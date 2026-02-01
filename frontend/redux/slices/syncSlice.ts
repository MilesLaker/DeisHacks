import { createSlice, PayloadAction, createAsyncThunk } from '@reduxjs/toolkit';
import { RootState } from '../store';

// We'll replace this with real URL later
const API_URL = 'https://script.google.com/macros/s/AKfycbxEdsw5CquVDGOb2NDjxok0-zgQZVfV1Ej5AyP-n_6UmaqJpb-ap7OOothrwr_Rq_7U/exec';

interface SyncItem {
    id: string; // Internal Queue ID
    eventId: string; // UUID for Backend Deduplication
    action: 'LOG_SERVICE' | 'UPDATE_GUEST' | 'REPLACE_CARD' | 'CLOTHING_PURCHASE' | 'ANONYMOUS_ENTRY';
    payload: any;
    timestamp: number;
    retryCount: number;
    lastError?: string;
}

interface SyncState {
    queue: SyncItem[];
    isSyncing: boolean;
    lastSyncTime: number | null;
}

const initialState: SyncState = {
    queue: [],
    isSyncing: false,
    lastSyncTime: null,
};

export const syncQueue = createAsyncThunk(
    'sync/process',
    async (_, { getState, dispatch }) => {
        const state = getState() as RootState;
        const queue = state.sync.queue;

        if (queue.length === 0) return;

        // Process items one by one strictly to maintain order
        for (const item of queue) {
            try {
                // Exponential Backoff Check (simple version: wait 2^retry seconds)
                // In real app, we'd skip item in this loop or wait. 
                // For simplicity, we just try.

                const response = await fetch(API_URL, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        action: item.action,
                        payload: { ...item.payload, eventId: item.eventId },
                    }),
                });

                const resText = await response.text();
                let resJson;
                try {
                    resJson = JSON.parse(resText);
                } catch (e) {
                    // Non-JSON response (Google Error HTML?). Treat as failure.
                    throw new Error('Invalid JSON response: ' + resText.substring(0, 100));
                }

                if (response.ok && resJson.status === 'success') {
                    // Critical Verification for Log Service
                    if (item.action === 'LOG_SERVICE') {
                        if (resJson.log?.written !== 1) {
                            throw new Error(`Log append failed. Server said: ${JSON.stringify(resJson)}`);
                        }
                    } else if (item.action === 'CLOTHING_PURCHASE') {
                        if (resJson.log?.written !== 1) {
                            throw new Error(`Clothing Log append failed. Server said: ${JSON.stringify(resJson)}`);
                        }
                    }

                    console.log(`[Sync] Success: ${item.eventId} (Build: ${resJson.buildId})`);
                    dispatch(removeFromQueue(item.id));
                } else {
                    console.warn(`[Sync] Backend Error for ${item.eventId} (Build: ${resJson.buildId || '??'}):`, resJson);
                    // Update retry count (logic needed in reducer, but for now we basically loop)
                    // We should probably break to avoid hammering if backend is down
                    throw new Error(resJson.message || 'Unknown Backend Error');
                }
            } catch (e: any) {
                console.error(`[Sync] Network/Logic Error for ${item.eventId}:`, e.message);
                dispatch(updateSyncStatus({
                    id: item.id,
                    error: e.message
                }));
                break; // Stop syncing queue on first error to maintain order
            }
        }
    }
);

const syncSlice = createSlice({
    name: 'sync',
    initialState,
    reducers: {
        addToQueue: (state, action: PayloadAction<{ action: SyncItem['action']; payload: any }>) => {
            // Generate a UUID-like string for idempotency
            const eventId = 'evt_' + Date.now().toString(36) + Math.random().toString(36).substr(2, 9);

            state.queue.push({
                id: eventId, // Use same ID for queue key
                eventId: eventId,
                action: action.payload.action,
                payload: action.payload.payload,
                timestamp: Date.now(),
                retryCount: 0,
            });
        },
        removeFromQueue: (state, action: PayloadAction<string>) => {
            state.queue = state.queue.filter(i => i.id !== action.payload);
        },
        updateSyncStatus: (state, action: PayloadAction<{ id: string; error: string }>) => {
            const item = state.queue.find(i => i.id === action.payload.id);
            if (item) {
                item.retryCount += 1;
                item.lastError = action.payload.error;
            }
        }
    },
    extraReducers: (builder) => {
        builder
            .addCase(syncQueue.pending, (state) => {
                state.isSyncing = true;
            })
            .addCase(syncQueue.fulfilled, (state) => {
                state.isSyncing = false;
                state.lastSyncTime = Date.now();
            })
            .addCase(syncQueue.rejected, (state) => {
                state.isSyncing = false;
            });
    },
});

export const { addToQueue, removeFromQueue, updateSyncStatus } = syncSlice.actions;
export default syncSlice.reducer;
