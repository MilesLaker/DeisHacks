import { createSlice, PayloadAction } from '@reduxjs/toolkit';
import { Guest } from '../../types';

interface GuestState {
    guests: Record<string, Guest>; // Keyed by NFC ID
    totalUniqueGuests: number;
    totalLostCards: number;
}

const initialState: GuestState = {
    guests: {},
    totalUniqueGuests: 0,
    totalLostCards: 0,
};

const guestSlice = createSlice({
    name: 'guest',
    initialState,
    reducers: {
        addGuest: (state, action: PayloadAction<Guest>) => {
            const newGuest = action.payload;
            if (!state.guests[newGuest.id]) {
                state.guests[newGuest.id] = newGuest;
                state.totalUniqueGuests += 1;
            }
        },
        updateGuest: (state, action: PayloadAction<Guest>) => {
            // Overwrite existing guest data
            const guest = action.payload;
            if (state.guests[guest.id]) {
                state.guests[guest.id] = guest;
            }
        },
        replaceCard: (state, action: PayloadAction<{ oldId: string; newId: string }>) => {
            const { oldId, newId } = action.payload;
            const existingGuest = state.guests[oldId];
            if (existingGuest) {
                // PRD: "append the new userID... BOTH old and new IDs can resolve"
                // Implementation: specific new ID points to same guest data.
                state.guests[newId] = { ...existingGuest, id: newId };
                // We do NOT delete the old ID.

                // Track lost card (globally)
                state.totalLostCards += 1;
            }
        },
        incrementLostCardCounter: (state) => {
            // Called when a card is lost but NO profile is found (new guest eventually)
            state.totalLostCards += 1;
        }
    },
});

export const { addGuest, updateGuest, replaceCard, incrementLostCardCounter } = guestSlice.actions;
export default guestSlice.reducer;
