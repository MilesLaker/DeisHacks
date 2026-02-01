import React, { useState } from 'react';
import { View, Text, StyleSheet, TouchableOpacity, Alert, Linking, TextInput, ScrollView } from 'react-native';
import { SafeAreaView } from 'react-native-safe-area-context';
import { useNavigation } from '@react-navigation/native';
import { NativeStackNavigationProp } from '@react-navigation/native-stack';
import { useAppSelector, useAppDispatch } from '../../redux/hooks';
import { syncQueue, addToQueue } from '../../redux/slices/syncSlice';

// Define navigation types locally or in a shared file if complex
type RootStackParamList = {
    Home: undefined;
    ServiceEntry: { guestId: string; isNew: boolean; isAnonymous?: boolean };
    NewUserSetup: { guestId: string };
    SearchGuest: { newCardId?: string };
    Clothing: { guestId: string };
};

type GenericNavigationProp = NativeStackNavigationProp<RootStackParamList>;

export default function HomeScreen() {
    const navigation = useNavigation<GenericNavigationProp>();
    const guests = useAppSelector((state) => state.guest.guests);
    const syncState = useAppSelector((state) => state.sync);
    const dispatch = useAppDispatch();

    // Debug State
    const [debugId, setDebugId] = useState('guest_001');

    // Sync on mount
    React.useEffect(() => {
        // @ts-ignore
        dispatch(syncQueue());
    }, []);

    // Core Scan Logic
    const handleScan = (tagId: string) => {
        console.log(`[NFC] Scanned ID: ${tagId}`);
        const existingGuest = guests[tagId];

        if (existingGuest) {
            console.log(`[NFC] Guest Found: ${existingGuest.name || 'Unnamed'} (${tagId}). Prompting.`);

            Alert.alert(
                'Guest Found',
                `Name: ${existingGuest.name || 'Unknown'}\nSelect Action:`,
                [
                    {
                        text: 'Service Entry',
                        onPress: () => navigation.navigate('ServiceEntry', { guestId: tagId, isNew: false })
                    },
                    {
                        text: 'Clothing Store',
                        onPress: () => navigation.navigate('Clothing', { guestId: tagId })
                    },
                    { text: 'Cancel', style: 'cancel' }
                ]
            );
        } else {
            console.log(`[NFC] Guest NOT Found. Prompting New/Lost.`);
            // New -> Ask "New Guest" or "Replacement"?
            Alert.alert(
                'Unknown Tag',
                `ID: ${tagId}\nIs this a new guest or a replacement card?`,
                [
                    {
                        text: 'New Guest',
                        onPress: () => {
                            console.log(`[NFC] User selected 'New Guest' for ${tagId}`);
                            navigation.navigate('NewUserSetup', { guestId: tagId });
                        }
                    },
                    {
                        text: 'Replacement Card',
                        onPress: () => {
                            console.log(`[NFC] User selected 'Replacement Card' for ${tagId}`);
                            navigation.navigate('SearchGuest', { newCardId: tagId });
                        }
                    },
                    { text: 'Cancel', style: 'cancel' }
                ]
            );
        }
    };

    const handleSimulatedTap = () => {
        const mockId = 'guest_' + Math.floor(Math.random() * 1000);
        console.log(`[Debug] Simulating Random Tap: ${mockId}`);
        handleScan(mockId);
    };

    const handleManualDebugScan = () => {
        if (!debugId) return;
        console.log(`[Debug] Simulating Manual Tap: ${debugId}`);
        handleScan(debugId);
    };

    const handleSyncNow = () => {
        // @ts-ignore
        dispatch(syncQueue());
    };

    const handleTestEvent = () => {
        dispatch(addToQueue({
            action: 'LOG_SERVICE',
            payload: {
                guestId: 'debug_test_guest',
                services: { shower: true, laundry: false, mealCount: 1, hygieneKitCount: 0, clothingCount: 0 },
                timestamp: new Date().toISOString()
            }
        }));
        // Auto trigger sync
        setTimeout(() => handleSyncNow(), 500);
    };

    return (
        <SafeAreaView style={styles.container}>
            <View style={styles.content}>
                <Text style={styles.title}>Service Entry</Text>
                <View style={styles.circle}>
                    <Text style={styles.status}>Ready to Scan</Text>
                </View>
                <Text style={styles.instruction}>Tap NFC Card to begin</Text>

                {/* DEV ONLY DEBUG CONTROLS */}
                {__DEV__ && (
                    <ScrollView style={styles.debugContainer} contentContainerStyle={{ alignItems: 'center' }}>
                        <Text style={styles.debugLabel}>--- DEV DEBUG & SYNC ---</Text>

                        <View style={styles.debugRow}>
                            <TextInput
                                style={styles.debugInput}
                                value={debugId}
                                onChangeText={setDebugId}
                                placeholder="Enter Test ID"
                            />
                            <TouchableOpacity style={styles.debugBtnSmall} onPress={handleManualDebugScan}>
                                <Text style={styles.debugBtnText}>Simulate</Text>
                            </TouchableOpacity>
                        </View>

                        {/* Sync Diagnostics */}
                        <View style={styles.syncBox}>
                            <Text style={styles.syncText}>Queue: {syncState.queue.length}</Text>
                            <Text style={styles.syncText}>Syncing: {syncState.isSyncing ? 'YES' : 'NO'}</Text>
                            {syncState.queue.length > 0 && (
                                <Text style={styles.errorText}>
                                    Last Err: {syncState.queue[0].lastError?.substring(0, 20) || 'None'}
                                </Text>
                            )}

                            <View style={styles.syncControls}>
                                <TouchableOpacity style={styles.syncBtn} onPress={handleSyncNow}>
                                    <Text style={styles.debugBtnText}>Sync Now</Text>
                                </TouchableOpacity>
                                <TouchableOpacity style={[styles.syncBtn, { marginLeft: 10, backgroundColor: 'orange' }]} onPress={handleTestEvent}>
                                    <Text style={styles.debugBtnText}>Test Event</Text>
                                </TouchableOpacity>
                            </View>
                        </View>

                        <TouchableOpacity style={styles.debugButton} onPress={handleSimulatedTap}>
                            <Text style={styles.debugText}>Simulate Random Tap</Text>
                        </TouchableOpacity>

                        <TouchableOpacity
                            style={[styles.debugButton, { marginTop: 10, backgroundColor: '#ddd' }]}
                            onPress={() => navigation.navigate('SearchGuest', {})}
                        >
                            <Text style={styles.debugText}>Staff Search (Lost Card)</Text>
                        </TouchableOpacity>
                    </ScrollView>
                )}

                <View style={{ width: '100%', alignItems: 'center', marginTop: 20 }}>
                    <TouchableOpacity
                        style={[styles.debugButton, { backgroundColor: '#424242', marginBottom: 20 }]}
                        onPress={() => navigation.navigate('ServiceEntry', { guestId: 'anonymous', isNew: false, isAnonymous: true })}
                    >
                        <Text style={[styles.debugText, { color: 'white', fontWeight: 'bold' }]}>Anonymous Entry</Text>
                    </TouchableOpacity>

                    {!__DEV__ && (
                        <TouchableOpacity
                            style={styles.linkButton}
                            onPress={() => Linking.openURL('https://docs.google.com/spreadsheets')}
                        >
                            <Text style={styles.linkText}>View Google Sheet Data</Text>
                        </TouchableOpacity>
                    )}
                </View>
            </View>
        </SafeAreaView>
    );
}

const styles = StyleSheet.create({
    container: {
        flex: 1,
        backgroundColor: '#fff',
    },
    content: {
        flex: 1,
        justifyContent: 'center',
        alignItems: 'center',
        padding: 20,
    },
    title: {
        fontSize: 24,
        fontWeight: 'bold',
        marginBottom: 20,
        color: '#333',
    },
    circle: {
        width: 150,
        height: 150,
        borderRadius: 75,
        backgroundColor: '#e3f2fd',
        justifyContent: 'center',
        alignItems: 'center',
        marginBottom: 20,
        borderWidth: 2,
        borderColor: '#2196f3',
    },
    status: {
        fontSize: 18,
        color: '#1565c0',
        fontWeight: '600',
    },
    instruction: {
        fontSize: 16,
        color: '#666',
        marginBottom: 10,
    },
    debugContainer: {
        marginTop: 10,
        width: '100%',
        padding: 10,
        backgroundColor: '#f8f9fa',
        borderWidth: 1,
        borderColor: '#dee2e6',
        borderRadius: 8,
        maxHeight: 400,
    },
    debugLabel: {
        fontSize: 12,
        color: '#999',
        marginBottom: 10,
        fontWeight: 'bold',
    },
    debugRow: {
        flexDirection: 'row',
        width: '100%',
        marginBottom: 10,
    },
    debugInput: {
        flex: 1,
        height: 35,
        borderColor: '#ccc',
        borderWidth: 1,
        borderRadius: 4,
        paddingHorizontal: 10,
        marginRight: 10,
        backgroundColor: 'white',
    },
    debugBtnSmall: {
        backgroundColor: '#6c757d',
        paddingHorizontal: 15,
        justifyContent: 'center',
        borderRadius: 4,
    },
    debugBtnText: {
        color: 'white',
        fontWeight: 'bold',
        fontSize: 12,
    },
    debugButton: {
        width: '100%',
        padding: 10,
        backgroundColor: '#eee',
        borderRadius: 8,
        alignItems: 'center',
        marginBottom: 5,
    },
    debugText: {
        color: '#333',
        fontSize: 14,
    },
    linkButton: {
        marginTop: 20,
    },
    linkText: {
        color: 'blue',
        textDecorationLine: 'underline',
    },
    syncBox: {
        width: '100%',
        backgroundColor: '#e9ecef',
        padding: 10,
        borderRadius: 5,
        marginBottom: 10,
    },
    syncText: {
        fontSize: 12,
        marginBottom: 2,
    },
    errorText: {
        fontSize: 10,
        color: 'red',
        marginBottom: 5,
    },
    syncControls: {
        flexDirection: 'row',
        marginTop: 5,
        justifyContent: 'center'
    },
    syncBtn: {
        backgroundColor: '#28a745',
        padding: 8,
        borderRadius: 4,
    }
});
