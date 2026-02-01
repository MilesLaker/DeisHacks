import React, { useState } from 'react';
import { View, Text, TextInput, FlatList, TouchableOpacity, StyleSheet, Alert } from 'react-native';
import { SafeAreaView } from 'react-native-safe-area-context';
import { useAppSelector, useAppDispatch } from '../../redux/hooks';
import { Guest } from '../../types';
import { replaceCard, incrementLostCardCounter, addGuest } from '../../redux/slices/guestSlice';
import { addToQueue } from '../../redux/slices/syncSlice';
import { useNavigation, useRoute } from '@react-navigation/native';

export default function SearchGuestScreen() {
    const [query, setQuery] = useState('');
    const guests = useAppSelector((state) => state.guest.guests);
    const dispatch = useAppDispatch();
    const navigation = useNavigation<any>();
    const route = useRoute<any>();

    const newCardId = route.params?.newCardId;

    const guestList = Object.values(guests);

    // Filter Logic matching PRD: ID, Name, Program Status, Last Visit
    const filteredGuests = guestList.filter((g) => {
        if (!query) return false;
        const lowerQ = query.toLowerCase();

        const hasProgramMatch =
            (g.programs.healthcare && "healthcare".includes(lowerQ)) ||
            (g.programs.seasonalNight && "seasonal".includes(lowerQ)) ||
            (g.programs.sustainability && "sustainability".includes(lowerQ));

        return (
            g.id.toLowerCase().includes(lowerQ) ||
            (g.name && g.name.toLowerCase().includes(lowerQ)) ||
            (g.lastVisit && g.lastVisit.includes(lowerQ)) ||
            hasProgramMatch
        );
    });

    const handleSelect = (guest: Guest) => {
        if (newCardId) {
            Alert.alert(
                'Confirm Replacement',
                `Link new card (${newCardId}) to guest "${guest.name || guest.id}"?`,
                [
                    { text: 'Cancel', style: 'cancel' },
                    {
                        text: 'Link & Continue',
                        onPress: () => {
                            // 1. Update Redux (Local)
                            dispatch(replaceCard({ oldId: guest.id, newId: newCardId }));

                            // 2. Queue for Sync (Backend)
                            dispatch(addToQueue({
                                action: 'REPLACE_CARD',
                                payload: { oldId: guest.id, newId: newCardId }
                            }));

                            // 3. Go to Service Entry for the NEW ID
                            navigation.replace('ServiceEntry', { guestId: newCardId, isNew: false });
                        },
                    },
                ]
            );
        } else {
            // View-only mode or manual lookup -> Go to Service Entry
            navigation.navigate('ServiceEntry', { guestId: guest.id, isNew: false });
        }
    };

    const handleNewGuestFallback = () => {
        if (!newCardId) return;

        // PRD: "Increment global lost-card counter AND create a new unique ID"
        // Since we are creating a New Guest WITH the NFC ID, we just proceed to Setup.
        // We also count this as a "Lost Card" event where no profile was found.
        dispatch(incrementLostCardCounter());

        // Navigate to New User Setup
        navigation.replace('NewUserSetup', { guestId: newCardId });
    };

    return (
        <SafeAreaView style={styles.container}>
            <Text style={styles.header}>Search Guests</Text>
            {newCardId && (
                <View style={styles.infoBox}>
                    <Text style={styles.infoText}>Replacing Card for NFC ID: {newCardId}</Text>
                    <Text style={styles.subInfo}>Search for the guest's existing profile below.</Text>
                </View>
            )}

            <TextInput
                style={styles.input}
                placeholder="Name, ID, Program, or Date..."
                value={query}
                onChangeText={setQuery}
                autoCapitalize="none"
            />

            <FlatList
                data={filteredGuests}
                keyExtractor={(item) => item.id}
                renderItem={({ item }) => (
                    <TouchableOpacity style={styles.item} onPress={() => handleSelect(item)}>
                        <View>
                            <Text style={styles.name}>{item.name || 'Unnamed Guest'}</Text>
                            <Text style={styles.details}>ID: {item.id}</Text>
                            <Text style={styles.details}>Last Visit: {item.lastVisit || 'Never'}</Text>
                        </View>
                        {/* Status Badges */}
                        <View style={styles.badges}>
                            {item.programs.healthcare && <Text style={styles.badge}>Health</Text>}
                            {item.programs.seasonalNight && <Text style={styles.badge}>Night</Text>}
                        </View>
                    </TouchableOpacity>
                )}
                ListEmptyComponent={
                    <View style={styles.emptyContainer}>
                        <Text style={styles.empty}>No matches found</Text>
                        {newCardId && (
                            <TouchableOpacity style={styles.fallbackButton} onPress={handleNewGuestFallback}>
                                <Text style={styles.fallbackText}>Not Found? Create New Guest</Text>
                            </TouchableOpacity>
                        )}
                    </View>
                }
            />
        </SafeAreaView>
    );
}

const styles = StyleSheet.create({
    container: {
        flex: 1,
        padding: 20,
        backgroundColor: '#fff',
    },
    header: {
        fontSize: 24,
        fontWeight: 'bold',
        marginBottom: 10,
    },
    infoBox: {
        backgroundColor: '#fff3cd',
        padding: 10,
        borderRadius: 8,
        marginBottom: 15,
    },
    infoText: {
        color: '#856404',
        fontWeight: 'bold',
    },
    subInfo: {
        color: '#856404',
        fontSize: 12,
    },
    input: {
        borderWidth: 1,
        borderColor: '#ccc',
        padding: 15,
        borderRadius: 8,
        marginBottom: 20,
        fontSize: 16,
    },
    item: {
        padding: 15,
        borderBottomWidth: 1,
        borderBottomColor: '#eee',
        flexDirection: 'row',
        justifyContent: 'space-between',
        alignItems: 'center',
    },
    name: {
        fontSize: 18,
        fontWeight: '600',
    },
    details: {
        color: '#888',
        fontSize: 12,
    },
    badges: {
        alignItems: 'flex-end',
    },
    badge: {
        backgroundColor: '#e2e3e5',
        paddingHorizontal: 8,
        paddingVertical: 2,
        borderRadius: 10,
        fontSize: 10,
        marginTop: 2,
        overflow: 'hidden',
    },
    emptyContainer: {
        alignItems: 'center',
        marginTop: 30,
    },
    empty: {
        color: '#888',
        marginBottom: 20,
    },
    fallbackButton: {
        backgroundColor: '#007bff',
        padding: 15,
        borderRadius: 8,
    },
    fallbackText: {
        color: '#fff',
        fontWeight: 'bold',
    },
});
