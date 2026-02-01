import React, { useState } from 'react';
import { View, Text, StyleSheet, TouchableOpacity, TextInput, Switch, ScrollView } from 'react-native';
import { SafeAreaView } from 'react-native-safe-area-context';
import { useNavigation, useRoute } from '@react-navigation/native';
import { useAppDispatch } from '../../redux/hooks';
import { addGuest } from '../../redux/slices/guestSlice';
import { addToQueue, syncQueue } from '../../redux/slices/syncSlice';
import { Guest } from '../../types';

export default function NewUserSetupScreen() {
    const navigation = useNavigation<any>();
    const route = useRoute<any>();
    const { guestId } = route.params;
    const dispatch = useAppDispatch();

    const [name, setName] = useState('');

    const handleSave = () => {
        const newGuest: Guest = {
            id: guestId,
            name: name.trim() || undefined, // Optional
            programs: {
                healthcare: false,
                seasonalNight: false,
                sustainability: false,
            },
            createdAt: new Date().toISOString(),
            lastVisit: new Date().toISOString()
        };

        // 1. Update Redux (Local)
        dispatch(addGuest(newGuest));

        // 2. Queue for Sync (Backend)
        // correlates with code.gs 'UPDATE_GUEST' which handles inserts
        dispatch(addToQueue({
            action: 'UPDATE_GUEST',
            payload: newGuest
        }));

        // 3. Trigger Sync
        // @ts-ignore
        dispatch(syncQueue());

        // 4. Go to Service Entry
        navigation.replace('ServiceEntry', { guestId, isNew: true });
    };

    return (
        <SafeAreaView style={styles.container}>
            <ScrollView contentContainerStyle={styles.scrollContent}>
                <Text style={styles.title}>New Guest Registration</Text>
                <Text style={styles.subtitle}>NFC ID: {guestId}</Text>

                <View style={styles.formGroup}>
                    <Text style={styles.label}>Name (Optional)</Text>
                    <TextInput
                        style={styles.input}
                        placeholder="Enter guest name"
                        value={name}
                        onChangeText={setName}
                    />
                </View>

                {/* Programs Removed as per new requirement */}


                <TouchableOpacity style={styles.button} onPress={handleSave}>
                    <Text style={styles.btnText}>Save & Continue</Text>
                </TouchableOpacity>
            </ScrollView>
        </SafeAreaView>
    );
}

const styles = StyleSheet.create({
    container: {
        flex: 1,
        backgroundColor: '#fff',
    },
    scrollContent: {
        padding: 20,
    },
    title: {
        fontSize: 24,
        fontWeight: 'bold',
        marginBottom: 5,
    },
    subtitle: {
        fontSize: 14,
        color: '#666',
        marginBottom: 30,
    },
    formGroup: {
        marginBottom: 25,
    },
    label: {
        fontSize: 16,
        marginBottom: 8,
        fontWeight: '600',
    },
    input: {
        borderWidth: 1,
        borderColor: '#ddd',
        borderRadius: 8,
        padding: 12,
        fontSize: 16,
        backgroundColor: '#f9f9f9',
    },
    sectionHeader: {
        fontSize: 18,
        fontWeight: 'bold',
        marginTop: 10,
        marginBottom: 15,
        color: '#333',
    },
    switchRow: {
        flexDirection: 'row',
        justifyContent: 'space-between',
        alignItems: 'center',
        paddingVertical: 12,
        borderBottomWidth: 1,
        borderBottomColor: '#eee',
    },
    switchLabel: {
        fontSize: 16,
    },
    button: {
        backgroundColor: '#2196f3',
        padding: 16,
        borderRadius: 8,
        alignItems: 'center',
        marginTop: 40,
    },
    btnText: {
        color: 'white',
        fontSize: 18,
        fontWeight: 'bold',
    },
});
