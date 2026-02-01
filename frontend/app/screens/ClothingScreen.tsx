import React, { useState, useEffect } from 'react';
import { View, Text, StyleSheet, TouchableOpacity, Alert, ScrollView } from 'react-native';
import { SafeAreaView } from 'react-native-safe-area-context';
import { useRoute, useNavigation } from '@react-navigation/native';
import { useAppSelector, useAppDispatch } from '../../redux/hooks';
import { addToQueue, syncQueue } from '../../redux/slices/syncSlice';
import { updateGuest } from '../../redux/slices/guestSlice';

export default function ClothingScreen() {
    const route = useRoute<any>();
    const navigation = useNavigation<any>();
    const { guestId } = route.params;
    const dispatch = useAppDispatch();

    const guest = useAppSelector(state => state.guest.guests[guestId]);
    const [quantity, setQuantity] = useState(0);

    // Default to 0 if undefined. Code.gs sends 0 if empty.
    const budget = guest?.feltonBucks || 0;

    const handleIncrement = () => {
        if (quantity < budget) {
            setQuantity(q => q + 1);
        } else {
            Alert.alert('Limit Reached', 'Cannot select more than current budget.');
        }
    };

    const handleDecrement = () => {
        if (quantity > 0) setQuantity(q => q - 1);
    };

    const handleSubmit = () => {
        if (quantity === 0) return;

        const timestamp = new Date().toISOString();

        // 1. Queue for Backend (Atomic CLOTHING_PURCHASE)
        dispatch(addToQueue({
            action: 'CLOTHING_PURCHASE', // New action type needs added to types?
            payload: {
                guestId,
                quantity,
                timestamp
            }
        }));

        // 2. Optimistic Update Local State
        if (guest) {
            dispatch(updateGuest({
                ...guest,
                feltonBucks: budget - quantity,
                lastVisit: timestamp
            }));
        }

        // 3. Trigger Sync
        // @ts-ignore
        dispatch(syncQueue());

        Alert.alert('Success', `Purchased ${quantity} items.`, [
            { text: 'OK', onPress: () => navigation.goBack() }
        ]);
    };

    const [refreshing, setRefreshing] = useState(false);

    const refreshBudget = async () => {
        setRefreshing(true);
        console.log('[Clothing] Refreshing budget for:', guestId);
        try {
            // but here we want immediate read
            const response = await fetch('https://script.google.com/macros/s/AKfycbxEdsw5CquVDGOb2NDjxok0-zgQZVfV1Ej5AyP-n_6UmaqJpb-ap7OOothrwr_Rq_7U/exec', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    action: 'GET_BUDGET',
                    payload: { guestId }
                }),
            });
            const text = await response.text();
            console.log('[Clothing] Refresh Response:', text);
            const res = JSON.parse(text);

            if (res.status === 'success' && guest) {
                dispatch(updateGuest({
                    ...guest,
                    feltonBucks: res.budget
                }));
                Alert.alert('Refreshed', `Budget Updated: ${res.budget}`);
            } else {
                Alert.alert('Error', res.message || 'Failed to refresh');
            }
        } catch (e) {
            console.error(e);
            Alert.alert('Error', 'Network Error');
        } finally {
            setRefreshing(false);
        }
    }

    return (
        <SafeAreaView style={styles.container}>
            <View style={styles.content}>
                <Text style={styles.title}>Clothing Store</Text>
                <Text style={styles.subtitle}>Guest: {guest.name || guestId}</Text>

                <View style={styles.card}>
                    <Text style={styles.budgetLabel}>Current Budget</Text>
                    <Text style={styles.budgetValue}>{budget} Felton Bucks</Text>
                    <TouchableOpacity onPress={refreshBudget} disabled={refreshing}>
                        <Text style={{ color: '#1565c0', marginTop: 10, textDecorationLine: 'underline' }}>
                            {refreshing ? 'Refreshing...' : 'Refresh Budget'}
                        </Text>
                    </TouchableOpacity>
                </View>

                {budget === 0 ? (
                    <Text style={styles.noBudget}>No clothing budget available.</Text>
                ) : (
                    <View style={styles.controls}>
                        <Text style={styles.label}>Select Items:</Text>
                        <View style={styles.counter}>
                            <TouchableOpacity style={styles.btn} onPress={handleDecrement}>
                                <Text style={styles.btnText}>-</Text>
                            </TouchableOpacity>
                            <Text style={styles.qty}>{quantity}</Text>
                            <TouchableOpacity style={styles.btn} onPress={handleIncrement}>
                                <Text style={styles.btnText}>+</Text>
                            </TouchableOpacity>
                        </View>
                    </View>
                )}

                <TouchableOpacity
                    style={[styles.submitBtn, (quantity === 0 || budget === 0) && styles.disabled]}
                    onPress={handleSubmit}
                    disabled={quantity === 0 || budget === 0}
                >
                    <Text style={styles.submitText}>Confirm Purchase ({quantity})</Text>
                </TouchableOpacity>
            </View>
        </SafeAreaView>
    );
}

const styles = StyleSheet.create({
    container: { flex: 1, backgroundColor: '#f5f5f5' },
    content: { padding: 20, alignItems: 'center' },
    title: { fontSize: 24, fontWeight: 'bold', marginBottom: 10 },
    subtitle: { fontSize: 16, color: '#666', marginBottom: 20 },
    card: {
        backgroundColor: '#e3f2fd',
        padding: 20,
        borderRadius: 10,
        alignItems: 'center',
        width: '100%',
        marginBottom: 30
    },
    budgetLabel: { fontSize: 16, color: '#1565c0' },
    budgetValue: { fontSize: 32, fontWeight: 'bold', color: '#0d47a1' },
    noBudget: { color: 'red', fontSize: 16, marginTop: 20 },
    controls: { alignItems: 'center', marginBottom: 30 },
    label: { fontSize: 18, marginBottom: 10 },
    counter: { flexDirection: 'row', alignItems: 'center' },
    btn: { width: 50, height: 50, borderRadius: 25, backgroundColor: '#ddd', justifyContent: 'center', alignItems: 'center' },
    btnText: { fontSize: 24, fontWeight: 'bold' },
    qty: { fontSize: 32, fontWeight: 'bold', marginHorizontal: 20 },
    submitBtn: {
        backgroundColor: '#4caf50',
        paddingVertical: 15,
        paddingHorizontal: 40,
        borderRadius: 10,
        width: '100%',
        alignItems: 'center'
    },
    disabled: { backgroundColor: '#ccc' },
    submitText: { color: 'white', fontSize: 18, fontWeight: 'bold' },
    error: { color: 'red', fontSize: 18, textAlign: 'center', marginTop: 50 }
});
