import React from 'react';
import { View, Text, StyleSheet, TouchableOpacity, ScrollView } from 'react-native';
import { SafeAreaView } from 'react-native-safe-area-context';
import { useAppDispatch, useAppSelector } from '../../redux/hooks';
import {
    toggleShower,
    toggleLaundry,
    incrementMeal,
    decrementMeal,
    incrementHygiene,
    decrementHygiene,
    incrementClothing,
    decrementClothing,
    resetServices,
} from '../../redux/slices/serviceSlice';
import { addToQueue, syncQueue } from '../../redux/slices/syncSlice';
import { updateGuest } from '../../redux/slices/guestSlice';
import { RouteProp, useNavigation, useRoute } from '@react-navigation/native';

type RootStackParamList = {
    ServiceEntry: { guestId: string; isNew: boolean; isAnonymous?: boolean };
};

type ServiceEntryRouteProp = RouteProp<RootStackParamList, 'ServiceEntry'>;

export default function ServiceEntryScreen() {
    const dispatch = useAppDispatch();
    const navigation = useNavigation();
    const route = useRoute<ServiceEntryRouteProp>();
    const { guestId, isNew, isAnonymous } = route.params;

    const services = useAppSelector((state) => state.service);

    // Select Guest to update it locally
    const guest = useAppSelector(state => state.guest.guests[guestId]);

    const onConfirm = () => {
        const timestamp = new Date().toISOString();

        if (isAnonymous) {
            dispatch(addToQueue({
                action: 'ANONYMOUS_ENTRY',
                payload: {
                    meals: services.mealCount,
                    timestamp
                }
            }));
        } else {
            // Create Canonical Payload
            const canonicalServices = {
                shower: services.shower,
                laundry: services.laundry,
                meals: services.mealCount,        // Map mealCount -> meals
                hygieneKits: services.hygieneKitCount, // Map hygieneKitCount -> hygieneKits
                // clothing handled in separate flow
            };

            console.log('[ServiceEntry] Submitting Payload:', { guestId, services: canonicalServices, timestamp });
            console.log('[ServiceEntry] Action: LOG_SERVICE');

            dispatch(addToQueue({
                action: 'LOG_SERVICE',
                payload: { guestId, services: canonicalServices, timestamp }
            }));

            if (guest) {
                dispatch(updateGuest({ ...guest, lastVisit: timestamp }));
            }
        }

        // Trigger sync
        // @ts-ignore
        dispatch(syncQueue());

        dispatch(resetServices());
        navigation.goBack();
    };

    return (
        <SafeAreaView style={styles.container}>
            <ScrollView contentContainerStyle={styles.scrollContent}>
                <Text style={styles.header}>Guest ID: {guestId}</Text>
                {isNew && <Text style={styles.newBadge}>New Guest</Text>}

                {!isAnonymous && (
                    <View style={styles.section}>
                        <Text style={styles.sectionTitle}>Facilities</Text>
                        <TouchableOpacity
                            style={[styles.toggleRow, services.shower && styles.activeRow]}
                            onPress={() => dispatch(toggleShower())}
                        >
                            <Text style={styles.label}>Shower</Text>
                            <Text style={styles.value}>{services.shower ? 'YES' : 'NO'}</Text>
                        </TouchableOpacity>

                        <TouchableOpacity
                            style={[styles.toggleRow, services.laundry && styles.activeRow]}
                            onPress={() => dispatch(toggleLaundry())}
                        >
                            <Text style={styles.label}>Laundry</Text>
                            <Text style={styles.value}>{services.laundry ? 'YES' : 'NO'}</Text>
                        </TouchableOpacity>
                    </View>
                )}

                <View style={styles.section}>
                    <Text style={styles.sectionTitle}>Distributions</Text>

                    <CounterRow
                        label="Meals"
                        count={services.mealCount}
                        onIncrement={() => dispatch(incrementMeal())}
                        onDecrement={() => dispatch(decrementMeal())}
                    />

                    {!isAnonymous && (
                        <CounterRow
                            label="Hygiene Kits"
                            count={services.hygieneKitCount}
                            onIncrement={() => dispatch(incrementHygiene())}
                            onDecrement={() => dispatch(decrementHygiene())}
                        />
                    )}

                    {/* Clothing moved to separate tab */}
                </View>

                <TouchableOpacity style={styles.submitButton} onPress={onConfirm}>
                    <Text style={styles.submitText}>Submit Entry</Text>
                </TouchableOpacity>

                {/* DEV DEBUG TEST */}
                {__DEV__ && (
                    <TouchableOpacity
                        style={[styles.submitButton, { backgroundColor: 'orange', marginTop: 10 }]}
                        onPress={() => {
                            const ts = new Date().toISOString();
                            console.log('[Debug] Sending Test Log Row...');
                            dispatch(addToQueue({
                                action: 'LOG_SERVICE',
                                payload: {
                                    guestId: guestId || 'debug_test',
                                    services: { shower: true, laundry: true, meals: 2, hygieneKits: 1 },
                                    timestamp: ts
                                }
                            }));
                            // @ts-ignore
                            dispatch(syncQueue());
                        }}
                    >
                        <Text style={styles.submitText}>TEST: Write Log Row</Text>
                    </TouchableOpacity>
                )}
            </ScrollView>
        </SafeAreaView>
    );
}

const CounterRow = ({ label, count, onIncrement, onDecrement }: any) => (
    <View style={styles.counterRow}>
        <Text style={styles.label}>{label}</Text>
        <View style={styles.controls}>
            <TouchableOpacity style={styles.btnInfo} onPress={onDecrement}>
                <Text style={styles.btnText}>-</Text>
            </TouchableOpacity>
            <Text style={styles.count}>{count}</Text>
            <TouchableOpacity style={styles.btnInfo} onPress={onIncrement}>
                <Text style={styles.btnText}>+</Text>
            </TouchableOpacity>
        </View>
    </View>
);

const styles = StyleSheet.create({
    container: {
        flex: 1,
        backgroundColor: '#f5f5f5',
    },
    scrollContent: {
        padding: 20,
    },
    header: {
        fontSize: 18,
        fontWeight: 'bold',
        marginBottom: 5,
    },
    newBadge: {
        color: 'green',
        fontWeight: 'bold',
        marginBottom: 20,
    },
    section: {
        backgroundColor: 'white',
        borderRadius: 10,
        padding: 15,
        marginBottom: 20,
        shadowColor: '#000',
        shadowOpacity: 0.05,
        shadowRadius: 5,
        elevation: 2,
    },
    sectionTitle: {
        fontSize: 16,
        fontWeight: '600',
        marginBottom: 15,
        color: '#444',
    },
    toggleRow: {
        flexDirection: 'row',
        justifyContent: 'space-between',
        paddingVertical: 15,
        borderBottomWidth: 1,
        borderBottomColor: '#eee',
    },
    activeRow: {
        backgroundColor: '#e3f2fd',
        paddingHorizontal: 10,
        borderRadius: 5,
    },
    counterRow: {
        flexDirection: 'row',
        justifyContent: 'space-between',
        alignItems: 'center',
        paddingVertical: 15,
        borderBottomWidth: 1,
        borderBottomColor: '#eee',
    },
    label: {
        fontSize: 16,
    },
    value: {
        fontSize: 16,
        fontWeight: 'bold',
        color: '#2196f3',
    },
    controls: {
        flexDirection: 'row',
        alignItems: 'center',
    },
    btnInfo: {
        width: 35,
        height: 35,
        borderRadius: 17.5,
        backgroundColor: '#eee',
        justifyContent: 'center',
        alignItems: 'center',
    },
    btnText: {
        fontSize: 20,
        fontWeight: 'bold',
    },
    count: {
        fontSize: 18,
        fontWeight: 'bold',
        marginHorizontal: 15,
        width: 30,
        textAlign: 'center',
    },
    submitButton: {
        backgroundColor: '#4caf50',
        padding: 15,
        borderRadius: 10,
        alignItems: 'center',
        marginBottom: 30,
    },
    submitText: {
        color: 'white',
        fontSize: 18,
        fontWeight: 'bold',
    },
});
