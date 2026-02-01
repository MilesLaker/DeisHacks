import React from 'react';
import { NavigationContainer } from '@react-navigation/native';
import { createNativeStackNavigator } from '@react-navigation/native-stack';

import HomeScreen from './screens/HomeScreen';
import ServiceEntryScreen from './screens/ServiceEntryScreen';
import NewUserSetupScreen from './screens/NewUserSetupScreen';
import SearchGuestScreen from './screens/SearchGuestScreen';
import ClothingScreen from './screens/ClothingScreen';

const Stack = createNativeStackNavigator();

export default function Navigation() {
    return (
        <NavigationContainer>
            <Stack.Navigator initialRouteName="Home">
                <Stack.Screen
                    name="Home"
                    component={HomeScreen}
                    options={{ headerShown: false }}
                />
                <Stack.Screen
                    name="ServiceEntry"
                    component={ServiceEntryScreen}
                    options={{ title: 'Service Entry' }}
                />
                <Stack.Screen
                    name="NewUserSetup"
                    component={NewUserSetupScreen}
                    options={{ title: 'New Visitor Registration' }}
                />
                <Stack.Screen
                    name="SearchGuest"
                    component={SearchGuestScreen}
                    options={{ title: 'Find Guest' }}
                />
                <Stack.Screen
                    name="Clothing"
                    component={ClothingScreen}
                    options={{ title: 'Clothing Store' }}
                />
            </Stack.Navigator>
        </NavigationContainer>
    );
}
