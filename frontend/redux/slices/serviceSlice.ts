import { createSlice, PayloadAction } from '@reduxjs/toolkit';

interface ServiceState {
  shower: boolean;
  laundry: boolean;
  mealCount: number;
  hygieneKitCount: number;
  clothingCount: number;
}

const initialState: ServiceState = {
  shower: false,
  laundry: false,
  mealCount: 0,
  hygieneKitCount: 0,
  clothingCount: 0,
};

const serviceSlice = createSlice({
  name: 'service',
  initialState,
  reducers: {
    toggleShower: (state) => {
      state.shower = !state.shower;
    },
    toggleLaundry: (state) => {
      state.laundry = !state.laundry;
    },
    incrementMeal: (state) => {
      state.mealCount += 1;
    },
    decrementMeal: (state) => {
      if (state.mealCount > 0) state.mealCount -= 1;
    },
    incrementHygiene: (state) => {
      state.hygieneKitCount += 1;
    },
    decrementHygiene: (state) => {
      if (state.hygieneKitCount > 0) state.hygieneKitCount -= 1;
    },
    incrementClothing: (state) => {
      state.clothingCount += 1;
    },
    decrementClothing: (state) => {
      if (state.clothingCount > 0) state.clothingCount -= 1;
    },
    resetServices: (state) => {
      return initialState;
    },
  },
});

export const {
  toggleShower,
  toggleLaundry,
  incrementMeal,
  decrementMeal,
  incrementHygiene,
  decrementHygiene,
  incrementClothing,
  decrementClothing,
  resetServices,
} = serviceSlice.actions;

export default serviceSlice.reducer;
