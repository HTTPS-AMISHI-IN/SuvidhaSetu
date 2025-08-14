import { createSlice } from '@reduxjs/toolkit';

const warrantyDataSlice = createSlice({
  name: 'warrantyData',
  initialState: {
    calculations: [],
    metadata: {},
    history: [],
    payments: [],
    filters: {},
    isProcessing: false,
    hasData: false,
    totalAssets: 0,
    totalValue: 0,
    
  },
  reducers: {
    storeWarrantyCalculations: (state, action) => {
      state.calculations = action.payload.calculations;
      state.metadata = action.payload.metadata;
      state.hasData = true;
      state.totalAssets = action.payload.calculations.length;
      state.totalValue = action.payload.calculations.reduce((sum, calc) => sum + (calc.totalWarrantyValue || 0), 0);
      state.history.push({
        timestamp: new Date().toISOString(),
        calculationsCount: action.payload.calculations.length
      });

      // Save to localStorage
      localStorage.setItem('warrantyData', JSON.stringify({
        calculations: state.calculations,
        metadata: state.metadata,
        hasData: state.hasData,
        totalAssets: state.totalAssets,
        totalValue: state.totalValue,
        history: state.history,
        
      }));
    },

    setWarrantyCalculatedData: (state, action) => {
      state.calculations = action.payload;
      state.hasData = true;
      state.totalAssets = action.payload.length;
      state.totalValue = action.payload.reduce((sum, calc) => sum + (calc.totalWarrantyValue || 0), 0);
      
    },

    setWarrantyPaymentData: (state, action) => {
      state.payments = action.payload;
    },

    setFilters: (state, action) => {
      state.filters = action.payload;
    },

    updateWarrantyPaymentStatus: (state, action) => {
      const { id, status } = action.payload;
      const item = state.payments?.find(p => p.id === id);
      if (item) item.status = status;
    },

    loadWarrantyFromMemory: (state) => {
      const savedData = localStorage.getItem('warrantyData');
      if (savedData) {
        const parsedData = JSON.parse(savedData);
        Object.assign(state, parsedData);
      }
    },

    clearWarrantyData: (state) => {
      state.calculations = [];
      state.metadata = {};
      state.history = [];
      state.payments = [];
      state.filters = {};
      state.hasData = false;
      state.totalAssets = 0;
      state.totalValue = 0;
      
      localStorage.removeItem('warrantyData');
    },

    setProcessing: (state, action) => {
      state.isProcessing = action.payload;
    }
  }
});

export const {
  storeWarrantyCalculations,
  setWarrantyCalculatedData,
  setWarrantyPaymentData,
  setFilters,
  updateWarrantyPaymentStatus,
  loadWarrantyFromMemory,
  clearWarrantyData,
  setProcessing
} = warrantyDataSlice.actions;

export default warrantyDataSlice.reducer;
