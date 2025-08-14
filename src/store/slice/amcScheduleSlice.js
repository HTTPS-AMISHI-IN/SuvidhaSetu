import { createSlice } from '@reduxjs/toolkit';

const initialState = {
  uploadedFile: null,
  calculatedData: null,
  loading: false,
  error: null,
  settings: {
    gstRate: 0.18,
    amcPercentage: 0.40,
    roiRates: [20, 22.5, 27.5, 30],
  },
};

const amcScheduleSlice = createSlice({
  name: 'amcSchedule',
  initialState,
  reducers: {
    setUploadedFile: (state, action) => {
      state.uploadedFile = action.payload;
      state.error = null;
    },
    setCalculatedData: (state, action) => {
      state.calculatedData = action.payload;
    },
    updateSettings: (state, action) => {
      state.settings = { ...state.settings, ...action.payload };
    },
    setLoading: (state, action) => {
      state.loading = action.payload;
    },
    setError: (state, action) => {
      state.error = action.payload;
    },
    clearData: (state) => {
      state.uploadedFile = null;
      state.calculatedData = null;
      state.error = null;
    },
    updateAmcPaymentStatus: (state, action) => {
      const { productId, quarter, year, isPaid, paidDate } = action.payload;

      if (!state.calculatedData) return;

      const product = state.calculatedData.products.find(p => p.id === productId);
      if (!product || !product.quarterlyData) return;

      const quarterEntry = product.quarterlyData.find(
        q => q.quarter === quarter && q.year === year
      );

      if (quarterEntry) {
        quarterEntry.isPaid = isPaid;
        quarterEntry.paidDate = paidDate;
      }
    },
    loadAMCFromMemory: (state, action) => {
      const stored = JSON.parse(localStorage.getItem('amcData'));
      if (stored) {
        state.calculatedData = stored.calculatedData;
        state.uploadedFile = stored.uploadedFile;
        state.settings = stored.settings;
      }
    }
  }
});

export const {
  setUploadedFile,
  setCalculatedData,
  updateSettings,
  setLoading,
  setError,
  clearData,
  loadAMCFromMemory,
  updateAmcPaymentStatus,
} = amcScheduleSlice.actions;

export default amcScheduleSlice.reducer;