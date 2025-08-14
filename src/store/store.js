// store.js
import { configureStore } from '@reduxjs/toolkit';
import excelReducer from './slice/excelSlice';
//import amcDataReducer from './slice/amcDataSlice';
import warrantyDataReducer from './slice/warrantyDataSlice';
import dashboardReducer from './slice/dashboardSlice';
import amcScheduleReducer from './slice/amcScheduleSlice';
//import warrantyReducer from './slice/warrantyDataSlice';

const store = configureStore({
  reducer: {
    excel: excelReducer,
   // amcData: amcDataReducer,
    warranty: warrantyDataReducer,
    dashboard: dashboardReducer,
    amcSchedule: amcScheduleReducer,
    //warranty: warrantyReducer
  },
  middleware: (getDefaultMiddleware) =>
    getDefaultMiddleware({
      serializableCheck: {
        ignoredActions: [
          'persist/PERSIST',
          'excel/processFile/fulfilled',
          'excel/processFile/pending',
          'excel/processFile/rejected'
        ],
        ignoredActionsPaths: ['meta.arg', 'payload.timestamp'],
        ignoredPaths: ['items.dates'],
      },
    }),
  devTools: process.env.NODE_ENV !== 'production',
});

export default store;
