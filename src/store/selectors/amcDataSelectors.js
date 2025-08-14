import { createSelector } from '@reduxjs/toolkit';

export const selectAMCCalculations = (state) => state.amcSchedule?.calculations || [];
export const selectAMCMetadata = (state) => state.amcSchedule?.metadata || {};
export const selectAMCHistory = (state) => state.amcSchedule?.history || [];
export const selectAMCIsProcessing = (state) => state.amcSchedule?.isProcessing || false;
export const selectHasAMCData = (state) => Boolean(state.amcSchedule?.calculations?.length);
export const selectTotalAMCAssets = (state) => {
  const calculations = state.amcSchedule?.calculations || [];
  return calculations.length;
};
export const selectTotalAMCValue = (state) => {
  const calculations = state.amcSchedule?.calculations || [];
  return calculations.reduce((total, calc) => total + (calc.totalValue || 0), 0);
};
export const selectAMCCalculationsForPayments = (state) => {
  const calculations = state.amcSchedule?.calculations || [];
  return calculations.map(calc => ({
    ...calc,
    type: 'amc',
    source: 'AMC Calculator'
  }));
};