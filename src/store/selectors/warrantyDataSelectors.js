export const selectWarrantyCalculations = (state) => state.warrantyData?.calculations || [];
export const selectWarrantyMetadata = (state) => state.warrantyData?.metadata || {};
export const selectWarrantyHistory = (state) => state.warrantyData?.history || [];
export const selectWarrantyIsProcessing = (state) => state.warrantyData?.isProcessing || false;
export const selectHasWarrantyData = (state) => Boolean(state.warrantyData?.calculations?.length);
export const selectTotalWarrantyAssets = (state) => {
  const calculations = state.warrantyData?.calculations || [];
  return calculations.length;
};
export const selectTotalWarrantyValue = (state) => {
  const calculations = state.warrantyData?.calculations || [];
  return calculations.reduce((total, calc) => total + (calc.totalValue || 0), 0);
};
export const selectWarrantyCalculationsForPayments = (state) => {
  const calculations = state.warrantyData?.calculations || [];
  return calculations.map(calc => ({
    ...calc,
    type: 'warranty',
    source: 'Warranty Calculator'
  }));
};
