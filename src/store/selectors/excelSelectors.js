// Fixed selectors for excelSelectors.js - to match your actual slice structure

// Your slice stores data in these properties:
// - data: {} (the sheets object)
// - sheetNames: [] (array of sheet names)  
// - activeSheet: '' (current sheet name)
// - status: 'idle' | 'loading' | 'succeeded' | 'failed'
// - error: null

// Core selectors that match your slice structure
export const selectExcelData = (state) => state.excel.data || {};
export const selectSheetNames = (state) => state.excel.sheetNames || [];
export const selectActiveSheet = (state) => state.excel.activeSheet || '';
export const selectExcelStatus = (state) => state.excel.status || 'idle';
export const selectError = (state) => state.excel.error;

// Derived selectors (computed from the core selectors)
export const selectIsLoading = (state) => selectExcelStatus(state) === 'loading';
export const selectHasData = (state) => {
  const data = selectExcelData(state);
  const sheetNames = selectSheetNames(state);
  return Object.keys(data).length > 0 && sheetNames.length > 0;
};

// This should be selectSheetNames, not selectSheets
export const selectSheets = (state) => {
  const data = selectExcelData(state);
  const sheetNames = selectSheetNames(state);
  
  return sheetNames.map(name => ({
    name,
    data: data[name] || [],
    headers: data[name] && data[name].length > 0 ? Object.keys(data[name][0]) : [],
    rowCount: data[name] ? data[name].length : 0
  }));
};

export const selectTotalRows = (state) => {
  const data = selectExcelData(state);
  return Object.values(data).reduce((total, sheetData) => {
    return total + (Array.isArray(sheetData) ? sheetData.length : 0);
  }, 0);
};

export const selectFileName = (state) => {
  // Access fileName from the slice - it does store fileName
  return state.excel.fileName || null;
};

export const selectUploadHistory = (state) => {
  // Your slice doesn't have upload history, return empty array
  return [];
};

export const selectSheetData = (state, sheetName) => {
  const data = selectExcelData(state);
  return data[sheetName] || [];
};