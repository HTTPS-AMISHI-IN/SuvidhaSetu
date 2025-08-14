import { createSlice, createAsyncThunk } from '@reduxjs/toolkit';
import * as XLSX from 'xlsx';

// Async thunk to process Excel file
export const processExcelFile = createAsyncThunk(
  'excel/processExcelFile',
  async (file, { rejectWithValue }) => {
    try {
      console.log('🚀 Starting Excel file processing...', {
        fileName: file.name,
        fileSize: file.size,
        fileType: file.type
      });

      if (!file.name.match(/\.(xlsx|xls)$/)) {
        console.error('❌ Invalid file format:', file.name);
        return rejectWithValue('Unsupported file format. Please upload a .xlsx or .xls file.');
      }

      console.log('✅ File format validation passed');
      console.log('📖 Reading file as array buffer...');

      const arrayBuffer = await file.arrayBuffer();
      console.log('📖 File read successfully, parsing Excel...');

      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetNames = workbook.SheetNames;

      console.log('📊 Excel workbook parsed successfully:', {
        sheetNames,
        sheetsCount: sheetNames.length
      });

      const sheets = {};
      let totalRows = 0;

      sheetNames.forEach(sheetName => {
        console.log(`📋 Processing sheet: ${sheetName}`);
        
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {
          defval: '',
          raw: false,
          blankrows: false,
        });
        
        console.log(`📊 Sheet "${sheetName}" processed: ${jsonData.length} rows`);
        sheets[sheetName] = jsonData;
        totalRows += jsonData.length;
      });

      console.log('📊 All sheets processed successfully:', {
        totalSheets: sheetNames.length,
        totalRows
      });

      // Prepare data to save
      const dataToSave = {
        sheets,
        sheetNames,
        fileName: file.name,
        fileSize: file.size,
        uploadDate: new Date().toISOString(),
        totalRows
      };

      // Save to localStorage
      console.log('💾 Saving data to localStorage...');
      localStorage.setItem('excelData', JSON.stringify(sheets));
      localStorage.setItem('excelSheetNames', JSON.stringify(sheetNames));
      localStorage.setItem('excelFileName', file.name);
      localStorage.setItem('excelFileSize', file.size.toString());
      localStorage.setItem('excelUploadDate', dataToSave.uploadDate);
      
      console.log('✅ Data saved to localStorage successfully');
      console.log('🎉 FILE UPLOADED SUCCESSFULLY! 🎉');
      console.log('📈 Excel data processed and ready to use');

      return { 
        sheets, 
        sheetNames, 
        fileName: file.name,
        fileSize: file.size,
        uploadDate: dataToSave.uploadDate,
        totalRows
      };
    } catch (error) {
      console.error('❌ Error reading Excel file:', error);
      return rejectWithValue('Failed to process Excel file. Please try again.');
    }
  }
);

// Load saved data from localStorage on app init
export const loadExcelFromMemory = createAsyncThunk(
  'excel/loadExcelFromMemory',
  async (_, { rejectWithValue }) => {
    try {
      console.log('📥 Loading Excel data from localStorage...');
      
      const savedData = localStorage.getItem('excelData');
      const savedSheetNames = localStorage.getItem('excelSheetNames');
      const savedFileName = localStorage.getItem('excelFileName');
      const savedFileSize = localStorage.getItem('excelFileSize');
      const savedUploadDate = localStorage.getItem('excelUploadDate');

      if (savedData && savedSheetNames) {
        const sheets = JSON.parse(savedData);
        const sheetNames = JSON.parse(savedSheetNames);
        
        console.log('✅ Excel data loaded from localStorage:', {
          fileName: savedFileName,
          sheetsCount: sheetNames.length,
          uploadDate: savedUploadDate
        });

        return {
          sheets,
          sheetNames,
          fileName: savedFileName || 'Excel File',
          fileSize: savedFileSize ? parseInt(savedFileSize) : 0,
          uploadDate: savedUploadDate || new Date().toISOString()
        };
      } else {
        console.log('ℹ️ No Excel data found in localStorage');
        return rejectWithValue('No stored Excel data found.');
      }
    } catch (error) {
      console.error('❌ Failed to load Excel data from memory:', error);
      return rejectWithValue('Error loading Excel data from memory.');
    }
  }
);

const excelSlice = createSlice({
  name: 'excel',
  initialState: {
    data: {},              // All parsed sheets
    sheetNames: [],        // Sheet name list
    activeSheet: '',       // Currently selected sheet
    status: 'idle',        // idle | loading | succeeded | failed
    error: null,
    fileName: null,        // Added fileName
    fileSize: 0,           // Added fileSize
    uploadDate: null,      // Added uploadDate
    totalRows: 0           // Added totalRows
  },
  reducers: {
    clearData: (state) => {
      console.log('🗑️ Clearing Excel data...');
      
      state.data = {};
      state.sheetNames = [];
      state.activeSheet = '';
      state.status = 'idle';
      state.error = null;
      state.fileName = null;
      state.fileSize = 0;
      state.uploadDate = null;
      state.totalRows = 0;
      
      // Clear localStorage
      localStorage.removeItem('excelData');
      localStorage.removeItem('excelSheetNames');
      localStorage.removeItem('excelFileName');
      localStorage.removeItem('excelFileSize');
      localStorage.removeItem('excelUploadDate');
      
      console.log('✅ Excel data cleared successfully');
    },
    setActiveSheet: (state, action) => {
      console.log('📋 Setting active sheet:', action.payload);
      state.activeSheet = action.payload;
    },
  },
  extraReducers: (builder) => {
    builder
      .addCase(processExcelFile.pending, (state) => {
        console.log('⏳ Excel file processing started...');
        state.status = 'loading';
        state.error = null;
      })
      .addCase(processExcelFile.fulfilled, (state, action) => {
        console.log('🎉 Excel file processing completed successfully!');
        console.log('💾 Updating Redux store with Excel data...');
        
        state.status = 'succeeded';
        state.data = action.payload.sheets;
        state.sheetNames = action.payload.sheetNames;
        state.activeSheet = action.payload.sheetNames[0] || '';
        state.fileName = action.payload.fileName;
        state.fileSize = action.payload.fileSize;
        state.uploadDate = action.payload.uploadDate;
        state.totalRows = action.payload.totalRows || 0;
        
        console.log('✅ Redux store updated successfully:', {
          fileName: state.fileName,
          sheetsCount: state.sheetNames.length,
          totalRows: state.totalRows
        });
      })
      .addCase(processExcelFile.rejected, (state, action) => {
        console.error('❌ Excel file processing failed:', action.payload);
        state.status = 'failed';
        state.error = action.payload || 'Failed to process Excel file.';
      })
      .addCase(loadExcelFromMemory.fulfilled, (state, action) => {
        console.log('✅ Excel data loaded from memory successfully');
        
        state.status = 'succeeded';
        state.data = action.payload.sheets;
        state.sheetNames = action.payload.sheetNames;
        state.activeSheet = action.payload.sheetNames[0] || '';
        state.fileName = action.payload.fileName;
        state.fileSize = action.payload.fileSize;
        state.uploadDate = action.payload.uploadDate;
        
        // Calculate total rows
        state.totalRows = Object.values(action.payload.sheets).reduce((total, sheetData) => {
          return total + (Array.isArray(sheetData) ? sheetData.length : 0);
        }, 0);
        
        console.log('📊 Memory data loaded:', {
          fileName: state.fileName,
          sheetsCount: state.sheetNames.length,
          totalRows: state.totalRows
        });
      })
      .addCase(loadExcelFromMemory.rejected, (state, action) => {
        console.log('ℹ️ No Excel data to load from memory');
        state.status = 'idle';
        state.error = action.payload || null;
      });
  }
});

export const { clearData, setActiveSheet } = excelSlice.actions;

export default excelSlice.reducer;