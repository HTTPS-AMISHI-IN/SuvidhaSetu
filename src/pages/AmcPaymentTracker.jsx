import React, { useState, useEffect, useCallback, useMemo } from "react";
import { useSelector} from "react-redux";
import { useNavigate } from "react-router-dom";
import VirtualDataTable from "../components/VirtualDataTable";
import CalculationProgress from "../components/CalculationProgress";
import CacheManager from "../components/CacheManager";
import { useAMCCalculationWorker } from "../hooks/useWebWorker";
import { useAMCCache } from "../hooks/useCalculationCache";
import { Upload, Plus, FileText, ArrowLeft, Calculator, } from "lucide-react";
import { AMCExportManager } from '../utils/exportUtils';
import { Download, FileSpreadsheet, Database } from 'lucide-react';

// Import Redux
import {
  selectExcelData,
  selectHasData,
  selectFileName,
} from "../store/selectors/excelSelectors";

// Manual Product Form Component
const ManualProductForm = ({ onAddProduct }) => {
  const [formData, setFormData] = useState({
    productName: "",
    invoiceValue: "",
    invoiceNumber: "",
    quantity: 1,
    location: "",
    uatDate: new Date().toISOString().split("T")[0],
  });

  const handleSubmit = () => {
    const { productName, invoiceValue, invoiceNumber, quantity, location, uatDate } = formData;

    if (
      !productName.trim() ||
      !invoiceValue ||
      parseFloat(invoiceValue) <= 0 ||
      !invoiceNumber.trim() || '',
      invoiceNumber.length < 5 ||
      invoiceNumber.length > 20 ||
      !location.trim() ||
      !uatDate
    ) {
      alert("Please fill in all required fields with valid values.");
      return;
    }

    const manualProduct = {
      id: `manual_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
      productName: productName.trim(),
      invoiceValue: parseFloat(invoiceValue),
      invoiceNumber: invoiceNumber.trim() || '',
      quantity: parseInt(quantity) || 1,
      location: location.trim(),
      uatDate: uatDate,
      roi: 0, 
    };

    onAddProduct(manualProduct);

    // Reset form
    setFormData({
      productName: "",
      invoiceValue: "",
      invoiceNumber: "",
      quantity: 1,
      location: "",
      uatDate: new Date().toISOString().split("T")[0],
    });
  };

  return (
    <>
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))",
          gap: "16px",
          marginBottom: "24px",
        }}
      >
        <div>
          <label
            style={{
              display: "block",
              marginBottom: "8px",
              fontWeight: 600,
              color: "#374151",
            }}
          >
            Product Name *
          </label>
          <input
            type="text"
            placeholder="Enter product name"
            value={formData.productName}
            onChange={(e) =>
              setFormData((prev) => ({ ...prev, productName: e.target.value }))
            }
            style={{
              width: "100%",
              padding: "8px 12px",
              border: "1px solid #d1d5db",
              borderRadius: "6px",
              fontSize: "0.9rem",
            }}
          />
        </div>

        <div>
          <label
            style={{
              display: "block",
              marginBottom: "8px",
              fontWeight: 600,
              color: "#374151",
            }}
          >
            Invoice Value (₹) *
          </label>
          <input
            type="number"
            placeholder="0"
            min="0"
            step="0.01"
            value={formData.invoiceValue}
            onChange={(e) =>
              setFormData((prev) => ({ ...prev, invoiceValue: e.target.value }))
            }
            style={{
              width: "100%",
              padding: "8px 12px",
              border: "1px solid #d1d5db",
              borderRadius: "6px",
              fontSize: "0.9rem",
            }}
          />
        </div>

        <div>
          <label
            style={{
              display: "block",
              marginBottom: "8px",
              fontWeight: 600,
              color: "#374151",
            }}
          >
            Invoice Number *
          </label>
          <input
            type="text"
            placeholder="Enter invoice number"
            value={formData.invoiceNumber}
            onChange={(e) =>
              setFormData((prev) => ({ ...prev, invoiceNumber: e.target.value.replace(/\D/g, '') })) // Allow only digits
            }
            maxLength={20}
            style={{
              width: "100%",
              padding: "8px 12px",
              border: "1px solid #d1d5db",
              borderRadius: "6px",
              fontSize: "0.9rem",
            }}
          />
        </div>

        <div>
          <label
            style={{
              display: "block",
              marginBottom: "8px",
              fontWeight: 600,
              color: "#374151",
            }}
          >
            Quantity
          </label>
          <input
            type="number"
            placeholder="1"
            min="1"
            value={formData.quantity}
            onChange={(e) =>
              setFormData((prev) => ({ ...prev, quantity: e.target.value }))
            }
            style={{
              width: "100%",
              padding: "8px 12px",
              border: "1px solid #d1d5db",
              borderRadius: "6px",
              fontSize: "0.9rem",
            }}
          />
        </div>

        <div>
          <label
            style={{
              display: "block",
              marginBottom: "8px",
              fontWeight: 600,
              color: "#374151",
            }}
          >
            Location *
          </label>
          <input
            type="text"
            placeholder="Enter location"
            value={formData.location}
            onChange={(e) =>
              setFormData((prev) => ({ ...prev, location: e.target.value }))
            }
            style={{
              width: "100%",
              padding: "8px 12px",
              border: "1px solid #d1d5db",
              borderRadius: "6px",
              fontSize: "0.9rem",
            }}
          />
        </div>

        <div>
          <label
            style={{
              display: "block",
              marginBottom: "8px",
              fontWeight: 600,
              color: "#374151",
            }}
          >
            UAT Date *
          </label>
          <input
            type="date"
            value={formData.uatDate}
            onChange={(e) =>
              setFormData((prev) => ({ ...prev, uatDate: e.target.value }))
            }
            style={{
              width: "100%",
              padding: "8px 12px",
              border: "1px solid #d1d5db",
              borderRadius: "6px",
              fontSize: "0.9rem",
            }}
          />
        </div>
      </div>

      <div style={{ textAlign: "center" }}>
        <button
          onClick={handleSubmit}
          style={{
            padding: "12px 24px",
            backgroundColor: "#3b82f6",
            color: "white",
            border: "none",
            borderRadius: "8px",
            fontSize: "0.9rem",
            fontWeight: 600,
            cursor: "pointer",
            transition: "all 0.2s ease",
          }}
        >
          Add Product
        </button>
      </div>
    </>
  );
};

const AmcPaymentTracker = () => {
 
  // Redux selectors
  const excelData = useSelector(selectExcelData);
  const hasExcelData = useSelector(selectHasData);
  const fileName = useSelector(selectFileName);
  const navigate = useNavigate();
  const [activeMainTab, setActiveMainTab] = useState("processExcel"); // "processExcel" or "manualEntry"
  const [processDataView, setProcessDataView] = useState("cards"); // "table" or "cards"
  const [paidQuarters, setPaidQuarters] = useState({});
  const [selectedYear, setSelectedYear] = useState(null);
  const [selectedQuarter, setSelectedQuarter] = useState(null);

// Function to update payment status
const updatePaidQuarters = useCallback((quarterKey, isPaid, date = "") => {
  setPaidQuarters(prev => ({
    ...prev,
    [quarterKey]: { paid: isPaid, date: date }
  }));
}, []);

  // Local state
  const [settings, setSettings] = useState({
    roiRates: [20, 22.5, 27.5, 30],
    amcPercentage: 0.4,
    gstRate: 0.18,
    amcYears: 4,
    chunkSize: 1000,
  });
  const [showDataPreview, setShowDataPreview] = useState(false);
  const [showCacheManager, setShowCacheManager] = useState(false);
  const [useCache] = useState(true);
  const [selectedLocation, setSelectedLocation] = useState("All Locations");
  const [selectedROI, setSelectedROI] = useState(null);
  const [showSplitDetails, setShowSplitDetails] = useState(false);
  const [selectedProduct, setSelectedProduct] = useState(null);
  const [showWithoutGST, setShowWithoutGST] = useState(false);
  const [manualProducts, setManualProducts] = useState([]);
  const [viewMode, setViewMode] = useState("table"); 

  // Update ROI rates array when AMC years change
  useEffect(() => {
    setSettings((prev) => {
      const currentLength = prev.roiRates.length;
      const newLength = prev.amcYears;

      if (currentLength === newLength) return prev; 

      let newRoiRates = [...prev.roiRates];

      if (newLength > currentLength) {
        // Add new rates for additional years (default to last rate + 2.5%)
        const lastRate = newRoiRates[newRoiRates.length - 1] || 30;
        for (let i = currentLength; i < newLength; i++) {
          newRoiRates.push(lastRate + (i - currentLength + 1) * 2.5);
        }
      } else {
        // Trim excess rates
        newRoiRates = newRoiRates.slice(0, newLength);
      }

      return {
        ...prev,
        roiRates: newRoiRates,
      };
    });
  }, [settings.amcYears]);

  // Process Excel data when available 
  const parseExcelDate = (dateValue) => {
    if (!dateValue) return new Date().toISOString().split("T")[0];

    // Handle various date formats
    const dateStr = String(dateValue).trim();

    // If it's already a valid date string
    if (dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
      return dateStr;
    }

    // Handle Excel serial numbers (like 44562 for Oct 18, 2021)
    if (!isNaN(dateStr) && dateStr.length <= 6) {
      const serialNumber = parseInt(dateStr);
      if (serialNumber > 0 && serialNumber < 100000) {
        
const excelEpoch = new Date(1900, 0, 1); 
        let resultDate = new Date(
          excelEpoch.getTime() + (serialNumber - 1) * 24 * 60 * 60 * 1000
        );

       
        // Excel incorrectly thinks 1900 was a leap year and includes Feb 29, 1900
        if (serialNumber > 59) {
          // 59 = Feb 28, 1900 in Excel
          resultDate = new Date(resultDate.getTime() - 24 * 60 * 60 * 1000); // Subtract 1 day
        }

        return resultDate.toISOString().split("T")[0];
      }
    }

    // Handle DD-MMM-YY format specifically (like "18-Oct-21") 
    const ddMmmYyMatch = dateStr.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2})$/);
    if (ddMmmYyMatch) {
      const [, day, monthStr, year] = ddMmmYyMatch;
      const monthMap = {
        jan: "01",
        feb: "02",
        mar: "03",
        apr: "04",
        may: "05",
        jun: "06",
        jul: "07",
        aug: "08",
        sep: "09",
        oct: "10",
        nov: "11",
        dec: "12",
      };
      const month = monthMap[monthStr.toLowerCase()];
      if (month) {
        // Handle 2-digit years - assume 2000s for years 00-99
        const fullYear = `20${year}`;
        const formattedDate = `${fullYear}-${month}-${day.padStart(2, "0")}`;
        
        return formattedDate;
      }
    }

    // Handle DD-MMM-YYYY format (4-digit year)
    const ddMmmYyyyMatch = dateStr.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
    if (ddMmmYyyyMatch) {
      const [, day, monthStr, year] = ddMmmYyyyMatch;
      const monthMap = {
        jan: "01",
        feb: "02",
        mar: "03",
        apr: "04",
        may: "05",
        jun: "06",
        jul: "07",
        aug: "08",
        sep: "09",
        oct: "10",
        nov: "11",
        dec: "12",
      };
      const month = monthMap[monthStr.toLowerCase()];
      if (month) {
        const formattedDate = `${year}-${month}-${day.padStart(2, "0")}`;
        
        return formattedDate;
      }
    }

    // Handle text dates like "Oct 18, 2021"
    try {
      const parsedDate = new Date(dateStr);
      if (!isNaN(parsedDate.getTime())) {
        const formattedDate = parsedDate.toISOString().split("T")[0];
        
        return formattedDate;
      }
    } catch (error) {
      console.warn(`⚠️ Failed to parse date: ${dateStr}`);
    }

    // Default fallback
    const today = new Date().toISOString().split("T")[0];
    console.warn(
      `⚠️ Using today's date as fallback for: ${dateStr} -> ${today}`
    );
    return today;
  };

  const styles = {
    button: {
      padding: "12px 24px",
      borderRadius: "12px",
      border: "none",
      cursor: "pointer",
      fontSize: "0.875rem",
      fontWeight: 600,
      transition: "all 0.2s ease-in-out",
      display: "flex",
      alignItems: "center",
      gap: "8px",
    },
    primaryButton: {
      background: "linear-gradient(135deg, #3b82f6 0%, #8b5cf6 100%)",
      color: "white",
    },
    secondaryButton: {
      background: "#f1f5f9",
      color: "#475569",
      border: "1px solid #e2e8f0",
    },
    input: {
      width: "100%",
      padding: "8px 12px",
      border: "1px solid #d1d5db",
      borderRadius: "6px",
      fontSize: "0.875rem",
      backgroundColor: "white",
      color: "#374151",
    },
    label: {
      display: "block",
      marginBottom: "8px",
      fontWeight: 600,
      color: "#374151",
    },
    tab: (isActive) => ({
      padding: "8px 16px",
      backgroundColor: isActive ? "#3b82f6" : "transparent",
      color: isActive ? "white" : "#64748b",
      border: "none",
      borderRadius: "6px",
      fontSize: "0.875rem",
      fontWeight: 600,
      cursor: "pointer",
      transition: "all 0.2s ease",
    }),
  };

  const processedExcelData = useMemo(() => {
    let processed = [];

    // Process Excel data if available
    if (excelData && Object.keys(excelData).length > 0) {
      const firstSheetName = Object.keys(excelData)[0];
      const sheetData = excelData[firstSheetName] || [];

      processed = sheetData
        .map((row, index) => ({
          id: `excel_${index}`,
          productName:
            row["Item Name"] ||
            row["Product Name"] ||
            row.productName ||
            row.name ||
            `Product ${index + 1}`,
          invoiceValue: (() => {
            let costStr = String(
              row["Cost"] ||
                row["Invoice Value"] ||
                row.invoiceValue ||
                row.cost ||
                0
            );
            costStr = costStr
              .replace(/[₹$,\s]/g, "")
              .replace(/\.00$/, "")
              .trim();
            const parsed = parseFloat(costStr || 0);
            if (parsed > 0 && parsed < 1000) {
              console.warn(
                `⚠️ Small cost value detected for ${
                  row["Item Name"] || "Unknown"
                }: ₹${parsed}. Original: ${row["Cost"]}`
              );
            }
            return parsed;
          })(),
          invoiceNumber: String(
            row["Invoice Number"] ||
            row.invoiceNumber ||
            row["Invoice No"] ||
            row["Invoice No."] || ''
          ),
          location: row.Location || row.location || "Unknown",
          //  DATE PARSING 
          uatDate: parseExcelDate(
            row["UAT Date"] ||row["UAT DATE"] ||
              row.uatDate ||
              row.date ||
              row["Date"] ||
              row["Installation Date"] ||
              row["Go Live Date"]
          ),
          quantity: parseInt(
            String(row.Quantity || row.quantity || 1).replace(/,/g, "") || 1
          ),
          category: row.Category || row.category || "General",
          roi: parseFloat(
            String(row.ROI || row.roi || 0).replace(/[(),%]/g, "") || 0
          ),
        }))
        .filter((item) => {
          if (!item.productName) {
            console.warn("⚠️ Product filtered out: Missing product name");
            return false;
          }
          if (item.invoiceValue <= 0) {
            console.warn(
              `⚠️ Product filtered out: Invalid cost for ${item.productName}: ₹${item.invoiceValue}`
            );
            return false;
          }
          
          return true;
        });
    }

    // Add manual products
    const combinedData = [...processed, ...manualProducts];
    return combinedData;
  }, [excelData, hasExcelData, fileName, manualProducts]);

  // Web Worker hook
  const {
    isReady: workerReady,
    error: workerError,
    progress,
    results,
    summary,
    calculateAMCForDataset,
    terminate,
  } = useAMCCalculationWorker();

  // Cache hook (now processedExcelData is available)
  const {
    isSupported: cacheSupported,
    cachedResult,
    hasCachedResult,
    storeCachedResult,
    isOnline,
    cacheStats,
  } = useAMCCache(processedExcelData, settings);

  // Check if we have calculations (from worker or cache)
  const hasCalculations = results.length > 0 || hasCachedResult;
  const isCalculating = progress.isProcessing;
  const rawResults =
    hasCachedResult && useCache ? cachedResult.results : results;
  const currentSummary =
    hasCachedResult && useCache ? cachedResult.metadata : summary;

  // Transform data to match Excel AMC Schedule format 
  const currentResults = useMemo(() => {
    if (!rawResults || rawResults.length === 0) return [];

    return rawResults.map((product) => {
      // Start with base product data
      const row = {
        productName: product.productName,
        location: product.location || "Unknown",
        invoiceValue: product.invoiceValue,
        invoiceNumber: product.invoiceNumber || '',
        quantity: product.quantity,
        amcStartDate: product.amcStartDate,
        uatDate: product.uatDate,
      };

      // Add quarter data as individual columns (quarter-year format for readability)
      if (product.quarters && Array.isArray(product.quarters)) {
        // Sort quarters chronologically before processing
        const sortedQuarters = [...product.quarters].sort((a, b) => {
          if (a.year !== b.year) return a.year - b.year;
          const quarterOrder = { JFM: 0, AMJ: 1, JAS: 2, OND: 3 };
          return quarterOrder[a.quarter] - quarterOrder[b.quarter];
        });

        sortedQuarters.forEach((quarter) => {
          const quarterKey = `${quarter.quarter}-${quarter.year}`;
          row[quarterKey] = quarter.totalAmount; // Amount with GST
        });
      }

      return row;
    });
  }, [rawResults]);

  // Process results for display with totals and options
  const {
    filteredResults,
    quarterTotals,
    yearTotals,
    locationOptions,
    roiOptions,
  } = useMemo(() => {
    if (!currentResults || currentResults.length === 0) {
      return {
        filteredResults: [],
        quarterTotals: {},
        yearTotals: {},
        locationOptions: ["All Locations"],
        roiOptions: [],
      };
    }

    // Get unique locations and ROI options
    const locations = [
      "All Locations",
      ...new Set(currentResults.map((r) => r.location).filter(Boolean)),
    ];
    const roiGroups = [settings.roiRates]; // For now, just use current settings

    // Simple filtering (can be enhanced later)
    const filtered = currentResults;

    // Calculate quarter and year totals from the flattened data
    const qTotals = {};
    const yTotals = {};

    // Extract quarter data from rawResults for totals calculation
    if (rawResults && rawResults.length > 0) {
      rawResults.forEach((product) => {
        if (product.quarters && Array.isArray(product.quarters)) {
          product.quarters.forEach((quarter) => {
            const key = `${quarter.quarter}-${quarter.year}`;

            // Quarter totals
            if (!qTotals[key]) {
              qTotals[key] = {
                withGst: 0,
                withoutGst: 0,
                quarter: quarter.quarter,
                year: quarter.year,
              };
            }
            qTotals[key].withGst += quarter.totalAmount || 0;
            qTotals[key].withoutGst += quarter.baseAmount || 0;

            // Year totals
            if (!yTotals[quarter.year]) {
              yTotals[quarter.year] = { withGst: 0, withoutGst: 0 };
            }
            yTotals[quarter.year].withGst += quarter.totalAmount || 0;
            yTotals[quarter.year].withoutGst += quarter.baseAmount || 0;
          });
        }
      });
    }

    return {
      filteredResults: filtered,
      quarterTotals: qTotals,
      yearTotals: yTotals,
      locationOptions: locations,
      roiOptions: roiGroups,
    };
  }, [currentResults, rawResults, settings.roiRates]);

  const filteredQuarterTotals = useMemo(() => {
  if (!selectedYear && !selectedQuarter) {
    return quarterTotals;
  }
  
  return Object.fromEntries(
    Object.entries(quarterTotals).filter(([key, totals]) => {
      const [quarter, year] = key.split("-");
      
      // Filter by year if selected
      if (selectedYear && year !== selectedYear) {
        return false;
      }
      
      // Filter by quarter if selected
      if (selectedQuarter && quarter !== selectedQuarter) {
        return false;
      }
      
      return true;
    })
  );
}, [quarterTotals, selectedYear, selectedQuarter]);

// Filter yearly cards based on selected year
const filteredYearTotals = useMemo(() => {
  if (!selectedYear) {
    return yearTotals;
  }
  
  return Object.fromEntries(
    Object.entries(yearTotals).filter(([year, totals]) => {
      return year === selectedYear;
    })
  );
}, [yearTotals, selectedYear]);

  // Handle AMC calculation
  const handleCalculate = useCallback(async () => {
    
    if (!hasExcelData && processedExcelData.length === 0) {
      
      alert(
        "Please upload an Excel file from the dashboard or add manual products below"
      );
      return;
    }

    if (processedExcelData.length === 0) {
    
      alert(
        "No valid products found. Please:\n- Upload an Excel file with Item Name, Cost, and Location columns, OR\n- Add products manually using the form above"
      );
      return;
    }

    if (!workerReady) {
      
      alert(
        "Calculation engine is loading. Please wait a moment and try again."
      );
      return;
    }

    try {
      
      if (hasCachedResult && useCache) {
        
        return;
      }

      // Start fresh calculation
      await calculateAMCForDataset(processedExcelData, settings);

      // Store results in cache after successful calculation
      if (results.length > 0) {
        await storeCachedResult(results, {
          timestamp: Date.now(),
          settings: { ...settings },
          productCount: processedExcelData.length,
          fileName: fileName || "unknown",
        });
      }
    } catch (error) {
      console.error("❌ Calculation error:", error);
    }
  }, [
    hasExcelData,
    isCalculating,
    workerReady,
    hasCachedResult,
    useCache,
    processedExcelData,
    settings,
    calculateAMCForDataset,
    results,
    storeCachedResult,
    fileName,
    excelData,
  ]);

  // Generate dynamic table columns based on calculation results (like Excel AMC Schedule sheet)
  const tableColumns = useMemo(() => {
    // Base columns (matching Excel AMC Schedule sheet)
    const baseColumns = [
      {
        key: "productName",
        title: "Item Name",
        width: 200,
        filterable: true,
      },
      {
        key: "location",
        title: "Location",
        width: 120,
        filterable: true,
      },
      {
        key: "invoiceValue",
        title: "Cost",
        width: 100,
        className: "text-right",
      },
      {
        key: "invoiceNumber",
        title: "Invoice Number",
        width: 150,
        className: "text-left",
        formatter: (value, row) => value || "-"
      },
      {
        key: "quantity",
        title: "Quantity",
        width: 80,
        className: "text-center",
      },
      {
        key: "amcStartDate",
        title: "AMC Start Date",
        width: 120,
      },
      {
        key: "uatDate",
        title: "UAT Date",
        width: 120,
      },
    ];

    // Generate quarter columns dynamically from results
    const quarterColumns = [];
    const quarterSet = new Set();

    // Extract unique quarters from ALL products (not just first one)
    if (currentResults && currentResults.length > 0) {
      // Look for quarter columns in ALL products to collect complete quarter set
      currentResults.forEach((product) => {
        Object.keys(product).forEach((key) => {
          
          if (key.match(/^[A-Z]{3}-\d{4}$/)) {
            quarterSet.add(key);
          }
        });
      });
    }

    // Sort quarters chronologically (year first, then quarter within year)
    const quarterOrder = { JFM: 0, AMJ: 1, JAS: 2, OND: 3 };
    const sortedQuarters = Array.from(quarterSet).sort((a, b) => {
      const [qA, yearA] = a.split("-");
      const [qB, yearB] = b.split("-");

      // First sort by year
      if (parseInt(yearA) !== parseInt(yearB)) {
        return parseInt(yearA) - parseInt(yearB);
      }

      // Then sort by quarter within the same year
      return quarterOrder[qA] - quarterOrder[qB];
    });

    // Create column definitions for each quarter
    sortedQuarters.forEach((quarterKey) => {
      quarterColumns.push({
        key: quarterKey,
        title: quarterKey,
        width: 120,
        className: "text-right",
      });
    });

    return [...baseColumns, ...quarterColumns];
  }, [currentResults]);

  // Cleanup worker on unmount
  useEffect(() => {
    return () => {
      terminate();
    };
  }, [terminate]);

  const containerStyle = {
    minHeight: "100vh",
    background:
      "linear-gradient(135deg, #f8fafc 0%, #e2e8f0 50%, #cbd5e1 100%)",
    color: "#1e293b",
    fontFamily:
      '"Inter", "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif',
    padding: "40px",
    position: "relative",
  };

  const backgroundOverlayStyle = {
    position: "absolute",
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    background: `
      radial-gradient(circle at 20% 80%, rgba(59, 130, 246, 0.1) 0%, transparent 50%),
      radial-gradient(circle at 80% 20%, rgba(139, 92, 246, 0.08) 0%, transparent 50%)
    `,
    pointerEvents: "none",
  };

  // Updated Quarter Payment Card Component with Shared Payment Status
const QuarterPaymentCard = ({ quarter, year, totals, showWithoutGST, paidQuarters, updatePaidQuarters }) => {
  const quarterKey = `${quarter}-${year}`;
  const isPaid = paidQuarters[quarterKey]?.paid || false;
  const paidDate = paidQuarters[quarterKey]?.date || "";
  
  return (
    <div
      style={{
        background: isPaid 
          ? "linear-gradient(135deg, #6a2ce6ff 0%, #6813a0ce 100%)" 
          : "linear-gradient(135deg, #3b82f6 0%, #8b5cf6 100%)",
        borderRadius: "16px",
        padding: "24px",
        color: "white",
        boxShadow: isPaid 
          ? "0 10px 30px -10px rgba(5, 150, 105, 0.3)"
          : "0 10px 30px -10px rgba(59, 130, 246, 0.3)",
        position: "relative",
        transition: "all 0.3s ease",
      }}
    >
      {/* Payment Status Badge */}
      <div style={{
        position: "absolute",
        top: "12px",
        right: "12px",
        display: "flex",
        alignItems: "center",
        gap: "8px"
      }}>
        <label style={{
          display: "flex",
          alignItems: "center",
          cursor: "pointer",
          fontSize: "0.75rem",
          fontWeight: 600,
          opacity: 0.9
        }}>
          <input
            type="checkbox"
            checked={isPaid}
            onChange={(e) => {
              if (e.target.checked) {
                const payDate = prompt("Enter payment date (YYYY-MM-DD):", new Date().toISOString().split("T")[0]);
                if (payDate) {
                  updatePaidQuarters(quarterKey, true, payDate);
                }
              } else {
                updatePaidQuarters(quarterKey, false, "");
              }
            }}
            style={{
              marginRight: "4px",
              accentColor: "white",
              transform: "scale(1.2)"
            }}
          />
          {isPaid ? "PAID" : "UNPAID"}
        </label>
      </div>
      
      <div style={{
        fontSize: "0.875rem",
        fontWeight: 600,
        marginBottom: "8px",
        opacity: 0.9,
        marginTop: "20px"
      }}>
        {quarter} {year}
      </div>
      
      <div style={{
        fontSize: "1.5rem",
        fontWeight: 800,
        marginBottom: "4px"
      }}>
        ₹{Math.round(totals.withGst).toLocaleString()}
      </div>
      
      <div style={{
        fontSize: "0.75rem",
        opacity: 0.8,
        marginBottom: "8px"
      }}>
        {totals.withGst >= 10000000 
          ? `₹${(totals.withGst / 10000000).toFixed(2)} Cr`
          : totals.withGst >= 100000
          ? `₹${(totals.withGst / 100000).toFixed(2)} Lakh`
          : "Amount"
        }
      </div>
      
      {showWithoutGST && (
        <div style={{
          fontSize: "1rem",
          marginBottom: "8px",
          padding: "4px 8px",
          backgroundColor: "rgba(182, 76, 252, 0.91)",
          borderRadius: "6px",
          display: "inline-block",
        }}>
          Without GST: ₹{Math.round(totals.withoutGst).toLocaleString()}
        </div>
      )}
      
      {/* Payment Date */}
      {isPaid && paidDate && (
        <div style={{
          fontSize: "0.75rem",
          opacity: 0.9,
          borderTop: "1px solid rgba(255, 255, 255, 0.65)",
          paddingTop: "8px",
          marginTop: "8px"
        }}>
          Paid on: {new Date(paidDate).toLocaleDateString('en-GB')}
        </div>
      )}
    </div>
  );
};

  const headerStyle = {
    background: "rgba(255, 255, 255, 0.95)",
    backdropFilter: "blur(20px)",
    borderRadius: "20px",
    padding: "32px",
    marginBottom: "40px",
    border: "1px solid rgba(255, 255, 255, 0.2)",
    boxShadow: "0 20px 40px -12px rgba(0, 0, 0, 0.15)",
    position: "relative",
    zIndex: 1,
  };

  const cardStyle = {
    background: "rgba(255, 255, 255, 0.95)",
    backdropFilter: "blur(20px)",
    borderRadius: "20px",
    padding: "32px",
    marginBottom: "32px",
    border: "1px solid rgba(255, 255, 255, 0.2)",
    boxShadow: "0 20px 40px -12px rgba(0, 0, 0, 0.15)",
    position: "relative",
    zIndex: 1,
  };

  // Export Component
const ExportControls = () => {
  const [showExportOptions, setShowExportOptions] = useState(false);

  const handleExport = (type) => {
    if (!hasCalculations) {
      alert('Please calculate AMC schedule first before exporting.');
      return;
    }

    const exporter = new AMCExportManager(filteredResults, settings, paidQuarters);
    const baseFilename = fileName ? fileName.replace(/\.[^/.]+$/, '') : 'AMC_Schedule';

    switch (type) {
      case 'excel':
        exporter.exportToExcel(baseFilename);
        break;
      case 'csv':
        exporter.exportToCSV(baseFilename);
        break;
      case 'pdf':
        exporter.exportToPDF(baseFilename);
        break;
      case 'json':
        exporter.exportToJSON(baseFilename);
        break;
      default:
        break;
    }

    setShowExportOptions(false);
  };

  return (
    <div style={{ position: 'relative' }}>
      <button
        onClick={() => setShowExportOptions(!showExportOptions)}
        disabled={!hasCalculations}
        style={{
          ...styles.button,
          ...styles.primaryButton,
          opacity: hasCalculations ? 1 : 0.5,
          cursor: hasCalculations ? 'pointer' : 'not-allowed',
        }}
      >
        <Download size={16} />
        Export Data
      </button>

      {showExportOptions && (
        <div
          style={{
            position: 'absolute',
            top: '100%',
            right: 0,
            backgroundColor: 'white',
            border: '1px solid #e2e8f0',
            borderRadius: '8px',
            boxShadow: '0 10px 25px rgba(0, 0, 0, 0.15)',
            zIndex: 1000,
            minWidth: '200px',
            marginTop: '8px',
          }}
        >
          <div style={{ padding: '8px 0' }}>
            <button
              onClick={() => handleExport('excel')}
              style={{
                width: '100%',
                padding: '12px 16px',
                border: 'none',
                backgroundColor: 'transparent',
                textAlign: 'left',
                cursor: 'pointer',
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
                fontSize: '0.875rem',
              }}
              onMouseEnter={(e) => e.target.style.backgroundColor = '#f8fafc'}
              onMouseLeave={(e) => e.target.style.backgroundColor = 'transparent'}
            >
              <FileSpreadsheet size={16} color="#059669" />
              Export to Excel
            </button>
            
            <button
              onClick={() => handleExport('csv')}
              style={{
                width: '100%',
                padding: '12px 16px',
                border: 'none',
                backgroundColor: 'transparent',
                textAlign: 'left',
                cursor: 'pointer',
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
                fontSize: '0.875rem',
              }}
              onMouseEnter={(e) => e.target.style.backgroundColor = '#f8fafc'}
              onMouseLeave={(e) => e.target.style.backgroundColor = 'transparent'}
            >
              <Database size={16} color="#3b82f6" />
              Export to CSV
            </button>
            
            <button
              onClick={() => handleExport('pdf')}
              style={{
                width: '100%',
                padding: '12px 16px',
                border: 'none',
                backgroundColor: 'transparent',
                textAlign: 'left',
                cursor: 'pointer',
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
                fontSize: '0.875rem',
              }}
              onMouseEnter={(e) => e.target.style.backgroundColor = '#f8fafc'}
              onMouseLeave={(e) => e.target.style.backgroundColor = 'transparent'}
            >
              <FileText size={16} color="#dc2626" />
              Export to PDF
            </button>
            
            <button
              onClick={() => handleExport('json')}
              style={{
                width: '100%',
                padding: '12px 16px',
                border: 'none',
                backgroundColor: 'transparent',
                textAlign: 'left',
                cursor: 'pointer',
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
                fontSize: '0.875rem',
              }}
              onMouseEnter={(e) => e.target.style.backgroundColor = '#f8fafc'}
              onMouseLeave={(e) => e.target.style.backgroundColor = 'transparent'}
            >
              <Database size={16} color="#8b5cf6" />
              Export to JSON
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

  return (
    <div style={containerStyle}>
      <div style={backgroundOverlayStyle}></div>

      {/* Header */}
      <div style={headerStyle}>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            marginBottom: "16px",
          }}
        >
          <div>
            <button
              onClick={() => navigate("/")}
              style={{
                ...styles.button,
                ...styles.secondaryButton,
                marginBottom: "16px",
              }}
            >
              <ArrowLeft size={16} />
                Back to Dashboard
            </button>

            <h1
              style={{
                fontSize: "2rem",
                fontWeight: 800,
                background: "linear-gradient(135deg, #3b82f6 0%, #8b5cf6 100%)",
                WebkitBackgroundClip: "text",
                WebkitTextFillColor: "transparent",
                margin: "0 0 8px 0",
                display: "flex",
                alignItems: "center",
                gap: "12px",
              }}
            >
              AMC Payment Tracker
            </h1>
            {hasExcelData && (
              <div
                style={{
                  marginTop: "12px",
                  padding: "6px 12px",
                  backgroundColor: "#eff6ff",
                  color: "#2563eb",
                  borderRadius: "20px",
                  fontSize: "0.8rem",
                  fontWeight: 600,
                  border: "1px solid #bfdbfe",
                  display: "inline-block",
                }}
              >
                Source File: {fileName}
              </div>
            )}
          </div>
        </div>
      </div> 
      
      {/* Worker Error Alert */}
      {workerError && (
        <div
          style={{
            ...cardStyle,
            backgroundColor: "#fef2f2",
            border: "1px solid #fecaca",
            color: "#dc2626",
          }}
        >
          <h4 style={{ margin: "0 0 8px 0", fontWeight: 600 }}>
            ⚠️ Worker Error
          </h4>
          <p style={{ margin: 0, fontSize: "0.9rem" }}>{workerError}</p>
        </div>
      )}

      {/* Cache Manager */}
      <CacheManager
        visible={showCacheManager}
        onClose={() => setShowCacheManager(false)}
      />

      {/* Cache Status Compact */}
      {cacheSupported && <CacheManager compact={true} />}

      {/* Progress Indicator */}
      <CalculationProgress
        progress={progress}
        summary={currentSummary}
        isCalculating={isCalculating}
        title="AMC Calculation Progress"
      />

      {/* Data Status */}
      {hasExcelData && (
        <div style={cardStyle}>
          <h3
            style={{
              fontSize: "1.25rem",
              fontWeight: 700,
              color: "#1e293b",
              margin: "0 0 16px 0",
              display: "flex",
              alignItems: "center",
              gap: "8px",
            }}
          >
            📊 Data Status
          </h3>

          {/* Data Preview Toggle */}
          <div style={{ marginBottom: "16px" }}>
            <button
              onClick={() => setShowDataPreview(!showDataPreview)}
              style={{
                padding: "8px 16px",
                backgroundColor: showDataPreview ? "#ef4444" : "#3b82f6",
                color: "white",
                border: "none",
                borderRadius: "6px",
                fontSize: "0.875rem",
                fontWeight: 600,
                cursor: "pointer",
                transition: "all 0.2s ease",
              }}
            >
              {showDataPreview
                ? "Hide Data Preview"
                : "Show Data Preview (First 5 rows)"}
            </button>
          </div>

          {/* Data Preview Table */}
          {showDataPreview && processedExcelData.length > 0 && (
            <div
              style={{
                marginBottom: "20px",
                border: "1px solid #e2e8f0",
                borderRadius: "8px",
                overflow: "auto",
              }}
            >
              <table
                style={{
                  width: "100%",
                  borderCollapse: "collapse",
                  fontSize: "0.875rem",
                }}
              >
                <thead>
                  <tr style={{ backgroundColor: "#f8fafc" }}>
                    <th
                      style={{
                        padding: "8px",
                        border: "1px solid #e2e8f0",
                        textAlign: "left",
                      }}
                    >
                      Product Name
                    </th>
                    <th
                      style={{
                        padding: "8px",
                        border: "1px solid #e2e8f0",
                        textAlign: "right",
                      }}
                    >
                      Invoice Value
                    </th>
                    <th
                      style={{
                        padding: "8px",
                        border: "1px solid #e2e8f0",
                        textAlign: "left",
                      }}
                    >
                      Invoice Number
                    </th>
                    <th
                      style={{
                        padding: "8px",
                        border: "1px solid #e2e8f0",
                        textAlign: "center",
                      }}
                    >
                      Qty
                    </th>
                    <th
                      style={{
                        padding: "8px",
                        border: "1px solid #e2e8f0",
                        textAlign: "left",
                      }}
                    >
                      Location
                    </th>
                    <th
                      style={{
                        padding: "8px",
                        border: "1px solid #e2e8f0",
                        textAlign: "left",
                      }}
                    >
                      UAT Date
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {processedExcelData.slice(0, 5).map((item, index) => (
                    <tr key={index}>
                      <td
                        style={{ padding: "8px", border: "1px solid #e2e8f0" }}
                      >
                        {item.productName}
                      </td>
                      <td
                        style={{
                          padding: "8px",
                          border: "1px solid #e2e8f0",
                          textAlign: "right",
                          fontWeight: 600,
                          color:
                            item.invoiceValue > 10000000
                              ? "#059669"
                              : "#dc2626",
                        }}
                      >
                        ₹ {item.invoiceValue.toLocaleString()}
                        {item.invoiceValue > 10000000 && (
                          <span
                            style={{
                              fontSize: "0.75rem",
                              color: "#059669",
                              marginLeft: "4px",
                            }}
                          >
                            (₹{(item.invoiceValue / 10000000).toFixed(2)} Cr)
                          </span>
                        )}
                      </td>
                      <td
                        style={{ padding: "8px", border: "1px solid #e2e8f0" }}
                        >
                          {item.invoiceNumber || 'N/A'}
                      </td>
                      <td
                        style={{
                          padding: "8px",
                          border: "1px solid #e2e8f0",
                          textAlign: "center",
                        }}
                      >
                        {item.quantity}
                      </td>
                      <td
                        style={{ padding: "8px", border: "1px solid #e2e8f0" }}
                      >
                        {item.location}
                      </td>
                      <td
                        style={{ padding: "8px", border: "1px solid #e2e8f0" }}
                      >
                        {item.uatDate}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>

              <div
                style={{
                  padding: "12px",
                  backgroundColor: "#f0f9ff",
                  borderTop: "1px solid #e2e8f0",
                  fontSize: "0.8rem",
                  color: "#1e40af",
                }}
              >
                <ul style={{ margin: "4px 0", paddingLeft: "16px" }}>
                  <li>
                    <span style={{ color: "#059669" }}>Green amounts</span> are
                    ≥₹1 Crore (correctly parsed large values)
                  </li>
                  <li>
                    <span style={{ color: "#dc2626" }}>Red amounts</span> are
                    &lt;₹1 Crore (check if these should be larger)
                  </li>
                  <li>
                    Total products loaded:{" "}
                    <strong>{processedExcelData.length}</strong>
                  </li>
                  <li>
                    Products ≥₹1 Cr:{" "}
                    <strong style={{ color: "#059669" }}>
                      {
                        processedExcelData.filter(
                          (p) => p.invoiceValue >= 10000000
                        ).length
                      }
                    </strong>
                  </li>
                  <li>
                    Products &lt;₹1 Cr:{" "}
                    <strong style={{ color: "#dc2626" }}>
                      {
                        processedExcelData.filter(
                          (p) => p.invoiceValue < 10000000
                        ).length
                      }
                    </strong>
                  </li>
                </ul>
              </div>
            </div>
          )}

          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))",
              gap: "16px",
            }}
          >
          </div>
        </div>
      )}

      {/* Filters and Settings */}
      {hasCalculations && (
        <div style={cardStyle}>
          <h3
            style={{
              fontSize: "1.25rem",
              fontWeight: 700,
              color: "#1e293b",
              margin: "0 0 24px 0",
              display: "flex",
              alignItems: "center",
              gap: "8px",
            }}
          >
            🔍 Filters & Display Options
          </h3>

          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fit, minmax(250px, 1fr))",
              gap: "20px",
              marginBottom: "24px",
            }}
          >
            <div>
              <label
                style={{
                  display: "block",
                  marginBottom: "8px",
                  fontWeight: 600,
                  color: "#374151",
                }}
              >
                Filter by Location
              </label>
              <select
                value={selectedLocation}
                onChange={(e) => setSelectedLocation(e.target.value)}
                style={{
                  width: "100%",
                  padding: "8px 12px",
                  border: "1px solid #d1d5db",
                  borderRadius: "6px",
                  fontSize: "0.9rem",
                }}
              >
                {locationOptions.map((location) => (
                  <option key={location} value={location}>
                    {location}
                  </option>
                ))}
              </select>
            </div>

            {/* Filter by Year */}
<div>
  <label
    style={{
      display: "block",
      marginBottom: "8px",
      fontWeight: 600,
      color: "#374151",
    }}
  >
    Filter by Year
  </label>
  <select
    value={selectedYear || ""}
    onChange={(e) => setSelectedYear(e.target.value || null)}
    style={{
      width: "100%",
      padding: "8px 12px",
      border: "1px solid #d1d5db",
      borderRadius: "6px",
      fontSize: "0.9rem",
    }}
  >
    <option value="">All Years</option>
    {Object.keys(yearTotals)
      .sort()
      .map((year) => (
        <option key={year} value={year}>
          {year}
        </option>
      ))}
  </select>
</div>

{/* Filter by Quarter */}
<div>
  <label
    style={{
      display: "block",
      marginBottom: "8px",
      fontWeight: 600,
      color: "#374151",
    }}
  >
    Filter by Quarter
  </label>
  <select
    value={selectedQuarter || ""}
    onChange={(e) => setSelectedQuarter(e.target.value || null)}
    style={{
      width: "100%",
      padding: "8px 12px",
      border: "1px solid #d1d5db",
      borderRadius: "6px",
      fontSize: "0.9rem",
    }}
  >
    <option value="">All Quarters</option>
    <option value="JFM">Q1 (JFM)</option>
    <option value="AMJ">Q2 (AMJ)</option>
    <option value="JAS">Q3 (JAS)</option>
    <option value="OND">Q4 (OND)</option>
  </select>
</div>


            <div>
              <label
                style={{
                  display: "block",
                  marginBottom: "8px",
                  fontWeight: 600,
                  color: "#374151",
                }}
              >
                Display Options
              </label>
              <div
                style={{ display: "flex", flexDirection: "column", gap: "8px" }}
              >
                <label
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: "8px",
                    fontSize: "0.9rem",
                  }}
                >
                  <input
                    type="checkbox"
                    checked={showWithoutGST}
                    onChange={(e) => setShowWithoutGST(e.target.checked)}
                  />
                  Show amounts without GST
                </label>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* New Card with Two Main Tabs */}
      <div
        style={{
    background: "rgba(255, 255, 255, 0.95)",
    backdropFilter: "blur(12px)",
    borderRadius: "20px",
    padding: "32px",
    boxShadow: "0 20px 40px rgba(0, 0, 0, 0.1)",
    border: "1px solid rgba(255, 255, 255, 0.2)",
    marginBottom: "32px",
  }}
>
  {/* Main Tab Headers */}
  <div
    style={{
      display: "flex",
      gap: "8px",
      marginBottom: "24px",
      backgroundColor: "#f1f5f9",
      padding: "4px",
      borderRadius: "12px",
    }}
  >
    <button
      onClick={() => setActiveMainTab("processExcel")}
      style={{
        ...styles.tab(activeMainTab === "processExcel"),
        flex: 1,
        padding: "12px 24px",
      }}
    >
      <Upload size={16} style={{ marginRight: "6px" }} /> Process Excel Data
    </button>
    <button
      onClick={() => setActiveMainTab("manualEntry")}
      style={{
        ...styles.tab(activeMainTab === "manualEntry"),
        flex: 1,
        padding: "12px 24px",
      }}
    >
      <Plus size={16} />  Add Product Manually
    </button>
  </div>

  {/* Tab Content */}
  {activeMainTab === "manualEntry" && (
    <div>
      {/* Manual Product Entry */}
      <div style={cardStyle}>
        <h3
          style={{
            fontSize: "1.25rem",
            fontWeight: 700,
            color: "#1e293b",
            margin: "0 0 24px 0",
            display: "flex",
            alignItems: "center",
            gap: "8px",
          }}
        >
          ➕ Manual Product Entry
        </h3>

        <ManualProductForm
          onAddProduct={(product) => {
            setManualProducts((prev) => [...prev, product]);
          }}
        />

        {/* Display Currently Added Manual Products */}
        {manualProducts.length > 0 && (
          <div
            style={{
              marginTop: "20px",
              padding: "16px",
              backgroundColor: "#f0f9ff",
              border: "1px solid #3b82f6",
              borderRadius: "8px",
            }}
          >
            <h4
              style={{
                fontSize: "1rem",
                fontWeight: 600,
                color: "#1e293b",
                margin: "0 0 12px 0",
                display: "flex",
                alignItems: "center",
                gap: "8px",
              }}
            >
              📋 Manual Products Added ({manualProducts.length})
            </h4>

            <div
              style={{ display: "flex", flexDirection: "column", gap: "8px" }}
            >
              {manualProducts.map((product, index) => (
                <div
                  key={product.id}
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    padding: "8px 12px",
                    backgroundColor: "white",
                    border: "1px solid #e2e8f0",
                    borderRadius: "6px",
                    fontSize: "0.875rem",
                  }}
                >
                  <div style={{ display: "flex", gap: "16px", flex: 1 }}>
                    <span style={{ fontWeight: 600, color: "#1e293b" }}>
                      {product.productName}
                    </span>
                    <span style={{ color: "#64748b" }}>
                      ₹{product.invoiceValue.toLocaleString()}
                    </span>
                    <span style={{ color: "#64748b" }}>
                      {product.invoiceNumber || 'N/A'}
                    </span>
                    <span style={{ color: "#64748b" }}>{product.location}</span>
                    <span style={{ color: "#64748b" }}>
                      UAT: {new Date(product.uatDate).toLocaleDateString()}
                    </span>
                  </div>
                  <button
                    onClick={() => {
                      setManualProducts((prev) =>
                        prev.filter((_, i) => i !== index)
                      );
                    }}
                    style={{
                      background: "none",
                      border: "none",
                      color: "#ef4444",
                      cursor: "pointer",
                      fontSize: "1.2rem",
                      padding: "4px",
                      borderRadius: "4px",
                    }}
                    title="Remove product"
                  >
                    ✕
                  </button>
                </div>
              ))}
            </div>

            <div
              style={{
                marginTop: "12px",
                padding: "8px",
                backgroundColor: "#dbeafe",
                borderRadius: "4px",
                fontSize: "0.8rem",
                color: "#1d4ed8",
                textAlign: "center",
              }}
            >
              Ready to calculate! Click "Calculate AMC Schedule" below to
              process these products.
            </div>
          </div>
        )}

        {/* AMC Settings & Calculation */}
        {(hasExcelData || processedExcelData.length > 0) && (
          <div>
            <h3
              style={{
                fontSize: "1.25rem",
                fontWeight: 700,
                color: "#1e293b",
                margin: "0 0 24px 0",
                display: "flex",
                alignItems: "center",
                gap: "8px",
              }}
            >
              ⚙️ AMC Calculation Settings
            </h3>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(auto-fit, minmax(250px, 1fr))",
                gap: "20px",
                marginBottom: "24px",
              }}
            >
              <div>
                <label
                  style={{
                    display: "block",
                    marginBottom: "8px",
                    fontWeight: 600,
                    color: "#374151",
                  }}
                >
                  AMC Percentage
                </label>
                <input
                  type="number"
                  value={settings.amcPercentage * 100}
                  onChange={(e) =>
                    setSettings((prev) => ({
                      ...prev,
                      amcPercentage: parseFloat(e.target.value) / 100,
                    }))
                  }
                  min="1"
                  max="100"
                  step="0.1"
                  style={{
                    width: "100%",
                    padding: "8px 12px",
                    border: "1px solid #d1d5db",
                    borderRadius: "6px",
                    fontSize: "0.9rem",
                  }}
                />
                <p
                  style={{
                    fontSize: "0.75rem",
                    color: "#6b7280",
                    margin: "4px 0 0 0",
                  }}
                >
                  Current: {(settings.amcPercentage * 100).toFixed(1)}%
                </p>
              </div>

              <div>
                <label
                  style={{
                    display: "block",
                    marginBottom: "8px",
                    fontWeight: 600,
                    color: "#374151",
                  }}
                >
                  GST Rate
                </label>
                <input
                  type="number"
                  value={settings.gstRate * 100}
                  onChange={(e) =>
                    setSettings((prev) => ({
                      ...prev,
                      gstRate: parseFloat(e.target.value) / 100,
                    }))
                  }
                  min="0"
                  max="50"
                  step="0.1"
                  style={{
                    width: "100%",
                    padding: "8px 12px",
                    border: "1px solid #d1d5db",
                    borderRadius: "6px",
                    fontSize: "0.9rem",
                  }}
                />
                <p
                  style={{
                    fontSize: "0.75rem",
                    color: "#6b7280",
                    margin: "4px 0 0 0",
                  }}
                >
                  Current: {(settings.gstRate * 100).toFixed(1)}%
                </p>
              </div>

              <div>
                <label
                  style={{
                    display: "block",
                    marginBottom: "8px",
                    fontWeight: 600,
                    color: "#374151",
                  }}
                >
                  AMC Duration (Years)
                </label>
                <select
                  value={settings.amcYears}
                  onChange={(e) =>
                    setSettings((prev) => ({
                      ...prev,
                      amcYears: parseInt(e.target.value),
                    }))
                  }
                  style={{
                    width: "100%",
                    padding: "8px 12px",
                    border: "1px solid #d1d5db",
                    borderRadius: "6px",
                    fontSize: "0.9rem",
                  }}
                >
                  <option value={3}>3 Years</option>
                  <option value={4}>4 Years</option>
                  <option value={5}>5 Years</option>
                  <option value={6}>6 Years</option>
                  <option value={7}>7 Years</option>
                  <option value={8}>8 Years</option>
                  <option value={9}>9 Years</option>
                  <option value={10}>10 Years</option>
                </select>
              </div>

              <div>
                <label
                  style={{
                    display: "block",
                    marginBottom: "8px",
                    fontWeight: 600,
                    color: "#374151",
                  }}
                >
                  Processing Chunk Size
                </label>
                <select
                  value={settings.chunkSize}
                  onChange={(e) =>
                    setSettings((prev) => ({
                      ...prev,
                      chunkSize: parseInt(e.target.value),
                    }))
                  }
                  style={{
                    width: "100%",
                    padding: "8px 12px",
                    border: "1px solid #d1d5db",
                    borderRadius: "6px",
                    fontSize: "0.9rem",
                  }}
                >
                  <option value={500}>500 (Slower, Stable)</option>
                  <option value={1000}>1,000 (Balanced)</option>
                  <option value={2000}>2,000 (Faster)</option>
                  <option value={5000}>5,000 (Maximum Speed)</option>
                </select>
                <p
                  style={{
                    fontSize: "0.75rem",
                    color: "#6b7280",
                    margin: "4px 0 0 0",
                  }}
                >
                  Larger chunks = faster processing
                </p>
              </div>
            </div>

            {/* ROI Rates */}
            <div style={{ marginBottom: "24px" }}>
              <label
                style={{
                  display: "block",
                  marginBottom: "12px",
                  fontWeight: 600,
                  color: "#374151",
                }}
              >
                ROI Rates (%) - {settings.amcYears} Year
                {settings.amcYears > 1 ? "s" : ""}
              </label>
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns:
                    settings.amcYears <= 5
                      ? "repeat(auto-fit, minmax(150px, 1fr))"
                      : "repeat(auto-fit, minmax(120px, 1fr))",
                  gap: "12px",
                  maxWidth: "100%",
                }}
              >
                {Array.from({ length: settings.amcYears }, (_, index) => {
                  const rate = settings.roiRates[index] || 0;
                  return (
                    <div
                      key={index}
                      style={{
                        display: "flex",
                        flexDirection: "column",
                        alignItems: "center",
                        gap: "4px",
                        padding: "12px",
                        backgroundColor: "#f8fafc",
                        border: "1px solid #e2e8f0",
                        borderRadius: "8px",
                      }}
                    >
                      <span
                        style={{
                          fontSize: "0.875rem",
                          color: "#4b5563",
                          fontWeight: 600,
                          marginBottom: "4px",
                        }}
                      >
                        Year {index + 1}
                      </span>
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: "4px",
                        }}
                      >
                        <input
                          type="number"
                          value={rate}
                          onChange={(e) => {
                            const newRates = [...settings.roiRates];
                            newRates[index] = parseFloat(e.target.value) || 0;
                            setSettings((prev) => ({
                              ...prev,
                              roiRates: newRates,
                            }));
                          }}
                          min="0"
                          max="100"
                          step="0.5"
                          placeholder="0"
                          style={{
                            width: "70px",
                            padding: "6px 8px",
                            border: "1px solid #d1d5db",
                            borderRadius: "4px",
                            fontSize: "0.875rem",
                            textAlign: "center",
                          }}
                        />
                        <span style={{ fontSize: "0.75rem", color: "#6b7280" }}>
                          %
                        </span>
                      </div>
                    </div>
                  );
                })}
              </div>
              <div
                style={{
                  marginTop: "12px",
                  padding: "8px 12px",
                  backgroundColor: "#f0f9ff",
                  border: "1px solid #bfdbfe",
                  borderRadius: "6px",
                  fontSize: "0.8rem",
                  color: "#1e40af",
                }}
              >
                💡 <strong>Tip:</strong> ROI rates typically increase over time.
                Default progression adds 2.5% per year when extending duration.
              </div>
            </div>
            {/* Calculate Button */}
            <div style={{ textAlign: "center" }}>
              <button
                onClick={handleCalculate}
                disabled={
                  (!hasExcelData && processedExcelData.length === 0) ||
                  isCalculating ||
                  !workerReady
                }
                style={{
                  padding: "16px 32px",
                  backgroundColor:
                    (!hasExcelData && processedExcelData.length === 0) ||
                    isCalculating ||
                    !workerReady
                      ? "#9ca3af"
                      : "#862debff",
                  color: "white",
                  border: "none",
                  borderRadius: "12px",
                  fontSize: "1.1rem",
                  fontWeight: 700,
                  cursor:
                    (!hasExcelData && processedExcelData.length === 0) ||
                    isCalculating ||
                    !workerReady
                      ? "not-allowed"
                      : "pointer",
                  transition: "all 0.2s ease",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  gap: "12px",
                  margin: "0 auto",
                  minWidth: "250px",
                }}
              >
                {isCalculating ? (
                  <>
                    <span
                      style={{
                        display: "inline-block",
                        width: "16px",
                        height: "16px",
                        border: "2px solid #ffffff",
                        borderTop: "2px solid transparent",
                        borderRadius: "50%",
                        animation: "spin 1s linear infinite",
                      }}
                    ></span>
                    Calculating...
                  </>
                ) : hasCachedResult ? (
                  <>Load from Cache</>
                ) : (
                  <> Get AMC Tracker</>
                )}
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  )}

  {activeMainTab === "processExcel" && (
  <div>
    <div
      style={{
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
        marginBottom: "24px",
      }}
    >
      <div>
        <h3
          style={{
            fontSize: "1.25rem",
            fontWeight: 700,
            color: "#1e293b",
            display: "flex",
            alignItems: "center",
            gap: "8px",
          }}
        >
          Process Excel Data
        </h3>
        
        {/* Show filename if available */}
        {fileName && (
          <p
            style={{
              fontSize: "0.875rem",
              color: "#64748b",
              margin: "4px 0 0 0",
            }}
          >
            Process Excel file: <strong>{fileName}</strong>
          </p>
        )}
      </div>

      {/* View Toggle for Process Excel Data */}
      {hasCalculations && (
        <div
          style={{
            display: "flex",
            gap: "8px",
            backgroundColor: "#f1f5f9",
            padding: "4px",
            borderRadius: "8px",
          }}
        >
          <button
            onClick={() => setProcessDataView("cards")}
            style={{
              ...styles.tab(processDataView === "cards"),
              padding: "8px 16px",
              fontSize: "0.875rem",
            }}
          >
            Cards View
          </button>
          <button
            onClick={() => setProcessDataView("table")}
            style={{
              ...styles.tab(processDataView === "table"),
              padding: "8px 16px",
              fontSize: "0.875rem",
            }}
          >
            Table View
          </button>
          
        <ExportControls />
        </div>
      )}
    </div>

    {/* Process Excel Data Content */}
    {!hasCalculations ? (
      // Show calculation trigger when no results
      <div style={{ textAlign: "center", padding: "40px 0" }}>
        <div
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: "12px",
            padding: "16px 32px",
            backgroundColor: "#dbeafe",
            color: "#1e40af",
            borderRadius: "12px",
            marginBottom: "24px",
          }}
        >
          <Calculator size={24} />
          <span style={{ fontSize: "1.1rem", fontWeight: 600 }}>
            Ready to calculate AMC schedules
          </span>
        </div>
        
        <div style={{ display: "flex", justifyContent: "center", width: "100%" }}>
          <button
            onClick={handleCalculate}
            disabled={isCalculating || !workerReady}
            style={{
              ...styles.button,
              ...styles.primaryButton,
              fontSize: "1rem",
              padding: "16px 32px",
            }}
          >
            {isCalculating ? "Calculating..." : "Calculate AMC Schedule"}
          </button>
        </div>
      </div>
    ) : (
      //based on selected view
      <>
        {processDataView === "table" && (
          hasCalculations && (
            <div>
              <VirtualDataTable
                data={filteredResults}
                columns={tableColumns}
                height={600}
                title={
                  <h3
            style={{
              fontSize: "1.25rem",
              fontWeight: 700,
              color: "#1e293b",
              margin: "0 0 24px 0",
              display: "flex",
              alignItems: "center",
              gap: "8px",
            }}
          >
            📊 AMC Schedule Table - {filteredResults.length} products
          </h3>}
                searchable={true}
                filterable={true}
                sortable={true}
                showTotals={true}
                showRowTotals={true}
                summary={{
                  ...currentSummary,
                  totalProducts: filteredResults.length,
                }}
                formatters={{
                  invoiceValue: (value) => `₹${value?.toLocaleString() || "0"}`,
                  invoiceNumber: (value) => value || "-",
                  quantity: (value) => value?.toLocaleString() || "0",
                  amcStartDate: (value) =>
                    value ? new Date(value).toLocaleDateString() : "-",
                  uatDate: (value) =>
                    value ? new Date(value).toLocaleDateString() : "-",
                  
                  ...Object.fromEntries(
                    tableColumns
                      .filter(
                        (col) =>
                          col.key.includes("-") &&
                          col.key.match(/^[A-Z]{3}-\d{4}$/)
                      )
                      .map((col) => [
                        col.key,
                        (value) => (value ? `₹${value.toLocaleString()}` : "₹0"),
                      ])
                  ),
                }}
              />
            </div>
          )
        )}

        {processDataView === "cards" && (
          <div>
            {/* Quarter-wise Summary - Payment Status Design */}
{hasCalculations && Object.keys(filteredQuarterTotals).length > 0 && (
  <div 
    style={{
      background: "rgba(255, 255, 255, 0.95)",
      backdropFilter: "blur(12px)",
      padding: "32px",
      marginBottom: "32px",
    }}
  >
    <h3
      style={{
        fontSize: "1.25rem",
        fontWeight: 700,
        color: "#1e293b",
        margin: "0 0 24px 0",
        display: "flex",
        alignItems: "center",
        gap: "8px",
      }}
    >
    Quarterly Totals
    </h3>

    <div
      style={{
        display: "grid",
        gridTemplateColumns: "repeat(auto-fit, minmax(250px, 1fr))",
        gap: "16px",
      }}
    >
      {Object.entries(filteredQuarterTotals)
        .sort(([a], [b]) => {
          const [quarterA, yearA] = a.split("-");
          const [quarterB, yearB] = b.split("-");
          if (yearA !== yearB) return parseInt(yearA) - parseInt(yearB);
          const qOrder = { JFM: 0, AMJ: 1, JAS: 2, OND: 3 };
          return qOrder[quarterA] - qOrder[quarterB];
        })
        .map(([key, totals]) => {
          const [quarter, year] = key.split("-");
          
          return (
            <QuarterPaymentCard
              key={key}
              quarter={quarter}
              year={year}
              totals={totals}
              showWithoutGST={showWithoutGST}
              paidQuarters={paidQuarters}
              updatePaidQuarters={updatePaidQuarters}
            />
          );
        })}
    </div>
  </div>
)}
            
      {/* Year-wise AMC Totals */}
{hasCalculations && Object.keys(filteredYearTotals).length > 0 && (
  <div
    style={{
      background: "rgba(255, 255, 255, 0.95)",
      backdropFilter: "blur(12px)",
      padding: "32px",
      marginBottom: "32px",
    }}
  >
    <h3
      style={{
        fontSize: "1.25rem",
        fontWeight: 700,
        color: "#1e293b",
        margin: "0 0 24px 0",
        display: "flex",
        alignItems: "center",
        gap: "8px",
      }}
    >
      Yearly Totals
    </h3>

    <div
      style={{
        display: "grid",
        gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))",
        gap: "16px",
      }}
    >
      {Object.entries(filteredYearTotals)
        .sort()
        .map(([year, totals]) => {
          // Calculate quarters for this year from rawResults
          const activeQuarters = rawResults.reduce((acc, product) => {
            if (product.quarters) {
              product.quarters.forEach(quarter => {
                if (quarter.year.toString() === year && quarter.totalAmount > 0) {
                  if (!acc.includes(quarter.quarter)) {
                    acc.push(quarter.quarter);
                  }
                }
              });
            }
            return acc;
          }, []);

          return (
            <div
              key={year}
              style={{
                background: "linear-gradient(135deg, #059669 0%, #0d9488 100%)",
                borderRadius: "16px",
                padding: "24px",
                color: "white",
                boxShadow: "0 10px 30px -10px rgba(5, 150, 105, 0.3)",
              }}
            >
              <div
                style={{
                  fontSize: "0.875rem",
                  fontWeight: 600,
                  marginBottom: "8px",
                  opacity: 0.9,
                }}
              >
                {year}
              </div>
              
              <div
                style={{
                  fontSize: "2rem",
                  fontWeight: 800,
                  marginBottom: "8px",
                }}
              >
                ₹{Math.round(totals.withGst).toLocaleString()}
              </div>
              
              {showWithoutGST && (
                <div
                  style={{
                    fontSize: "1rem",
                    marginBottom: "8px",
                    fontWeight: 600,
                    color: "rgba(250, 254, 129, 1)",
                    padding: "4px 8px",
                    borderRadius: "6px",
                    display: "inline-block",
                  }}
                >
                  Without GST: ₹ {Math.round(totals.withoutGst).toLocaleString()}
                </div>
              )}
              
              <div
                style={{
                  fontSize: "0.75rem",
                  opacity: 0.8,
                  borderTop: "1px solid rgba(255, 255, 255, 0.2)",
                  paddingTop: "8px",
                  marginTop: "8px",
                }}
              >
                {activeQuarters.length} active quarters
              </div>
            </div>
          );
        })}
    </div>
  </div>
)}          </div>
        )}
      </>
    )}
  </div>
)}

</div>

      {/* No Data Message */}
      {!hasExcelData && processedExcelData.length === 0 && (
        <div
          style={{
            ...cardStyle,
            textAlign: "center",
            backgroundColor: "#fefce8",
            border: "1px solid #fde047",
          }}
        >
          <div style={{ fontSize: "3rem", marginBottom: "16px" }}>📊</div>
          <h3
            style={{
              fontSize: "1.25rem",
              fontWeight: 700,
              color: "#92400e",
              margin: "0 0 8px 0",
            }}
          >
            No Excel Data Found
          </h3>
          <p
            style={{
              fontSize: "1rem",
              color: "#a16207",
              margin: "0 0 16px 0",
              lineHeight: 1.6,
            }}
          >
            Please go back to the dashboard and upload your Excel file first.{" "}
            <br />
            The AMC Calculator requires Excel data to perform calculations.
          </p>
          <button
            onClick={() => window.history.back()}
            style={{
              padding: "12px 24px",
              backgroundColor: "#ca8a04",
              color: "white",
              border: "none",
              borderRadius: "8px",
              fontSize: "0.9rem",
              fontWeight: 600,
              cursor: "pointer",
              transition: "all 0.2s ease",
            }}
          >
            ← Back to Dashboard
          </button>
        </div>
      )}
      {/* Payment Summary - Now with Real Status Tracking */}
{hasCalculations && Object.keys(quarterTotals).length > 0 && (
  <div style={{ marginBottom: "32px" }}>
    <div style={{
      background: "rgba(255, 255, 255, 0.95)",
      backdropFilter: "blur(12px)",
      borderRadius: "20px",
      padding: "32px",
      boxShadow: "0 20px 40px rgba(0, 0, 0, 0.1)",
      border: "1px solid rgba(255, 255, 255, 0.2)",
    }}>
      <h3 style={{
        fontSize: "1.25rem",
        fontWeight: 700,
        color: "#1e293b",
        marginBottom: "24px",
        display: "flex",
        alignItems: "center",
        gap: "8px"
      }}>
        💰 Final Payment Summary
      </h3>
      
      {(() => {
        // Calculate payment summary from actual payment status
        let totalPaid = 0;
        let totalBalance = 0;
        let paidCount = 0;
        let totalCount = Object.keys(quarterTotals).length;
        
        // Use actual payment status from paidQuarters state
        Object.entries(quarterTotals).forEach(([quarter, totals]) => {
          const quarterKey = `${quarter}`;  // quarter is already in "JFM-2024" format
          const isPaid = paidQuarters[quarterKey]?.paid || false;
          
          if (isPaid) {
            totalPaid += totals.withGst;
            paidCount++;
          } else {
            totalBalance += totals.withGst;
          }
        });
        
        return (
          <div style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))",
            gap: "16px"
          }}>
            {/* Total Amount */}
            <div style={{
              textAlign: "center",
              padding: "16px",
              backgroundColor: "#ffffff",
              borderRadius: "12px",
              border: "1px solid #e2e8f0"
            }}>
              <div style={{
                fontSize: "0.875rem",
                fontWeight: 600,
                color: "#64748b",
                marginBottom: "4px",
              }}>
                Total Amount
              </div>
              <div style={{
                fontSize: "1.25rem",
                fontWeight: 800,
                color: "#374151"
              }}>
                ₹{(totalPaid + totalBalance).toLocaleString()}
              </div>
            </div>
            
            {/* Paid Amount */}
            <div style={{
              textAlign: "center",
              padding: "16px",
              backgroundColor: "#f0fdf4",
              borderRadius: "12px",
              border: "1px solid #bbf7d0"
            }}>
              <div style={{
                fontSize: "0.875rem",
                fontWeight: 600,
                color: "#059669",
                marginBottom: "4px"
              }}>
                Paid Amount
              </div>
              <div style={{
                fontSize: "1.25rem",
                fontWeight: 800,
                color: "#059669"
              }}>
                ₹{totalPaid.toLocaleString()}
              </div>
              <div style={{
                fontSize: "0.75rem",
                color: "#059669",
                opacity: 0.8
              }}>
                {paidCount} of {totalCount} quarters
              </div>
            </div>
            
            {/* Balance Amount */}
            <div style={{
              textAlign: "center",
              padding: "16px",
              backgroundColor: "#fef3c7",
              borderRadius: "12px",
              border: "1px solid #fcd34d"
            }}>
              <div style={{
                fontSize: "0.875rem",
                fontWeight: 600,
                color: "#d97706",
                marginBottom: "4px"
              }}>
                Balance Amount
              </div>
              <div style={{
                fontSize: "1.25rem",
                fontWeight: 800,
                color: "#d97706"
              }}>
                ₹{totalBalance.toLocaleString()}
              </div>
              <div style={{
                fontSize: "0.75rem",
                color: "#d97706",
                opacity: 0.8
              }}>
                {totalCount - paidCount} quarters pending
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  </div>
)}

    </div>
  )};

export default AmcPaymentTracker;