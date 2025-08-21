import React, { useState, useEffect, useMemo, useCallback } from "react";
import { useDispatch, useSelector } from "react-redux";
import { useNavigate } from "react-router-dom";
import { Upload, Plus, FileText, ArrowLeft, Calculator, Package, CheckCircle, AlertCircle, Trash2, Download, FileSpreadsheet, Database } from "lucide-react";
import * as XLSX from "xlsx";
import VirtualDataTable from "../components/VirtualDataTable";
import { storeWarrantyCalculations } from "../store/slice/warrantyDataSlice";
import { selectExcelData, selectFileName, selectHasData, selectActiveSheet } from "../store/selectors/excelSelectors";
import { WarrantyExportManager } from '../utils/exportUtils';

const spinKeyframes = `
  @keyframes spin {
    0% { transform: rotate(0deg); }
    100% { transform: rotate(360deg); }
  }
`;
  const parseExcelDate = (dateValue) => {
    if (!dateValue) return new Date().toISOString().split("T")[0];
  
    const dateStr = String(dateValue).trim();
    
    if (dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
      return dateStr;
    }
  
    if (!isNaN(dateStr) && dateStr.length <= 6) {
      const serialNumber = parseInt(dateStr);
      if (serialNumber > 0 && serialNumber < 100000) {
        
        const excelEpoch = new Date(1899, 11, 30); 
        const resultDate = new Date(excelEpoch.getTime() + serialNumber * 24 * 60 * 60 * 1000);
        
        
        if (serialNumber > 59) {
          resultDate.setTime(resultDate.getTime() - 24 * 60 * 60 * 1000);
        }
        return resultDate.toISOString().split("T")[0];
      }
    }

    // Handling DD-MMM-YY format specifically for excel input
    const ddMmmYyMatch = dateStr.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2})$/);
    if (ddMmmYyMatch) {
      const [, day, monthStr, year] = ddMmmYyMatch;
      const monthMap = {
        'jan': 0, 'feb': 1, 'mar': 2, 'apr': 3,
        'may': 4, 'jun': 5, 'jul': 6, 'aug': 7,
        'sep': 8, 'oct': 9, 'nov': 10, 'dec': 11
      };
      const monthIndex = monthMap[monthStr.toLowerCase()];
      if (monthIndex !== undefined) {
        const fullYear = 2000 + parseInt(year);
        const date = new Date(Date.UTC(fullYear, monthIndex, parseInt(day)));
        const formattedDate = date.toISOString().split("T")[0];
        return formattedDate;
      }
    }
    
    try {
      const parsedDate = new Date(dateStr);
      if (!isNaN(parsedDate.getTime())) {
        const formattedDate = parsedDate.toISOString().split("T")[0];
        return formattedDate;
      }
    } catch (error) {
      console.warn(`⚠️ Warranty: Failed to parse date: ${dateStr}`);
    }

    const today = new Date().toISOString().split("T")[0];
    return today;
  };
  
  if (typeof document !== "undefined") {
    const style = document.createElement("style");
    style.textContent = spinKeyframes;
    document.head.appendChild(style);
  }

 const WarrantyPaymentTracker = () => {
  const dispatch = useDispatch();
  const navigate = useNavigate();

  // Redux state
  const excelData = useSelector(selectExcelData);
  const fileName = useSelector(selectFileName);
  const hasExcelData = useSelector(selectHasData);
  const activeSheet = useSelector(selectActiveSheet);

  const [showDataPreview, setShowDataPreview] = useState(false);
  const [showCacheManager, setShowCacheManager] = useState(false);
  const [useCache] = useState(true);
  const [selectedLocation, setSelectedLocation] = useState("All Locations");
  const [selectedStatus, setSelectedStatus] = React.useState("All");
  const [selectedYear, setSelectedYear] = React.useState("All");
  const [showWithoutGST, setShowWithoutGST] = useState(false);
  const [paidQuarters, setPaidQuarters] = useState(() => {
    const saved = localStorage.getItem('paidQuarters');
    return saved ? JSON.parse(saved) : {};
  });

  const updatePaidQuarters = useCallback((quarterKey, isPaid, date = null) => {
  const newPaidQuarters = { ...paidQuarters };
  if (isPaid) {
    newPaidQuarters[quarterKey] = {
      paid: true,
      date: date || new Date().toISOString().split("T")[0]
    };
  } else {
    delete newPaidQuarters[quarterKey];
  }
  setPaidQuarters(newPaidQuarters);
  localStorage.setItem('paidQuarters', JSON.stringify(newPaidQuarters));
}, [paidQuarters]);

  // Local state
  const [isCalculating, setIsCalculating] = useState(false);
  const [currentTime, setCurrentTime] = useState("");
  const [activeTab, setActiveTab] = useState("upload"); // 'upload', 'manual', 'schedule'
  const [viewMode, setViewMode] = useState("cards"); // 'table', 'cards'
  const [location, setLocation] = useState("");
  const [warrantyProducts, setWarrantyProducts] = useState([]);
  const [manualProduct, setManualProduct] = useState({
    itemName: "",
    invoiceValue: "",
    invoiceNumber: "",
    quantity: 1,
    uatDate: new Date().toISOString().split("T")[0],
    warrantyStart: "",
    warrantyYears: 3,
    warrantyStartSameAsUAT: true,
  });
  const [calculatedSchedule, setCalculatedSchedule] = useState([]);
  const [showGST, setShowGST] = useState(true);

  // Quarter date logic 
  const getQuarterDates = useCallback((year) => {
    return {
      OND: [new Date(year, 9, 5), new Date(year + 1, 0, 4)], // Oct 5 - Jan 4
      JFM: [new Date(year + 1, 0, 5), new Date(year + 1, 3, 4)], // Jan 5 - Apr 4
      AMJ: [new Date(year + 1, 3, 5), new Date(year + 1, 6, 4)], // Apr 5 - Jul 4
      JAS: [new Date(year + 1, 6, 5), new Date(year + 1, 9, 4)], // Jul 5 - Oct 4
    };
  }, []);

  const calculateWarrantySchedule = useCallback(
    (startDate, invoiceValue, gstRate = 0.18, warrantyPercent = 0.15, years = 3) => {
      const schedule = new Map();
      const splitDetails = new Map();

      const totalWarrantyAmount = invoiceValue * warrantyPercent;
      const quarterlyAmount = totalWarrantyAmount / (years * 4); 

      const warrantyEnd = new Date(startDate);
      warrantyEnd.setFullYear(warrantyEnd.getFullYear() + years);
      warrantyEnd.setDate(warrantyEnd.getDate() - 1);

      let firstQuarterActualAmount = 0;
      let firstQuarterKey = null;
      let lastQuarterKey = null;

      // Find all quarters that overlap with warranty period
      const quarterlyPayments = [];

      for (
        let year = startDate.getFullYear() - 1;
        year <= warrantyEnd.getFullYear() + 1;
        year++
      ) {
        const quarters = getQuarterDates(year);

        for (const [qtrName, [qStart, qEnd]] of Object.entries(quarters)) {
          if (qEnd < startDate || qStart > warrantyEnd) {
            continue;
          }

          const overlapStart = new Date(
            Math.max(startDate.getTime(), qStart.getTime())
          );
          const overlapEnd = new Date(
            Math.min(warrantyEnd.getTime(), qEnd.getTime())
          );

          if (overlapStart > overlapEnd) {
            continue;
          }

          const displayYear = qStart.getFullYear();
          const key = `${displayYear}-${qtrName}`;

          const totalDaysInQuarter =
            Math.floor((qEnd - qStart) / (1000 * 60 * 60 * 24)) + 1;
          const overlapDays =
            Math.floor((overlapEnd - overlapStart) / (1000 * 60 * 60 * 24)) + 1;

          quarterlyPayments.push({
            key,
            quarter: qtrName,
            quarterStart: qStart,
            quarterEnd: qEnd,
            overlapStart,
            overlapEnd,
            totalDaysInQuarter,
            overlapDays,
            isFirst: overlapStart.getTime() === startDate.getTime(),
            isLast: overlapEnd.getTime() === warrantyEnd.getTime(),
          });
        }
      }

      // Sort quarters chronologically
      quarterlyPayments.sort((a, b) => a.quarterStart - b.quarterStart);
      quarterlyPayments.forEach((quarter, index) => {
        let proratedAmount;

        if (quarter.isFirst) {
          // First quarter: Prorate based on actual days in warranty period
          proratedAmount = (quarter.overlapDays / quarter.totalDaysInQuarter) * quarterlyAmount;
          firstQuarterActualAmount = proratedAmount;
          firstQuarterKey = quarter.key;

        } else if (quarter.isLast) {
          // Last quarter: Pay ONLY the deficit from first quarter (quarterly_amount - first_quarter_amount)
          const firstQuarterDeficit = quarterlyAmount - firstQuarterActualAmount;
          proratedAmount = firstQuarterDeficit;
          lastQuarterKey = quarter.key;
          
        } else {
          // Middle quarters: Full quarterly amount if fully covered
          if (quarter.overlapDays === quarter.totalDaysInQuarter) {
            proratedAmount = quarterlyAmount;
          } else {
            // Partial middle quarter (rare case)
            proratedAmount = (quarter.overlapDays / quarter.totalDaysInQuarter) * quarterlyAmount;
          }
        
        }

        const withoutGst = Math.round(proratedAmount * 100) / 100;
        const withGst = Math.round(withoutGst * (1 + gstRate) * 100) / 100;

        // Store in schedule
        if (schedule.has(quarter.key)) {
          const existing = schedule.get(quarter.key);
          schedule.set(quarter.key, [
            existing[0] + withGst,
            existing[1] + withoutGst,
          ]);
        } else {
          schedule.set(quarter.key, [withGst, withoutGst]);
        }

        // Store detailed breakdown
        splitDetails.set(quarter.key, {
          quarter: quarter.quarter,
          quarterStart: quarter.quarterStart,
          quarterEnd: quarter.quarterEnd,
          proratedDays: quarter.overlapDays,
          totalDaysInQuarter: quarter.totalDaysInQuarter,
          amountWithoutGST: withoutGst,
          amountWithGST: withGst,
          calculationType: quarter.isFirst 
            ? "First Quarter (Prorated)" 
            : quarter.isLast 
            ? "Last Quarter (Deficit Only)" 
            : quarter.overlapDays === quarter.totalDaysInQuarter
            ? "Full Quarter"
            : "Partial Quarter",
          isFirst: quarter.isFirst,
          isLast: quarter.isLast,
        });
      });

      // Verify total calculation
      const calculatedTotal = Array.from(schedule.values()).reduce(
        (sum, [withGst, withoutGst]) => sum + withoutGst, 
        0
      );
      const expectedTotal = totalWarrantyAmount;
      const difference = Math.abs(expectedTotal - calculatedTotal);
 
      return {
        schedule: Object.fromEntries(schedule),
        splitDetails: Object.fromEntries(splitDetails),
        metadata: {
          totalExpected: expectedTotal,
          totalCalculated: calculatedTotal,
          difference: difference,
          quarterCount: quarterlyPayments.length,
          firstQuarterAdjustment: quarterlyAmount - firstQuarterActualAmount
        }
      };
    },
    [getQuarterDates]
  );

  const processExcelData = useCallback(() => {
    if (!excelData || Object.keys(excelData).length === 0) {
      alert("No Excel data available. Please upload an Excel file first.");
      return;
    }

    const sheetName = activeSheet || Object.keys(excelData)[0];
    const sheetData = excelData[sheetName];

    if (!Array.isArray(sheetData) || sheetData.length === 0) {
      alert(
        `No data found in sheet "${sheetName}". Please check your Excel file.`
      );
      return;
    }

    setIsCalculating(true);

    try {
      const processedProducts = sheetData.map((row, index) => {
     
        const itemName =
          row["Item Name"] ||
          row["itemName"] ||
          row["Product Name"] ||
          row["productName"] ||
          `Product ${index + 1}`;
        const invoiceValue = (() => {
          let costStr = String(
            row["Cost"] ||
              row["cost"] ||
              row["Price"] ||
              row["price"] ||
              row["Invoice Value"] ||
              row["invoiceValue"] ||
              0
          );
          const originalStr = costStr.toLowerCase();
          const isCrores =
            originalStr.includes("cr") || originalStr.includes("crore");
          const isLakhs =
            originalStr.includes("lakh") || originalStr.includes("lac");

          costStr = costStr
            .replace(/[₹$,\s]/g, "") 
            .replace(/\.00$/, "") // 
            .replace(/[Cc][Rr].*$/, "") 
            .replace(/[Ll]akh.*$/, "") 
            .trim();

          let parsed = parseFloat(costStr || 0);

          if (isCrores) {
            parsed = parsed * 10000000; 
            
          } else if (isLakhs) {
            parsed = parsed * 100000;
          }

          return parsed;
        })();

        const invoiceNumber = String(
            row["Invoice Number"] || row["invoiceNumber"] || row["Invoice No"] || row["Invoice No."] || ''
        )
        const quantity = parseInt(row["Quantity"] || row["quantity"] || row["Qty"] ||row["qty"] ||  1);
     
        const uatDate = parseExcelDate(
            row["UAT Date"] ||
            row["UAT DATE"] ||
            row["uatDate"] ||
            row["Purchase Date"] ||
            row["purchaseDate"]
        );
        
        const warrantyStart = row["Warranty Start"]
          ? parseExcelDate(row["Warranty Start"] || row["WARRANTY START"] )
          : uatDate;
          
        const warrantyYears = parseInt(
          row["Warranty Years"] || row["warrantyYears"] || 3
        );

        return {
          id: `excel-${index}`,
          itemName,
          invoiceValue,
          invoiceNumber,
          quantity,
          uatDate,  
          warrantyStart,  
          warrantyYears,
          location: location || "Default Location",
          source: "excel",
        };
      });

      setWarrantyProducts(processedProducts);
      setActiveTab("schedule");

      setTimeout(() => {
        setIsCalculating(false);
      }, 1000);
    } catch (error) {
      alert("Error processing Excel data. Please check the file format.");
      setIsCalculating(false);
    }
  }, [excelData, activeSheet, location]);

  // Add manual product
  const addManualProduct = useCallback(() => {
    if (!manualProduct.itemName || !manualProduct.invoiceValue) {
      alert("Please fill in required fields: Item Name and Cost");
      return;
    }

    const warrantyStartDate = manualProduct.warrantyStartSameAsUAT
      ? manualProduct.uatDate
      : manualProduct.warrantyStart || manualProduct.uatDate;

    const processedCost = (() => {
      let costStr = String(manualProduct.invoiceValue || 0);
      const originalStr = costStr.toLowerCase();
      const isCrores =
        originalStr.includes("cr") || originalStr.includes("crore");
      const isLakhs =
        originalStr.includes("lakh") || originalStr.includes("lac");

      
      costStr = costStr
        .replace(/[₹$,\s]/g, "") 
        .replace(/\.00$/, "") 
        .replace(/[Cc][Rr].*$/, "") 
        .replace(/[Ll]akh.*$/, "") 
        .trim();

      let parsed = parseFloat(costStr || 0);

      
      if (isCrores) {
        parsed = parsed * 10000000; 
      } else if (isLakhs) {
        parsed = parsed * 100000;
      }

      return parsed;
    })();

    const newProduct = {
      id: `manual-${Date.now()}`,
      itemName: manualProduct.itemName,
      invoiceValue: processedCost,
      invoiceNumber: manualProduct.invoiceNumber,
      quantity: parseInt(manualProduct.quantity),
      uatDate: manualProduct.uatDate,
      warrantyStart: warrantyStartDate,
      warrantyYears: parseInt(manualProduct.warrantyYears),
      location: location || "Default Location",
      source: "manual",
    };

    setWarrantyProducts((prev) => [...prev, newProduct]);

    setManualProduct({
      itemName: "",
      invoiceValue: "",
      invoiceNumber: "",
      quantity: 1,
      uatDate: new Date().toISOString().split("T")[0],
      warrantyStart: "",
      warrantyYears: 3,
      warrantyStartSameAsUAT: true,
    });
  }, [manualProduct, location]);

    const warrantyTableColumns = useMemo(() => {
   
    const baseColumns = [
      {
        key: "itemName",
        title: "Item Name",
        width: 200,
        filterable: true,
      },
      {
        key: "uatDate",
        title: "UAT Date",
        width: 120,
      },
      {
        key: "warrantyStart",
        title: "Warranty Start",
        width: 120,
      },
      {
        key: "invoiceValue",
        title: "Invoice Value",
        width: 120,
        className: "text-right",
      },
      {
        key: "invoiceNumber",
        title: "Invoice Number",
        width: 180,
        className: "text-left",
      },
      {
        key: "quantity",
        title: "Quantity",
        width: 80,
        className: "text-center",
      },
      {
        key: "location",
        title: "Location",
        width: 120,
        filterable: true,
      },
    ];

    const quarterColumns = [];
    const quarterSet = new Set();

    if (calculatedSchedule && calculatedSchedule.length > 0) {
      calculatedSchedule.forEach((product) => {
        Object.keys(product).forEach((key) => {
          if (key.match(/^(JFM|AMJ|JAS|OND) \d{4}$/)) {
            quarterSet.add(key);
          }
        });
      });
    }
    const quarterOrder = ["JFM", "AMJ", "JAS", "OND"];
    const sortedQuarters = Array.from(quarterSet).sort((a, b) => {
      const [qA, yearA] = a.split(" ");
      const [qB, yearB] = b.split(" ");
   
      const yearNumA = parseInt(yearA);
      const yearNumB = parseInt(yearB);

      if (yearNumA !== yearNumB) {
        return yearNumA - yearNumB;
      }
   
      return quarterOrder.indexOf(qA) - quarterOrder.indexOf(qB);
    });

    sortedQuarters.forEach((quarterKey) => {
      quarterColumns.push({
        key: quarterKey,
        title: quarterKey,
        width: 120,
        className: "text-right",
      });
    });

    const totalColumns = [];
    if (calculatedSchedule && calculatedSchedule.length > 0) {
      const firstRow = calculatedSchedule[0];
      if (firstRow) {
        Object.keys(firstRow).forEach((key) => {
          if (key.includes("Total (") && key.includes(" Years)")) {
            totalColumns.push({
              key: key,
              title: key,
              width: 140,
              className: "text-right",
            });
          }
        });
      }
    }
    return [...baseColumns, ...quarterColumns, ...totalColumns];
  }, [calculatedSchedule]);

  const warrantyFormatters = useMemo(() => {
    const formatters = {
      invoiceValue: (value) => (value ? `₹${value.toLocaleString()}` : "₹0"),
      invoiceNumber: (value) => value || "-",
      quantity: (value) => value?.toLocaleString() || "0",
      uatDate: (value) => value || "-",
      warrantyStart: (value) => value || "-",
      location: (value) => value || "-",
    };

    if (calculatedSchedule && calculatedSchedule.length > 0) {
      const firstRow = calculatedSchedule[0];
      if (firstRow) {
        Object.keys(firstRow).forEach((key) => {
          if (key.match(/^(JFM|AMJ|JAS|OND) \d{4}$/)) {
            formatters[key] = (value) =>
              value ? `₹${value.toLocaleString()}` : "₹0";
          }
          if (key.includes("Total (") && key.includes(" Years)")) {
            formatters[key] = (value) =>
              value ? `₹${value.toLocaleString()}` : "₹0";
          }
        });
      }
    }

    return formatters;
  }, [calculatedSchedule]);

  const warrantySummary = useMemo(() => {
    if (!calculatedSchedule || calculatedSchedule.length === 0) {
      return {
        totalProducts: 0,
        totalValue: 0,
        successful: 0,
        errors: 0,
      };
    }

    const grandTotalRow = calculatedSchedule.find(
      (row) => row.itemName === "Grand Total"
    );
    const productRows = calculatedSchedule.filter(
      (row) => row.itemName !== "Grand Total"
    );

    let totalValue = 0;
    if (grandTotalRow) {
      const totalColumn = Object.keys(grandTotalRow).find(
        (key) => key.includes("Total (") && key.includes(" Years)")
      );
      totalValue = totalColumn ? grandTotalRow[totalColumn] : 0;
    }

    return {
      totalProducts: productRows.length,
      totalValue: totalValue,
      successful: productRows.length,
      errors: 0,
    };
  }, [calculatedSchedule]);

     const calculateQuarterlySchedule = useCallback(() => {
      if (warrantyProducts.length === 0) {
        alert("No warranty products to calculate. Please add products first.");
        return;
      }
      setIsCalculating(true);
      try {
        const quarterlyRows = [];
        const allQuarters = new Set();
        
        warrantyProducts.forEach((product) => {
          const warrantyStart = new Date(product.warrantyStart);
          const totalCost = product.invoiceValue;

          const { schedule } = calculateWarrantySchedule(
            warrantyStart,
            totalCost,
            0.18, 
            0.15, 
            product.warrantyYears
          );

          const row = {
            itemName: product.itemName,
            uatDate: new Date(product.uatDate).toLocaleDateString("en-GB"),
            warrantyStart: warrantyStart.toLocaleDateString("en-GB"),
            invoiceValue: product.invoiceValue,
            invoiceNumber: product.invoiceNumber || "",
            quantity: product.quantity,
            location: product.location,
            source: product.source,
          };

          let totalAmount = 0;

          Object.entries(schedule).forEach(([key, [withGst, withoutGst]]) => {
            const [year, quarter] = key.split("-");
            const colName = `${quarter} ${year}`;
            const value = showGST ? withGst : withoutGst;
            row[colName] = Math.round(value * 100) / 100;
            totalAmount += showGST ? withGst : withoutGst;
            allQuarters.add(colName);
          });

          const finalTotal = Math.round(totalAmount * 100) / 100;
          row[`Total (${product.warrantyYears} Years)`] = finalTotal;
          
          quarterlyRows.push(row);
        });

        const quarterOrder = ["JFM", "AMJ", "JAS", "OND"];
        const sortedQuarters = Array.from(allQuarters).sort((a, b) => {
          const [qA, yearA] = a.split(" ");
          const [qB, yearB] = b.split(" ");
   
          const yearNumA = parseInt(yearA);
          const yearNumB = parseInt(yearB);
 
          if (yearNumA !== yearNumB) {
            return yearNumA - yearNumB;
          }

          return quarterOrder.indexOf(qA) - quarterOrder.indexOf(qB);
        });

        quarterlyRows.forEach((row) => {
          sortedQuarters.forEach((quarter) => {
            if (!(quarter in row)) {
              row[quarter] = 0.0;
            }
          });
        });

        const totals = {
          itemName: "Grand Total",
          uatDate: "",
          warrantyStart: "",
          invoiceValue: "",
          quantity: "",
          location: "",
          source: "",
        };

        sortedQuarters.forEach((quarter) => {
          totals[quarter] = quarterlyRows.reduce(
            (sum, row) => sum + (row[quarter] || 0),
            0
          );
          totals[quarter] = Math.round(totals[quarter] * 100) / 100;
        });

        // Calculate grand total for all warranty years
        let grandTotal = 0;
        quarterlyRows.forEach((row) => {
          const totalColumns = Object.keys(row).filter(
            (key) => key.includes("Total (") && key.includes(" Years)")
          );
          totalColumns.forEach((totalCol) => {
            grandTotal += row[totalCol] || 0;
          });
        });

        // Add total columns for each unique warranty period
        const uniqueWarrantyPeriods = [...new Set(warrantyProducts.map(p => p.warrantyYears))];
        uniqueWarrantyPeriods.forEach((years) => {
          const totalColName = `Total (${years} Years)`;
          if (!totals[totalColName]) {
            totals[totalColName] = quarterlyRows
              .filter((row) => row[totalColName] !== undefined)
              .reduce((sum, row) => sum + (row[totalColName] || 0), 0);
            totals[totalColName] = Math.round(totals[totalColName] * 100) / 100;
          }
        });

        // If there's only one warranty period, use that for the main total
        if (uniqueWarrantyPeriods.length === 1) {
          const mainTotalCol = `Total (${uniqueWarrantyPeriods[0]} Years)`;
          totals[mainTotalCol] = Math.round(grandTotal * 100) / 100;
        }

        const finalSchedule = [...quarterlyRows, totals];
        setCalculatedSchedule(finalSchedule);

        // Store in Redux for payment tracker
        dispatch(
          storeWarrantyCalculations({
            calculations: quarterlyRows,
            metadata: {
              totalProducts: warrantyProducts.length,
              totalValue: grandTotal,
              calculatedAt: new Date().toISOString(),
              location: location,
              quarters: sortedQuarters,
              showGST: showGST,
            },
          })
        );

        setActiveTab("schedule");
      } catch (error) {
        alert("Error calculating warranty schedule. Please check your data.");
      } finally {
        setIsCalculating(false);
      }
    }, [
      warrantyProducts,
      showGST,
      calculateWarrantySchedule,
      location,
      dispatch,
    ]);

   // Export to Excel (fixed to match UI column ordering)
    const exportToExcel = useCallback(() => {
      if (!calculatedSchedule || calculatedSchedule.length === 0) {
        alert("No data to export. Please calculate the warranty schedule first.");
        return;
      }

      try {
        // Collect all unique quarters from the UI display format (QTR YYYY)
        const quarterSet = new Set();
        const totalColumns = new Set();
        
        calculatedSchedule.forEach((row) => {
          Object.keys(row).forEach((key) => {
            // Look for the UI display format: "JFM 2021", "AMJ 2021", etc.
            if (key.match(/^(JFM|AMJ|JAS|OND) \d{4}$/)) {
              quarterSet.add(key);
            } else if (key.includes("Total (") && key.includes(" Years)")) {
              totalColumns.add(key);
            }
          });
        });

        // Sort quarters chronologically 
        const quarterOrder = ["JFM", "AMJ", "JAS", "OND"];
        const sortedQuarters = Array.from(quarterSet).sort((a, b) => {
          const [qA, yearA] = a.split(" ");
          const [qB, yearB] = b.split(" ");
          
          const yearNumA = parseInt(yearA);
          const yearNumB = parseInt(yearB);
          
          // First sort by year
          if (yearNumA !== yearNumB) {
            return yearNumA - yearNumB;
          }
  
          // Then sort by quarter within the same year
          return quarterOrder.indexOf(qA) - quarterOrder.indexOf(qB);
        });

        // Define the correct column order for export
        const baseColumns = ["Item Name", "UAT Date", "Warranty Start", "Invoice Value", "Quantity", "Location"];
        const orderedColumns = [...baseColumns, ...sortedQuarters, ...Array.from(totalColumns)];

        // Transform data with proper column ordering
        const exportData = calculatedSchedule.map((row) => {
          const transformedRow = {};
          
          // Copy basic fields with proper headers in correct order
          transformedRow["Item Name"] = row.itemName;
          transformedRow["UAT Date"] = row.uatDate;
          transformedRow["Warranty Start"] = row.warrantyStart;
          transformedRow["Invoice Value"] = row.invoiceValue ? `₹${row.invoiceValue.toLocaleString()}` : "";
          transformedRow["Invoice Number"] = row.invoiceNumber || ""; 
          transformedRow["Quantity"] = row.quantity || "";
          transformedRow["Location"] = row.location || "";
          
          // Add quarter columns in chronological order using UI display format
          sortedQuarters.forEach((quarterDisplay) => {
            // quarterDisplay is already in "QTR YYYY" format from the UI
            const value = row[quarterDisplay];
            transformedRow[quarterDisplay] = value ? `₹${value.toLocaleString()}` : "₹0";
          });
          
          // Add total columns
          Array.from(totalColumns).forEach((totalCol) => {
            const value = row[totalCol];
            transformedRow[totalCol] = value ? `₹${value.toLocaleString()}` : "₹0";
          });
          
          return transformedRow;
        });

        // Create worksheet with ordered columns
        const ws = XLSX.utils.json_to_sheet(exportData, { header: orderedColumns });
        
        // Auto-resize columns for better readability
        const colWidths = [];
        orderedColumns.forEach((header, index) => {
          let maxWidth = header.length;
          exportData.forEach((row) => {
            const cellValue = String(row[header] || "");
            maxWidth = Math.max(maxWidth, cellValue.length);
          });
          // Set reasonable min/max widths
          colWidths[index] = { wch: Math.min(Math.max(maxWidth + 2, 10), 25) };
        });
        
        ws['!cols'] = colWidths;

        // Create workbook and add worksheet
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Warranty Schedule");

        // Generate filename with current settings
        const fileName = `Warranty_Schedule_${
          showGST ? "With_GST" : "Without_GST"
        }_${new Date().toISOString().split("T")[0]}.xlsx`;
        
        // Write file
        XLSX.writeFile(wb, fileName);
      } catch (error) {
        alert("Error exporting to Excel. Please try again.");
      }
    }, [calculatedSchedule, showGST]);

  // Warranty Export Component
const WarrantyExportControls = () => {
  const [showExportOptions, setShowExportOptions] = useState(false);

  const handleExport = useCallback(async (type) => {
    if (!calculatedSchedule || calculatedSchedule.length === 0) {
      alert('Please calculate warranty schedule first before exporting.');
      return;
    }

    try {
      const exportManager = new WarrantyExportManager({
        products: warrantyProducts,
        schedule: calculatedSchedule,
        paidQuarters,
        settings: { gstRate: 0.18 }
      });

      const baseFilename = fileName ? fileName.replace(/\.[^/.]+$/, '') : 'Warranty_Schedule';

      switch (type) {
        case 'excel':
          await exportManager.exportToExcel(baseFilename);
          break;
        case 'csv':
          exportManager.exportToCSV(baseFilename);
          break;
        case 'pdf':
          exportManager.exportToPDF(`${baseFilename}_Report`);
          break;
        case 'json':
          exportManager.exportToJSON(`${baseFilename}_Data`);
          break;
        default:
          break;
      }

      setShowExportOptions(false);
    } catch (error) {
      alert(`Error exporting to ${type.toUpperCase()}. Please try again.`);
      console.error('Export error:', error);
    }
  }, [warrantyProducts, calculatedSchedule, paidQuarters, fileName]);

  const hasCalculations = calculatedSchedule && calculatedSchedule.length > 0;

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


  // Remove product
  const removeProduct = useCallback((productId) => {
    setWarrantyProducts((prev) => prev.filter((p) => p.id !== productId));
  }, []);

  // Styling 
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
    if (!excelData || Object.keys(excelData).length === 0) return [];
    const sheetName = activeSheet || Object.keys(excelData)[0];
    const sheetData = excelData[sheetName];
    if (!Array.isArray(sheetData) || sheetData.length === 0) return [];
    return sheetData.map((row, index) => ({
      productName: row["Item Name"] || row["itemName"] || row["Product Name"] || `Product ${index + 1}`,
      invoiceValue: (() => {
        let costStr = String(row["Cost"] || row["cost"] || row["Price"] || row["Invoice Value"] || 0);
        const originalStr = costStr.toLowerCase();
        const isCrores = originalStr.includes("cr") || originalStr.includes("crore");
        const isLakhs = originalStr.includes("lakh") || originalStr.includes("lac");
        costStr = costStr.replace(/[₹$,\s]/g, "").replace(/\.00$/, "").replace(/[Cc][Rr].*$/, "").replace(/[Ll]akh.*$/, "").trim();
        let parsed = parseFloat(costStr || 0);
        if (isCrores) parsed = parsed * 10000000;
        else if (isLakhs) parsed = parsed * 100000;
        return parsed;
      })(),
      invoiceNumber: String(row["Invoice Number"] || row["invoiceNumber"] || row["Invoice No"] || row["Invoice No."] || ''),
      quantity: parseInt(row["Quantity"] || row["quantity"] || row["Qty"] || 1),
      location: row["Location"] || row["location"] || "Default Location",
      uatDate: parseExcelDate(row["UAT Date"] || row["uatDate"] || row["Purchase Date"])
    }));
  }, [excelData, activeSheet]);

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
              Warranty Payment Tracker
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
                        ₹{item.invoiceValue.toLocaleString()}
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
                <strong>💡 Data Validation:</strong>
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
      </div>)}

      {/* Tabs */}
      <div style={cardStyle}>
        <div
          style={{
            display: "flex",
            gap: "4px",
            marginBottom: "24px",
            padding: "4px",
            backgroundColor: "#f1f5f9",
            borderRadius: "8px",
          }}
        >
          <button
            onClick={() => setActiveTab("upload")}
            style={styles.tab(activeTab === "upload")}
          >
            <Upload size={16} style={{ marginRight: "6px" }} />
            Process Excel Data
          </button>
          <button
            onClick={() => setActiveTab("manual")}
            style={styles.tab(activeTab === "manual")}
          >
            <Plus size={16} style={{ marginRight: "6px" }} />
            Manual Entry
          </button>
        </div>

        {/* Excel Upload Tab */}
        {activeTab === "upload" && (
          <div>
            <h3
              style={{
                fontSize: "1.25rem",
                fontWeight: 700,
                marginBottom: "16px",
                color: "#374151",
              }}
            >
              Process Excel Data
            </h3>
            <p
              style={{
                color: "#64748b",
                marginBottom: "24px",
                lineHeight: 1.6,
              }}
            >
              Use your uploaded Excel file to automatically create warranty
              products.
            </p>

            {hasExcelData ? (
              <div
                style={{
                  display: "flex",
                  flexDirection: "column",
                  gap: "16px",
                }}
              >
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: "16px",
                    padding: "16px",
                    backgroundColor: "#f0fdf4",
                    border: "1px solid #bbf7d0",
                    borderRadius: "12px",
                  }}
                >
                  <CheckCircle size={24} style={{ color: "#059669" }} />
                  <div>
                    <div style={{ color: "#374151", fontWeight: 600 }}>
                      Excel file loaded: {fileName}
                    </div>
                    <div style={{ color: "#64748b", fontSize: "0.875rem" }}>
                      {(() => {
                        if (!excelData || Object.keys(excelData).length === 0)
                          return 0;
                        const sheetName =
                          activeSheet || Object.keys(excelData)[0];
                        const sheetData = excelData[sheetName];
                        return Array.isArray(sheetData) ? sheetData.length : 0;
                      })()}{" "}
                      rows available
                    </div>
                  </div>
                </div>

                <button
                  onClick={processExcelData}
                  disabled={isCalculating}
                  style={{
                    ...styles.button,
                    ...styles.primaryButton,
                    width: "100%",
                    justifyContent: "center",
                  }}
                >
                  {isCalculating ? (
                    <>
                      <div
                        style={{
                          width: "20px",
                          height: "20px",
                          border: "2px solid transparent",
                          borderTop: "2px solid white",
                          borderRadius: "50%",
                          animation: "spin 1s linear infinite",
                        }}
                      ></div>
                      Processing...
                    </>
                  ) : (
                    <>
                      <FileText size={20} />
                      Process Excel Data for Warranty
                    </>
                  )}
                </button>
              </div>
            ) : (
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: "16px",
                  padding: "16px",
                  backgroundColor: "#fef3c7",
                  border: "1px solid #fcd34d",
                  borderRadius: "12px",
                }}
              >
                <AlertCircle size={24} style={{ color: "#d97706" }} />
                <div>
                  <div style={{ color: "#374151", fontWeight: 600 }}>
                    No Excel file available
                  </div>
                  <div style={{ color: "#64748b", fontSize: "0.875rem" }}>
                    Please upload an Excel file from the Dashboard first
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Manual Entry Tab */}
        {activeTab === "manual" && (
          <div>
            <h3
              style={{
                fontSize: "1.25rem",
                fontWeight: 700,
                marginBottom: "16px",
                color: "#374151",
              }}
            >
              Add Warranty Product
            </h3>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(auto-fit, minmax(250px, 1fr))",
                gap: "16px",
                marginBottom: "24px",
              }}
            >
              <div>
                <label style={styles.label}>Item Name *</label>
                <input
                  type="text"
                  value={manualProduct.itemName}
                  onChange={(e) =>
                    setManualProduct((prev) => ({
                      ...prev,
                      itemName: e.target.value,
                    }))
                  }
                  placeholder="Enter item name"
                  style={styles.input}
                />
              </div>
              <div>
                <label style={styles.label}>Invoice Value (₹) *</label>
                <input
                  type="text"
                  value={manualProduct.invoiceValue}
                  onChange={(e) =>
                    setManualProduct((prev) => ({
                      ...prev,
                      invoiceValue: e.target.value,
                    }))
                  }
                  placeholder="e.g., 2.5 Cr, 50 Lakhs, or 2500000"
                  style={styles.input}
                />
              </div>
              <div>
                <label style={styles.label}>Invoice Number</label>
  <input
    type="text"
    value={manualProduct.invoiceNumber}
    onChange={(e) =>
      setManualProduct((prev) => ({
        ...prev,
        invoiceNumber: e.target.value,
      }))
    }
    placeholder="Enter invoice number"
    style={styles.input}
  />
</div>
              <div>
                <label style={styles.label}>Quantity</label>
                <input
                  type="number"
                  value={manualProduct.quantity}
                  onChange={(e) =>
                    setManualProduct((prev) => ({
                      ...prev,
                      quantity: parseInt(e.target.value) || 1,
                    }))
                  }
                  min="1"
                  style={styles.input}
                />
              </div>
              <div>
                <label style={styles.label}>UAT Date</label>
                <input
                  type="date"
                  value={manualProduct.uatDate}
                  onChange={(e) =>
                    setManualProduct((prev) => ({
                      ...prev,
                      uatDate: e.target.value,
                    }))
                  }
                  style={styles.input}
                />
              </div>
              <div>
                <label style={styles.label}>Warranty Duration (Years)</label>
                <select
                  value={manualProduct.warrantyYears}
                  onChange={(e) =>
                    setManualProduct((prev) => ({
                      ...prev,
                      warrantyYears: parseInt(e.target.value),
                    }))
                  }
                  style={styles.input}
                >
                  {[1, 2, 3, 4, 5].map((year) => (
                    <option key={year} value={year}>
                      {year} Year{year > 1 ? "s" : ""}
                    </option>
                  ))}
                </select>
              </div>
              <div style={{ gridColumn: "1 / -1" }}>
                <label
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: "8px",
                    color: "#374151",
                    marginBottom: "16px",
                  }}
                >
                  <input
                    type="checkbox"
                    checked={manualProduct.warrantyStartSameAsUAT}
                    onChange={(e) =>
                      setManualProduct((prev) => ({
                        ...prev,
                        warrantyStartSameAsUAT: e.target.checked,
                      }))
                    }
                    style={{ accentColor: "#3b82f6" }}
                  />
                  Warranty starts on UAT date
                </label>

                {!manualProduct.warrantyStartSameAsUAT && (
                  <div>
                    <label style={styles.label}>Warranty Start Date</label>
                    <input
                      type="date"
                      value={manualProduct.warrantyStart}
                      onChange={(e) =>
                        setManualProduct((prev) => ({
                          ...prev,
                          warrantyStart: e.target.value,
                        }))
                      }
                      style={styles.input}
                    />
                  </div>
                )}
              </div>
            </div>

            <button
              onClick={addManualProduct}
              style={{
                ...styles.button,
                ...styles.primaryButton,
                width: "100%",
                justifyContent: "center",
              }}
            >
              <Plus size={20} />
              Add Product
            </button>

            {/* Products List */}
            {warrantyProducts.length > 0 && (
              <div style={{ marginTop: "24px" }}>
                <h4
                  style={{
                    color: "#374151",
                    fontWeight: 600,
                    marginBottom: "16px",
                  }}
                >
                  Added Products ({warrantyProducts.length})
                </h4>
                <div
                  style={{
                    display: "flex",
                    flexDirection: "column",
                    gap: "8px",
                  }}
                >
                  {warrantyProducts.map((product) => (
                    <div
                      key={product.id}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "space-between",
                        padding: "12px",
                        backgroundColor: "#f8fafc",
                        border: "1px solid #e2e8f0",
                        borderRadius: "8px",
                      }}
                    >
                      <div style={{ flex: 1 }}>
                        <div
                          style={{
                            color: "#374151",
                            fontWeight: 500,
                            marginBottom: "4px",
                          }}
                        >
                          {product.itemName}
                        </div>
                        <div
                          style={{
                            color: "#64748b",
                            fontSize: "0.875rem",
                          }}
                        >
                          ₹{product.invoiceValue.toLocaleString()} • Qty:{" "}
                          {product.quantity} • {product.warrantyYears} years •{" "}
                          {product.source}
                        </div>
                      </div>
                      <button
                        onClick={() => removeProduct(product.id)}
                        style={{
                          padding: "8px",
                          backgroundColor: "transparent",
                          border: "none",
                          cursor: "pointer",
                          color: "#ef4444",
                          borderRadius: "4px",
                          transition: "color 0.2s ease",
                        }}
                        onMouseEnter={(e) => {
                          e.target.style.color = "#dc2626";
                        }}
                        onMouseLeave={(e) => {
                          e.target.style.color = "#ef4444";
                        }}
                      >
                        <Trash2 size={16} />
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {activeTab === "schedule" && (
          <div>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: "24px",
                flexWrap: "wrap",
                gap: "16px",
              }}
            >
              <h3
                style={{
                  fontSize: "1.25rem",
                  fontWeight: 700,
                  color: "#374151",
                }}
              >
                📊 Warranty Payment Totals
              </h3>

              <div
                style={{ display: "flex", gap: "12px", alignItems: "center" }}
              >
                {warrantyProducts.length > 0 && (
                  <button
                    onClick={calculateQuarterlySchedule}
                    disabled={isCalculating}
                    style={{ ...styles.button, ...styles.primaryButton }}
                  >
                    {isCalculating ? (
                      <>
                        <div
                          style={{
                            width: "20px",
                            height: "20px",
                            border: "2px solid transparent",
                            borderTop: "2px solid white",
                            borderRadius: "50%",
                            animation: "spin 1s linear infinite",
                          }}
                        ></div>
                        Calculating...
                      </>
                    ) : (
                      <>
                        <Calculator size={20} />
                        Calculate Schedule
                      </>
                    )}
                  </button>
                )}
                <WarrantyExportControls />
              </div>
            </div>

            {/* View Mode Toggle */}
            {calculatedSchedule.length > 0 && (
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: "12px",
                  marginBottom: "24px",
                  padding: "4px",
                  backgroundColor: "#f1f5f9",
                  borderRadius: "8px",
                  width: "fit-content",
                }}
              >
                <button
                  onClick={() => setViewMode("cards")}
                  style={{
                    ...styles.tab(viewMode === "cards"),
                    display: "flex",
                    alignItems: "center",
                    gap: "6px",
                  }}
                >
                  📊 Cards View
                </button>
                <button
                  onClick={() => setViewMode("table")}
                  style={{
                    ...styles.tab(viewMode === "table"),
                    display: "flex",
                    alignItems: "center",
                    gap: "6px",
                  }}
                >
                  📋 Table View
                </button>
              </div>
            )}

            {warrantyProducts.length === 0 ? (
              <div
                style={{
                  textAlign: "center",
                  padding: "48px 24px",
                }}
              >
                <Package
                  size={64}
                  style={{ color: "#9ca3af", margin: "0 auto 16px" }}
                />
                <h4
                  style={{
                    fontSize: "1.25rem",
                    fontWeight: 600,
                    color: "#374151",
                    marginBottom: "8px",
                  }}
                >
                  No Products Added
                </h4>
                <p
                  style={{
                    color: "#64748b",
                    marginBottom: "24px",
                  }}
                >
                  Add products using Excel upload or manual entry to calculate
                  warranty schedules.
                </p>
              </div>
            ) : calculatedSchedule.length === 0 ? (
              <div
                style={{
                  textAlign: "center",
                  padding: "48px 24px",
                }}
              >
              </div>
            ) : viewMode === "table" ? (
              <VirtualDataTable
                data={calculatedSchedule}
                columns={warrantyTableColumns}
                height={600}
                title={`Warranty Schedule - ${calculatedSchedule.length} items`}
                onExport={() => exportToExcel()}
                searchable={true}
                filterable={true}
                sortable={true}
                summary={warrantySummary}
                formatters={warrantyFormatters}
              />
            ) : (
              <div>
                {/* Quarterly Total Cards */}
                <div style={{ marginBottom: "32px" }}>
                  <h4 style={{
                    fontSize: "1.5rem",
                    fontWeight: 700,
                    color: "#374151",
                    marginBottom: "16px",
                    display: "flex",
                    alignItems: "center",
                    gap: "8px"
                  }}>
                    Quarterly Payment Totals
                  </h4>
                  <div style={{
                    display: "grid",
                    gridTemplateColumns: "repeat(auto-fit, minmax(250px, 1fr))",
                    gap: "16px"
                  }}>
                    {(() => {
      const quarterTotals = new Map();
      const quarterOrder = ["JFM", "AMJ", "JAS", "OND"];
      
      // Calculate quarterly totals
      calculatedSchedule.forEach((row) => {
        if (row.itemName === "Grand Total") return;
        
        Object.keys(row).forEach((key) => {
          if (key.match(/^(JFM|AMJ|JAS|OND) \d{4}$/)) {
            const amount = row[key] || 0;
            if (quarterTotals.has(key)) {
              quarterTotals.set(key, quarterTotals.get(key) + amount);
            } else {
              quarterTotals.set(key, amount);
            }
          }
        });
      });
      
      // Sort quarters chronologically
      const sortedQuarters = Array.from(quarterTotals.entries()).sort((a, b) => {
        const [qA, yearA] = a[0].split(" ");
        const [qB, yearB] = b[0].split(" ");
        
        const yearNumA = parseInt(yearA);
        const yearNumB = parseInt(yearB);
        
        if (yearNumA !== yearNumB) {
          return yearNumA - yearNumB;
        }
        
        return quarterOrder.indexOf(qA) - quarterOrder.indexOf(qB);
      });
      
      return sortedQuarters.map(([quarter, total]) => {
        const isPaid = paidQuarters[quarter]?.paid || false;
        const paidDate = paidQuarters[quarter]?.date || "";
        
        return (
          <div
            key={quarter}
            style={{
              background: isPaid 
                ? "linear-gradient(135deg, #059669 0%, #0d9488 100%)" 
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
                        updatePaidQuarters(quarter, true, payDate);
                      }
                    } else {
                      updatePaidQuarters(quarter, false);
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
              {quarter}
            </div>
            <div style={{
              fontSize: "1.5rem",
              fontWeight: 800,
              marginBottom: "4px"
            }}>
              ₹{total.toLocaleString()}
            </div>
            <div style={{
              fontSize: "0.75rem",
              opacity: 0.8,
              marginBottom: "8px"
            }}>
              {total >= 10000000 
                ? `₹${(total / 10000000).toFixed(2)} Cr`
                : total >= 100000
                ? `₹${(total / 100000).toFixed(2)} Lakh`
                : "Amount"
              }
            </div>
            
            {/* Payment Date */}
            {isPaid && paidDate && (
              <div style={{
                fontSize: "0.75rem",
                opacity: 0.9,
                borderTop: "1px solid rgba(255, 255, 255, 0.2)",
                paddingTop: "8px",
                marginTop: "8px"
              }}>
                💳 Paid on: {new Date(paidDate).toLocaleDateString('en-GB')}
              </div>
            )}
          </div>
        );
      });
    })()}
  </div>
</div>
                      {/* Yearly Total Cards */}
                      <div>
                        <h4 style={{
                          fontSize: "1.5rem",
                          fontWeight: 700,
                          color: "#374151",
                          marginBottom: "16px",
                          display: "flex",
                          alignItems: "center",
                          gap: "10px"
                        }}>
                          Yearly Totals
                        </h4>
                        <div style={{
                          display: "grid",
                          gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))",
                          gap: "16px"
                          }}>
                            {(() => {
                              const yearlyTotals = new Map();
                              // Calculate yearly totals from quarterly data
                              calculatedSchedule.forEach((row) => {
                                if (row.itemName === "Grand Total") return;
                                Object.keys(row).forEach((key) => {
                                  if (key.match(/^(JFM|AMJ|JAS|OND) \d{4}$/)) {
                                    const [quarter, year] = key.split(" ");
                                    const amount = row[key] || 0;
                                    if (yearlyTotals.has(year)) {
                                      yearlyTotals.set(year, yearlyTotals.get(year) + amount);
                                    } else {
                                      yearlyTotals.set(year, amount);
                                    }}});
                                  });    
                                  // Sort year
                                  const sortedYears = Array.from(yearlyTotals.entries()).sort((a, b) => parseInt(a[0]) - parseInt(b[0]));
                                  return sortedYears.map(([year, total]) => (
                                    <div
                                    key={year}
                                    style={{
                                      background: "linear-gradient(135deg, #059669 0%, #0d9488 100%)",
                                      borderRadius: "16px",
                                      padding: "24px",
                                      color: "white",
                                      boxShadow: "0 10px 30px -10px rgba(5, 150, 105, 0.3)",
                                    }}>
                                      <div style={{
                                        fontSize: "0.875rem",
                                        fontWeight: 600,
                                        marginBottom: "8px",
                                        opacity: 0.9
                                        }}>
                                          {year}
                                      </div>
                                      <div style={{
                                        fontSize: "2rem",
                                        fontWeight: 800,
                                        marginBottom: "8px",
                                      }}>₹ {total.toLocaleString()}
                                    </div>
                                    <div style={{
                                      fontSize: "0.75rem",
                                      opacity: 0.8,
                                      borderTop: "1px solid rgba(255, 255, 255, 0.2)",
                                      paddingTop: "8px",
                                      marginTop: "8px"
                                    }}>
                                      {(() => {
                                        // Count quarters for this year
                                        const quarters = calculatedSchedule
                                        .filter(row => row.itemName !== "Grand Total")
                                        .reduce((acc, row) => {
                                          Object.keys(row).forEach(key => {
                                            if (key.includes(year) && key.match(/^(JFM|AMJ|JAS|OND)/)) {
                                              const quarter = key.split(" ")[0];
                                              if (row[key] > 0 && !acc.includes(quarter)) {
                                                acc.push(quarter);
                                              }}});
                                              return acc;
                                            }, []);
                                            return `${quarters.length} active quarters`;
                                          })()}
                                        </div>
                                      </div>
                                      ));
                                    })()}
                                  </div>
                                </div>
                              </div>
                            )}
          </div>
        )}
        </div>

        {/* Payment Summary*/}
        <div style={{ marginBottom: "32px" }}>
          <div style={{
            background: "rgba(255, 255, 255, 0.95)",
            borderRadius: "16px",
            padding: "24px",
            border: "1px solid #ffffffff",
            boxShadow: "0 10px 30px -10px rgba(0, 0, 0, 0.1)",
            }}>
              <h4 style={{
                fontSize: "1.25rem",
                fontWeight: 700,
                color: "#374151",
                marginBottom: "20px",
                display: "flex",
                alignItems: "center",
                gap: "8px"
                }}>
                  Final Payment Summary
                </h4>
                {(() => {
                  const quarterTotals = new Map();
                  // Calculate quarterly totals
                  calculatedSchedule.forEach((row) => {
                    if (row.itemName === "Grand Total") return;
                    Object.keys(row).forEach((key) => {
                      if (key.match(/^(JFM|AMJ|JAS|OND) \d{4}$/)) {
                        const amount = row[key] || 0;
                        if (quarterTotals.has(key)) {
                          quarterTotals.set(key, quarterTotals.get(key) + amount);
                        } else {
                          quarterTotals.set(key, amount);
                        }}});
                      });
                      let totalPaid = 0;
                      let totalBalance = 0;
                      let paidCount = 0;
                      let totalCount = quarterTotals.size;
                      quarterTotals.forEach((amount, quarter) => {
                        if (paidQuarters[quarter]?.paid) {
                          totalPaid += amount;
                          paidCount++;
                        } else {
                          totalBalance += amount;
                        }
                      });
                      return (
                        <div style={{
                          display: "grid",
                          gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))",
                          gap: "16px"
                        }}>

                        <div style={{
                          textAlign: "center",
                          padding: "16px",
                          backgroundColor: "#ffffffff",
                          borderRadius: "12px",
                          border: "1px solid #e2e8f0"
                          }}>
                            <div style={{
                              fontSize: "0.875rem",
                              fontWeight: 600,
                              color: "#64748b",
                              marginBottom: "4px",
                            }}> Total Amount
                          </div>
                          <div style={{
                            fontSize: "1.25rem",
                            fontWeight: 800,
                            color: "#374151"
                          }}>
                            ₹ {(totalPaid + totalBalance).toLocaleString()}
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
                              }}> Paid Amount
                            </div>
                            <div style={{
                              fontSize: "1.25rem",
                              fontWeight: 800,
                              color: "#059669"
                            }}> ₹{totalPaid.toLocaleString()}
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
</div>
  );
};

export default WarrantyPaymentTracker;
