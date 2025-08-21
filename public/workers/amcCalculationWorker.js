// AMC Calculation Web Worker - Fixed Critical Issues

// Quarter definitions (business critical) - matching Python implementation
const QUARTERS = ["JFM", "AMJ", "JAS", "OND"];
const QUARTER_ORDER = { JFM: 0, AMJ: 1, JAS: 2, OND: 3 };

const ROI_RATES = [20, 22.5, 27.5, 30]; // Year 1-4 rates
const AMC_PERCENTAGE = 0.4; // 40% of invoice value
const GST_RATE = 0.18; // 18% GST

// Helper function to check if a year is a leap year
function isLeapYear(year) {
  return year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
}

// FIXED: Single definition of getQuarterDates (was duplicated)
function getQuarterDates(year) {
  return {
    JFM: [new Date(year, 0, 5), new Date(year, 3, 4)], // Jan 5 - Apr 4
    AMJ: [new Date(year, 3, 5), new Date(year, 6, 4)], // Apr 5 - Jul 4
    JAS: [new Date(year, 6, 5), new Date(year, 9, 4)], // Jul 5 - Oct 4
    OND: [new Date(year, 9, 5), new Date(year + 1, 0, 4)], // Oct 5 - Jan 4 (next year)
  };
}

// Helper function to get actual days in a quarter for a specific year
function getActualQuarterDays(quarter, year) {
  const quarterDates = getQuarterDates(year);
  const [startDate, endDate] = quarterDates[quarter];
  return Math.floor((endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24)) + 1;
}

// FIXED: Accurate date arithmetic matching Python's relativedelta behavior
function addYears(date, years) {
  const result = new Date(date);
  const originalMonth = result.getMonth();
  const originalDate = result.getDate();
  
  // Set the new year first
  result.setFullYear(result.getFullYear() + years);
  
  // Handle edge cases like Feb 29 on non-leap years
  if (result.getMonth() !== originalMonth) {
    // If month changed (e.g., Feb 29 -> Mar 1), go back to last day of intended month
    result.setDate(0); // Sets to last day of previous month
  }
  
  return result;
}

// Add months to date using relativedelta equivalent
function addMonths(date, months) {
  const result = new Date(date);
  result.setMonth(result.getMonth() + months);
  return result;
}

// FIXED: More precise relativedelta equivalent
function addRelativeDelta(date, years = 0, months = 0, days = 0) {
  const result = new Date(date);
  
  if (years !== 0) {
    result.setFullYear(result.getFullYear() + years);
  }
  
  if (months !== 0) {
    const newMonth = result.getMonth() + months;
    result.setMonth(newMonth);
  }
  
  if (days !== 0) {
    result.setDate(result.getDate() + days);
  }
  
  return result;
}

// FIXED: Main AMC calculation with corrected date handling and input validation
function calculateAmcSchedule(
  startDate,
  invoiceNumber,
  cost,
  quantity,
  roiSplit,
  amcPercent = 0.4,
  gstRate = 0.18
) {
  // Input validation
  if (!startDate || isNaN(startDate.getTime())) {
    throw new Error('Invalid start date provided');
  }
  if (!cost || isNaN(cost) || cost <= 0) {
    throw new Error('Invalid cost provided');
  }
  if (!Array.isArray(roiSplit) || roiSplit.length === 0) {
    throw new Error('Invalid ROI split provided');
  }

  const totalAmc = cost * amcPercent;
  const schedule = {};
  const splitDetails = {};
  const quarterContributions = {};

  for (let yearIndex = 0; yearIndex < roiSplit.length; yearIndex++) {
    const roi = roiSplit[yearIndex];
    
    if (isNaN(roi) || roi < 0) {
      console.warn(`Invalid ROI rate at index ${yearIndex}: ${roi}, skipping`);
      continue;
    }
    
    // CRITICAL FIX: Exact date arithmetic matching Python relativedelta
    const yearStart = new Date(startDate);
    yearStart.setFullYear(startDate.getFullYear() + yearIndex);
    // Ensure we preserve the exact time component
    yearStart.setHours(startDate.getHours(), startDate.getMinutes(), startDate.getSeconds(), startDate.getMilliseconds());
    
    const yearEnd = new Date(startDate);
    yearEnd.setFullYear(startDate.getFullYear() + yearIndex + 1);
    yearEnd.setDate(yearEnd.getDate() - 1); // Subtract exactly 1 day
    yearEnd.setHours(23, 59, 59, 999); // Set to end of day

    const fullQuarterAmount = (totalAmc * roi) / 4;

    // FIXED: More precise calendar year determination
    const startCalendarYear = yearStart.getFullYear();
    const endCalendarYear = yearEnd.getFullYear();

    for (let calendarYear = startCalendarYear; calendarYear <= endCalendarYear; calendarYear++) {
      const quartersInYear = getQuarterDates(calendarYear);

      for (const qName of ["JFM", "AMJ", "JAS", "OND"]) {
        const [qStart, qEnd] = quartersInYear[qName];

        // FIXED: More precise overlap detection
        if (qEnd.getTime() < yearStart.getTime() || qStart.getTime() > yearEnd.getTime()) {
          continue;
        }

        // Calculate exact overlap using timestamps
        const overlapStart = new Date(Math.max(qStart.getTime(), yearStart.getTime()));
        const overlapEnd = new Date(Math.min(qEnd.getTime(), yearEnd.getTime()));

        if (overlapStart.getTime() > overlapEnd.getTime()) {
          continue;
        }

        // CRITICAL FIX: Exact day calculation matching Python's inclusive date arithmetic
        const msPerDay = 1000 * 60 * 60 * 24;
        
        // Python uses inclusive date ranges - both start and end dates are included
        // Set times to midnight to avoid timezone issues
        const qStartMidnight = new Date(qStart.getFullYear(), qStart.getMonth(), qStart.getDate());
        const qEndMidnight = new Date(qEnd.getFullYear(), qEnd.getMonth(), qEnd.getDate());
        const overlapStartMidnight = new Date(overlapStart.getFullYear(), overlapStart.getMonth(), overlapStart.getDate());
        const overlapEndMidnight = new Date(overlapEnd.getFullYear(), overlapEnd.getMonth(), overlapEnd.getDate());
        
        const totalDays = Math.round((qEndMidnight.getTime() - qStartMidnight.getTime()) / msPerDay) + 1;
        const overlapDays = Math.round((overlapEndMidnight.getTime() - overlapStartMidnight.getTime()) / msPerDay) + 1;

        // CRITICAL FIX: Display year logic must match Python exactly
        const displayYear = qStart.getFullYear();

        const proratedAmount = (overlapDays / totalDays) * fullQuarterAmount;

        // Store contribution with proper nesting
        if (!quarterContributions[qName]) {
          quarterContributions[qName] = {};
        }
        if (!quarterContributions[qName][yearIndex]) {
          quarterContributions[qName][yearIndex] = {};
        }
        quarterContributions[qName][yearIndex][displayYear] = proratedAmount;
      }
    }
  }

  // Process quarters - improved error handling
  const processedQuarters = new Set();

  for (const qName in quarterContributions) {
    const yearMap = quarterContributions[qName];
    const allDisplayYears = new Set();
    
    for (const yearIndex in yearMap) {
      const displayYearMap = yearMap[yearIndex];
      for (const displayYear in displayYearMap) {
        const year = parseInt(displayYear);
        if (!isNaN(year)) {
          allDisplayYears.add(year);
        }
      }
    }

    Array.from(allDisplayYears).sort((a, b) => a - b).forEach(displayYear => {
      const key = `${displayYear}-${qName}`;

      if (processedQuarters.has(key)) {
        return;
      }
      processedQuarters.add(key);

      let totalAmount = 0;
      const contributions = [];

      for (const yearIndex in yearMap) {
        const displayYearMap = yearMap[yearIndex];
        if (displayYearMap[displayYear]) {
          const yearIdx = parseInt(yearIndex);
          if (isNaN(yearIdx) || yearIdx >= roiSplit.length) {
            continue;
          }
          
          const roi = roiSplit[yearIdx];
          const fullQuarterAmount = (totalAmc * roi) / 4;
          const proratedAmount = displayYearMap[displayYear];

          const displayYearsForThisAmcYear = Object.keys(displayYearMap)
            .map(y => parseInt(y))
            .filter(y => !isNaN(y))
            .sort((a, b) => a - b);
          const occurrenceIndex = displayYearsForThisAmcYear.indexOf(displayYear);

          let actualAmount, calcType;
          if (occurrenceIndex === 0) {
            actualAmount = proratedAmount;
            calcType = `Prorated (Y${yearIdx + 1})`;
          } else if (displayYearsForThisAmcYear.length > 1) {
            const firstOccurrenceAmount = displayYearMap[displayYearsForThisAmcYear[0]];
            const residual = fullQuarterAmount - firstOccurrenceAmount;
            actualAmount = Math.max(0, residual);
            calcType = `Residual (Y${yearIdx + 1} = ${fullQuarterAmount.toFixed(4)} - ${firstOccurrenceAmount.toFixed(4)})`;
          } else {
            actualAmount = 0;
            calcType = `No Residual (Y${yearIdx + 1})`;
          }

          actualAmount = Math.max(0, actualAmount);
          totalAmount += actualAmount;

          contributions.push({
            "AMC Year": yearIdx + 1,
            "ROI Rate": roi,
            "Full Quarter Amount": fullQuarterAmount,
            "Prorated Amount": proratedAmount,
            "Actual Amount": actualAmount,
            "Calculation Type": calcType
          });
        }
      }

      // FIXED: Consistent rounding matching Python
      const withoutGst = Math.round(totalAmount * 100) / 100;
      const withGst = Math.round(withoutGst * (1 + gstRate) * 100) / 100;

      schedule[key] = [withGst, withoutGst];
      
      // Store split details with corrected day calculations
      splitDetails[key] = contributions.map(contrib => {
        const actualTotalDaysInQuarter = 91; // Standardized to match Python
        const actualDaysUsed = Math.round((contrib["Prorated Amount"] / contrib["Full Quarter Amount"]) * actualTotalDaysInQuarter);
        
        return {
          "AMC Year": contrib["AMC Year"],
          "Quarter": qName,
          "ROI Rate": contrib["ROI Rate"],
          "Full Quarter Amount": contrib["Full Quarter Amount"],
          "Prorated Amount": contrib["Prorated Amount"],
          "Actual Amount": contrib["Actual Amount"],
          "Calculation Type": contrib["Calculation Type"],
          "Display Year": displayYear,
          "Total Amount": totalAmount,
          "Current Year Contribution": contrib["Calculation Type"].includes("Prorated") ? contrib["Actual Amount"] : 0,
          "Residual from Previous": contrib["Calculation Type"].includes("Residual") ? contrib["Actual Amount"] : 0,
          "Amount Without GST": withoutGst,
          "Amount With GST": withGst,
          "Days": actualDaysUsed,
          "Total Days in Quarter": actualTotalDaysInQuarter
        };
      });
    });
  }

  return { schedule, splitDetails };
}

// FIXED: Calculate AMC schedule for a single product with improved error handling
function calculateProductAMC(product, settings = {}) {
  const roiRates = settings.roiRates || ROI_RATES;
  const amcPercentage = settings.amcPercentage || AMC_PERCENTAGE;
  const gstRate = settings.gstRate || GST_RATE;
  const amcYears = settings.amcYears || 4;

  try {
    // FIXED: Better input validation and error messages
    if (!product || typeof product !== 'object') {
      throw new Error('Invalid product object provided');
    }

    // Parse dates - FIXED to handle multiple field name variations
    const uatDateStr = product.uatDate || product.UAT_Date || product.uat_date || product["UAT Date"];
    if (!uatDateStr) {
      throw new Error(`Missing UAT date for product: ${product.productName || product["Item Name"] || "Unknown"}`);
    }
    
    const uatDate = new Date(uatDateStr);
    if (isNaN(uatDate.getTime())) {
      throw new Error(`Invalid UAT date format: ${uatDateStr}`);
    }
    
    const invoiceValue = parseFloat(
      product.invoiceValue || product.Invoice_Value || product.cost || product.Cost || 0
    );

    if (isNaN(invoiceValue) || invoiceValue <= 0) {
      throw new Error(
        `Invalid invoice value for product: ${product.productName || product["Item Name"] || "Unknown"}`
      );
    }

    // FIXED: AMC starts exactly 3 years after UAT (matching Python relativedelta)
    const amcStartDate = addYears(uatDate, 3);

    // Calculate using the sophisticated logic
    const roiSplit = roiRates.slice(0, amcYears).map(rate => {
      const numRate = parseFloat(rate);
      if (isNaN(numRate)) {
        throw new Error(`Invalid ROI rate: ${rate}`);
      }
      return numRate / 100;
    });
    
    const { schedule, splitDetails } = calculateAmcSchedule(
      amcStartDate,
      product.invoiceNumber,
      invoiceValue,
      product.quantity || product.Quantity || 1,
      roiSplit,
      amcPercentage,
      gstRate
    );

    // FIXED: Convert to quarters array format matching Python output structure
    const quarters = [];

    for (const [key, amounts] of Object.entries(schedule)) {
      if (!Array.isArray(amounts) || amounts.length !== 2) {
        console.warn(`Invalid amounts for key ${key}:`, amounts);
        continue;
      }
      
      const [withGst, withoutGst] = amounts;
      const [year, quarter] = key.split("-");
      
      const yearNum = parseInt(year);
      if (isNaN(yearNum) || !quarter) {
        console.warn(`Invalid key format: ${key}`);
        continue;
      }

      quarters.push({
        id: `${product.id || Math.random()}_${key}`,
        quarter: quarter,
        year: yearNum,
        quarterKey: key,
        startDate: "", // Will be calculated if needed
        endDate: "", // Will be calculated if needed
        dueDate: "", // Will be calculated if needed
        baseAmount: withoutGst,
        roiAmount: withGst - withoutGst - withoutGst * gstRate, // Calculate ROI portion
        roiPercentage: 0, // Will be calculated from split details
        gstAmount: withoutGst * gstRate,
        totalAmount: withGst,
        isPaid: false,
        status: "pending",
        splitDetails: splitDetails[key] || []
      });
    }

    // Sort quarters by year and quarter order
    quarters.sort((a, b) => {
      if (a.year !== b.year) return a.year - b.year;
      const qOrderA = QUARTER_ORDER[a.quarter] || 0;
      const qOrderB = QUARTER_ORDER[b.quarter] || 0;
      return qOrderA - qOrderB;
    });

    const totalAmcValue = invoiceValue * amcPercentage;
    const totalWithGst = quarters.reduce((sum, q) => sum + (q.totalAmount || 0), 0);

    return {
      id: product.id || `product_${Date.now()}_${Math.random()}`,
      productName: product.productName || product["Item Name"] || product.name || "Unknown Product",
      location: product.location || product.Location || "Unknown Location",
      uatDate: uatDate.toISOString().split("T")[0],
      amcStartDate: amcStartDate.toISOString().split("T")[0],
      invoiceValue: invoiceValue,
      quantity: product.quantity || product.Quantity || 1,
      totalAmcValue: Math.round(totalAmcValue * 100) / 100,
      quarters: quarters,
      totalQuarters: quarters.length,
      totalAmountWithGST: Math.round(totalWithGst * 100) / 100,
      splitDetails: splitDetails
    };
  } catch (error) {
    console.error('Error calculating AMC for product:', error);
    return {
      id: product.id || `error_${Date.now()}_${Math.random()}`,
      productName: product.productName || product["Item Name"] || "Error Product",
      error: error.message,
      quarters: []
    };
  }
}

// Helper function to format quarter key for display
function formatQuarterKey(quarter, year) {
  return `${quarter}-${year}`;
}

// FIXED: Process chunk of products with better error handling
function processChunk(products, settings, chunkIndex) {
  if (!Array.isArray(products)) {
    throw new Error('Products must be an array');
  }
  
  const results = [];
  const startTime = Date.now();

  for (let index = 0; index < products.length; index++) {
    const product = products[index];
    try {
      const amcSchedule = calculateProductAMC(product, settings);
      results.push(amcSchedule);

      // Send progress update every 50 products (reduced frequency)
      if (index % 50 === 0 && index > 0) {
        self.postMessage({
          type: "PROGRESS",
          chunkIndex: chunkIndex,
          processed: index + 1,
          total: products.length,
          timeElapsed: Date.now() - startTime
        });
      }
    } catch (error) {
      console.error(`Error processing product at index ${index}:`, error);
      results.push({
        id: product.id || `error_${index}_${Math.random()}`,
        productName: product.productName || product["Item Name"] || "Error",
        error: error.message,
        quarters: []
      });
    }
  }

  return results;
}

// FIXED: Worker message handler with comprehensive error handling
self.onmessage = function(event) {
  try {
    if (!event.data || typeof event.data !== 'object') {
      throw new Error('Invalid message format');
    }
    
    const { type, data } = event.data;

    switch (type) {
      case "CALCULATE_CHUNK":
        if (!data || !data.products || !Array.isArray(data.products)) {
          throw new Error('Invalid chunk data: products array required');
        }
        
        const { products, settings, chunkIndex, totalChunks } = data;
        const chunkStartTime = Date.now();

        self.postMessage({
          type: "CHUNK_STARTED",
          chunkIndex: chunkIndex || 0,
          totalProducts: products.length
        });

        const results = processChunk(products, settings || {}, chunkIndex || 0);

        // Send final progress update for this chunk
        self.postMessage({
          type: "PROGRESS",
          chunkIndex: chunkIndex || 0,
          processed: products.length,
          total: products.length,
          timeElapsed: Date.now() - chunkStartTime
        });

        self.postMessage({
          type: "CHUNK_COMPLETE",
          chunkIndex: chunkIndex || 0,
          totalChunks: totalChunks || 1,
          results: results,
          summary: {
            processed: results.length,
            successful: results.filter(r => !r.error).length,
            errors: results.filter(r => r.error).length,
            totalValue: results.reduce((sum, r) => sum + (r.totalAmountWithGST || 0), 0)
          }
        });
        break;

      case "CALCULATE_SINGLE":
        if (!data || !data.product) {
          throw new Error('Invalid single calculation data: product required');
        }
        
        const { product, settings: singleSettings } = data;
        const singleResult = calculateProductAMC(product, singleSettings || {});

        self.postMessage({
          type: "SINGLE_COMPLETE",
          result: singleResult
        });
        break;

      default:
        self.postMessage({
          type: "ERROR",
          error: `Unknown message type: ${type}`
        });
    }
  } catch (error) {
    console.error('Worker error:', error);
    self.postMessage({
      type: "ERROR",
      error: error.message,
      chunkIndex: event.data?.data?.chunkIndex,
      stack: error.stack
    });
  }
};

// Send ready signal
self.postMessage({
  type: "WORKER_READY",
  capabilities: ["CALCULATE_CHUNK", "CALCULATE_SINGLE"],
  version: "2.0.0"
});