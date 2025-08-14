
// Quarter definitions (business critical) - matching Python implementation
const QUARTERS = ["JFM", "AMJ", "JAS", "OND"];
const QUARTER_ORDER = { JFM: 0, AMJ: 1, JAS: 2, OND: 3 };

const ROI_RATES = [20, 22.5, 27.5, 30]; // Year 1-4 rates
const AMC_PERCENTAGE = 0.4; // 40% of invoice value
const GST_RATE = 0.18; // 18% GST

// Get quarter date ranges for a given year - matching Python logic
function getQuarterDates(year) {
  return {
    JFM: [new Date(year, 0, 5), new Date(year, 3, 4)], // Jan 5 - Apr 4
    AMJ: [new Date(year, 3, 5), new Date(year, 6, 4)], // Apr 5 - Jul 4
    JAS: [new Date(year, 6, 5), new Date(year, 9, 4)], // Jul 5 - Oct 4
    OND: [new Date(year, 9, 5), new Date(year + 1, 0, 4)], // Oct 5 - Jan 4 (next year)
  };
}

// Add months to date (helper function)
function addMonths(date, months) {
  const result = new Date(date);
  result.setMonth(result.getMonth() + months);
  return result;
}

// Add years to date (helper function)
function addYears(date, years) {
  const result = new Date(date);
  result.setFullYear(result.getFullYear() + years);
  return result;
}

// Calculate AMC schedule with complex quarter overlap logic - JavaScript port of Python function
function calculateAmcSchedule(
  startDate,
  cost,
  quantity,
  roiSplit,
  amcPercent,
  gstRate
) {
  // Set default values for compatibility
  if (typeof amcPercent === "undefined") amcPercent = 0.4;
  if (typeof gstRate === "undefined") gstRate = 0.18;

  const totalAmc = cost * amcPercent;
  const schedule = {};
  const splitDetails = {};

  // Track quarter contributions: quarter_name -> {year_index: {display_year: prorated_amount}}
  const quarterContributions = {};

  for (let yearIndex = 0; yearIndex < roiSplit.length; yearIndex++) {
    const roi = roiSplit[yearIndex];
    // Calculate AMC year boundaries
    const yearStart = addYears(startDate, yearIndex);
    const yearEnd = new Date(addYears(startDate, yearIndex + 1));
    yearEnd.setDate(yearEnd.getDate() - 1);

    // Calculate full quarter amount for this AMC year
    const fullQuarterAmount = (totalAmc * roi) / 4;

    // Determine which calendar years to check for quarters
    const startCalendarYear = yearStart.getFullYear();
    const endCalendarYear = yearEnd.getFullYear();

    // Check all quarters that might overlap with this AMC year
    for (
      let calendarYear = startCalendarYear;
      calendarYear <= endCalendarYear;
      calendarYear++
    ) {
      const quartersInYear = getQuarterDates(calendarYear);

      for (let q = 0; q < QUARTERS.length; q++) {
        const qName = QUARTERS[q];
        const qStart = quartersInYear[qName][0];
        const qEnd = quartersInYear[qName][1];

        // Check if quarter overlaps with AMC year
        if (qEnd < yearStart || qStart > yearEnd) {
          continue;
        }

        // Calculate overlap
        const overlapStart = new Date(
          Math.max(qStart.getTime(), yearStart.getTime())
        );
        const overlapEnd = new Date(
          Math.min(qEnd.getTime(), yearEnd.getTime())
        );

        if (overlapStart > overlapEnd) {
          continue;
        }

        const totalDays =
          Math.floor(
            (qEnd.getTime() - qStart.getTime()) / (1000 * 60 * 60 * 24)
          ) + 1;
        const overlapDays =
          Math.floor(
            (overlapEnd.getTime() - overlapStart.getTime()) /
              (1000 * 60 * 60 * 24)
          ) + 1;

        // Determine the display year
        const displayYear =
          qName === "OND" ? qStart.getFullYear() : qStart.getFullYear();

        // Calculate prorated amount for this AMC year in this quarter occurrence
        const proratedAmount = (overlapDays / totalDays) * fullQuarterAmount;

        // Store the contribution using plain objects
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

  // Calculate the actual amounts to be allocated
  const processedQuarters = {};

  for (const qName in quarterContributions) {
    const yearMap = quarterContributions[qName];

    // Get all display years for this quarter across all AMC years
    const allDisplayYears = [];
    for (const yearIndex in yearMap) {
      const displayYearMap = yearMap[yearIndex];
      for (const displayYear in displayYearMap) {
        if (allDisplayYears.indexOf(parseInt(displayYear)) === -1) {
          allDisplayYears.push(parseInt(displayYear));
        }
      }
    }

    // Process each display year occurrence of this quarter
    allDisplayYears.sort().forEach(function (displayYear) {
      const key = displayYear + "-" + qName;

      if (processedQuarters[key]) {
        return;
      }
      processedQuarters[key] = true;

      let totalAmount = 0;
      const contributions = [];

      // Sum contributions from all AMC years that overlap with this quarter occurrence
      for (const yearIndex in yearMap) {
        const displayYearMap = yearMap[yearIndex];
        if (displayYearMap[displayYear]) {
          const roi = roiSplit[parseInt(yearIndex)];
          const fullQuarterAmount = (totalAmc * roi) / 4;
          const proratedAmount = displayYearMap[displayYear];

          // Check if this AMC year has other occurrences of this quarter
          const displayYearsForThisAmcYear = Object.keys(displayYearMap)
            .map(function (y) {
              return parseInt(y);
            })
            .sort();
          const occurrenceIndex =
            displayYearsForThisAmcYear.indexOf(displayYear);

          let actualAmount, calcType;
          if (occurrenceIndex === 0) {
            // Always include prorated amount for first occurrence
            actualAmount = proratedAmount;
            calcType = "Prorated (Y" + (parseInt(yearIndex) + 1) + ")";
          } else if (displayYearsForThisAmcYear.length > 1) {
            // Only apply residual if there are multiple contributions
            const firstOccurrenceAmount =
              displayYearMap[displayYearsForThisAmcYear[0]];
            const residual = fullQuarterAmount - firstOccurrenceAmount;
            actualAmount = Math.max(0, residual);
            calcType =
              "Residual (Y" +
              (parseInt(yearIndex) + 1) +
              " = " +
              fullQuarterAmount.toFixed(4) +
              " - " +
              firstOccurrenceAmount.toFixed(4) +
              ")";
          } else {
            // No residual needed
            actualAmount = 0;
            calcType = "No Residual (Y" + (parseInt(yearIndex) + 1) + ")";
          }

          // Ensure amount is not negative
          actualAmount = Math.max(0, actualAmount);
          totalAmount += actualAmount;

          contributions.push({
            amcYear: parseInt(yearIndex) + 1,
            roiRate: roi,
            fullQuarterAmount: fullQuarterAmount,
            proratedAmount: proratedAmount,
            actualAmount: actualAmount,
            calculationType: calcType,
          });
        }
      }

      // Calculate GST amounts
      const withoutGst = Math.round(totalAmount * 100) / 100;
      const withGst = Math.round(withoutGst * (1 + gstRate) * 100) / 100;

      // Add to schedule
      schedule[key] = [withGst, withoutGst];

      // Store detailed split information
      splitDetails[key] = contributions.map(function (contrib) {
        return {
          amcYear: contrib.amcYear,
          quarter: qName,
          roiRate: contrib.roiRate,
          fullQuarterAmount: contrib.fullQuarterAmount,
          proratedAmount: contrib.proratedAmount,
          actualAmount: contrib.actualAmount,
          calculationType: contrib.calculationType,
          displayYear: displayYear,
          totalAmount: totalAmount,
          currentYearContribution:
            contrib.calculationType.indexOf("Prorated") !== -1
              ? contrib.actualAmount
              : 0,
          residualFromPrevious:
            contrib.calculationType.indexOf("Residual") !== -1
              ? contrib.actualAmount
              : 0,
          amountWithoutGst: withoutGst,
          amountWithGst: withGst,
          days: Math.round(
            (contrib.proratedAmount / contrib.fullQuarterAmount) * 91
          ),
          totalDaysInQuarter: 91,
        };
      });
    });
  }

  return { schedule: schedule, splitDetails: splitDetails };
}

// Calculate AMC schedule for a single product using the new logic
function calculateProductAMC(product, settings) {
  // Handle default parameters for compatibility
  if (typeof settings === "undefined") settings = {};

  const roiRates = settings.roiRates || ROI_RATES;
  const amcPercentage = settings.amcPercentage || AMC_PERCENTAGE;
  const gstRate = settings.gstRate || GST_RATE;
  const amcYears = settings.amcYears || 4;

  try {
    // Parse dates
    const uatDate = new Date(
      product.uatDate || product.UAT_Date || product.uat_date
    );
    const invoiceValue = parseFloat(
      product.invoiceValue || product.Invoice_Value || product.cost || 0
    );

    if (isNaN(invoiceValue) || invoiceValue <= 0) {
      throw new Error(
        `Invalid invoice value for product: ${product.productName || "Unknown"}`
      );
    }

    // AMC starts exactly 3 years after UAT
    const amcStartDate = new Date(uatDate);
    amcStartDate.setFullYear(amcStartDate.getFullYear() + 3);

    // Calculate using the sophisticated logic
    const roiSplit = roiRates.slice(0, amcYears).map(function (rate) {
      return rate / 100;
    });
    const { schedule, splitDetails } = calculateAmcSchedule(
      amcStartDate,
      invoiceValue,
      product.quantity || 1,
      roiSplit,
      amcPercentage,
      gstRate
    );

    // Convert to quarters array format
    const quarters = [];

    for (const key in schedule) {
      const amounts = schedule[key];
      const withGst = amounts[0];
      const withoutGst = amounts[1];
      const keyParts = key.split("-");

      // Key format is "year-quarter" (e.g., "2024-JFM")
      const year = keyParts[0];
      const quarter = keyParts[1];

      quarters.push({
        id: (product.id || Math.random()) + "_" + key,
        quarter: quarter,
        year: parseInt(year),
        quarterKey: key,
        startDate: "", // Will be calculated if needed
        endDate: "", // Will be calculated if needed
        dueDate: "", // Will be calculated if needed
        baseAmount: withoutGst,
        roiAmount: withGst - withoutGst - withoutGst * gstRate,
        roiPercentage: 0, // Will be calculated from split details
        gstAmount: withoutGst * gstRate,
        totalAmount: withGst,
        isPaid: false,
        status: "pending",
        splitDetails: splitDetails[key] || [],
      });
    }

    // Sort quarters by year and quarter order
    quarters.sort(function (a, b) {
      if (a.year !== b.year) return a.year - b.year;
      const qOrderA = QUARTER_ORDER[a.quarter] || 0;
      const qOrderB = QUARTER_ORDER[b.quarter] || 0;
      return qOrderA - qOrderB;
    });

    const totalAmcValue = invoiceValue * amcPercentage;
    const totalWithGst = quarters.reduce(function (sum, q) {
      return sum + q.totalAmount;
    }, 0);

    return {
      id: product.id || Math.random(),
      productName:
        product.productName ||
        product.Product_Name ||
        product.name ||
        "Unknown Product",
      location: product.location || product.Location || "Unknown Location",
      uatDate: uatDate.toISOString().split("T")[0],
      amcStartDate: amcStartDate.toISOString().split("T")[0],
      invoiceValue: invoiceValue,
      quantity: product.quantity || 1,
      totalAmcValue: Math.round(totalAmcValue * 100) / 100,
      quarters: quarters,
      totalQuarters: quarters.length,
      totalAmountWithGST: Math.round(totalWithGst * 100) / 100,
      splitDetails: splitDetails,
    };
  } catch (error) {
    console.error("Error calculating AMC for product:", product, error);
    return {
      id: product.id || Math.random(),
      productName: product.productName || "Error Product",
      error: error.message,
      quarters: [],
    };
  }
}

// Helper function to format quarter key for display
function formatQuarterKey(quarter, year) {
  return `${quarter}-${year}`;
}

// Process chunk of products
function processChunk(products, settings, chunkIndex) {
  const results = [];
  const startTime = Date.now();

  for (let index = 0; index < products.length; index++) {
    const product = products[index];
    try {
      const amcSchedule = calculateProductAMC(product, settings);
      results.push(amcSchedule);

      // Send progress update every 100 products
      if (index % 100 === 0) {
        self.postMessage({
          type: "PROGRESS",
          chunkIndex: chunkIndex,
          processed: index + 1,
          total: products.length,
          timeElapsed: Date.now() - startTime,
        });
      }
    } catch (error) {
      console.error("Error processing product:", error);
      results.push({
        id: product.id || Math.random(),
        productName: product.productName || "Error",
        error: error.message,
        quarters: [],
      });
    }
  }
  return results;
}

// Worker message handler
self.onmessage = function (event) {
  const type = event.data.type;
  const data = event.data.data;

  try {
    switch (type) {
      case "CALCULATE_CHUNK":
        const products = data.products;
        const settings = data.settings;
        const chunkIndex = data.chunkIndex;
        const totalChunks = data.totalChunks;
        const chunkStartTime = Date.now();

        self.postMessage({
          type: "CHUNK_STARTED",
          chunkIndex: chunkIndex,
          totalProducts: products.length,
        });

        const results = processChunk(products, settings, chunkIndex);

        // Send final progress update for this chunk
        self.postMessage({
          type: "PROGRESS",
          chunkIndex: chunkIndex,
          processed: products.length, // Full chunk processed
          total: products.length,
          timeElapsed: Date.now() - chunkStartTime,
        });

        self.postMessage({
          type: "CHUNK_COMPLETE",
          chunkIndex: chunkIndex,
          totalChunks: totalChunks,
          results: results,
          summary: {
            processed: results.length,
            successful: results.filter(function (r) {
              return !r.error;
            }).length,
            errors: results.filter(function (r) {
              return r.error;
            }).length,
            totalValue: results.reduce(function (sum, r) {
              return sum + (r.totalAmountWithGST || 0);
            }, 0),
          },
        });
        break;

      case "CALCULATE_SINGLE":
        const product = data.product;
        const singleSettings = data.settings;
        const singleResult = calculateProductAMC(product, singleSettings);

        self.postMessage({
          type: "SINGLE_COMPLETE",
          result: singleResult,
        });
        break;

      default:
        self.postMessage({
          type: "ERROR",
          error: "Unknown message type: " + type,
        });
    }
  } catch (error) {
    self.postMessage({
      type: "ERROR",
      error: error.message,
      chunkIndex: data.chunkIndex,
    });
  }
};

// Send ready signal
self.postMessage({
  type: "WORKER_READY",
  capabilities: ["CALCULATE_CHUNK", "CALCULATE_SINGLE"],
});