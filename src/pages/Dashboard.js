//dashboard-redux.js
import React, { useEffect, useCallback, useState } from "react";
import { useNavigate } from "react-router-dom";
import { useSelector, useDispatch } from "react-redux";

// Redux Actions
import {
  processExcelFile,
  loadExcelFromMemory,
  clearData as clearExcelData,
} from "../store/slice/excelSlice";

import {
  loadAMCFromMemory,
  clearData,
} from "../store/slice/amcScheduleSlice";

import {
  loadWarrantyFromMemory,
  clearWarrantyData,
} from "../store/slice/warrantyDataSlice";

import {
  setHoveredCard,
  setIsMobile,
  setCurrentTime,
  setSearchTerm,
  setSelectedCategory,
  toggleNotifications,
  updateStats,
  nextTip,
  setShowSearch,
} from "../store/slice/dashboardSlice";

// Redux Selectors
import {
  selectFileName,
  selectSheets,
  selectIsLoading,
  selectError,
  selectHasData,
  selectTotalRows,
} from "../store/selectors/excelSelectors";

import {
  selectHasAMCData,
  selectTotalAMCAssets,
  selectTotalAMCValue,
} from "../store/selectors/amcDataSelectors";

import {
  selectHasWarrantyData,
  selectTotalWarrantyAssets,
  selectTotalWarrantyValue,
} from "../store/selectors/warrantyDataSelectors";

import {
  selectHoveredCard,
  selectIsMobile,
  selectCurrentTime,
  selectSearchTerm,
  selectSelectedCategory,
  selectStats,
  selectShowSearch,
  selectRecentActivity,
  selectFilteredTools,
} from "../store/selectors/dashboardSelectors";

const ExcelUploadSection = () => {
  const dispatch = useDispatch();

  const fileName = useSelector(selectFileName);
  const sheets = useSelector(selectSheets);
  const isLoading = useSelector(selectIsLoading);
  const error = useSelector(selectError);
  const hasData = useSelector(selectHasData);
  const totalRows = useSelector(selectTotalRows);

  // AMC selectors
  const hasAMCData = useSelector(selectHasAMCData);
  const totalAMCAssets = useSelector(selectTotalAMCAssets);
  const totalAMCValue = useSelector(selectTotalAMCValue);

  // Warranty selectors
  const hasWarrantyData = useSelector(selectHasWarrantyData);
  const totalWarrantyAssets = useSelector(selectTotalWarrantyAssets);
  const totalWarrantyValue = useSelector(selectTotalWarrantyValue);

  const [dragActive, setDragActive] = useState(false);

  const handleFile = useCallback(
    async (file) => {
      if (!file) {
        console.error("❌ No file provided to handleFile");
        return;
      }

      const validTypes = [
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.ms-excel",
        ".xlsx",
        ".xls",
      ];

      const isValidType = validTypes.some(
        (type) => file.type === type || file.name.toLowerCase().endsWith(type)
      );

      if (!isValidType) {
        alert("Please upload a valid Excel file (.xlsx or .xls)");
        return;
      }

      try {
        // Dispatch the action and wait for it to complete
        const result = await dispatch(processExcelFile(file));

        // Check if the action was successful
        if (result.type && result.type.endsWith("/fulfilled")) {

          // Show success notification
          alert("✅ File uploaded successfully!");
        } else if (result.type && result.type.endsWith("/rejected")) {
          console.error("❌ File upload failed:", result.error);
        }
      } catch (error) {
        console.error("❌ Error during file processing:", error);
      }
    },
    [dispatch]
  );

  // Add useEffect to monitor state changes
  useEffect(() => {}, [hasData, fileName, sheets, totalRows, isLoading, error]);

  const handleDrag = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();

    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  }, []);

  const handleDrop = useCallback(
    (e) => {
      e.preventDefault();
      e.stopPropagation();
      setDragActive(false);

      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        handleFile(e.dataTransfer.files[0]);
      } else {
        console.error("❌ No files found in drop event");
      }
    },
    [handleFile]
  );

  const handleFileInput = useCallback(
    (e) => {
      if (e.target.files && e.target.files[0]) {
        handleFile(e.target.files[0]);
      } else {
        console.error("❌ No files found in file input");
      }
    },
    [handleFile]
  );

  const handleClearAllData = useCallback(() => {
    if (
      window.confirm(
        "Are you sure you want to clear all data? This will remove Excel files, AMC calculations, and Warranty calculations."
      )
    ) {
      dispatch(clearExcelData());
      dispatch(clearData());
      dispatch(clearWarrantyData());
    }
  }, [dispatch]);

  return (
    <div
      style={{
        background: "rgba(255, 255, 255, 0.95)",
        backdropFilter: "blur(20px)",
        borderRadius: "20px",
        padding: "32px",
        marginBottom: "40px",
        border: "1px solid rgba(255, 255, 255, 0.2)",
        boxShadow: "0 20px 40px -12px rgba(0, 0, 0, 0.15)",
      }}
    >
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          marginBottom: "24px",
        }}
      >
        <div>
          <h3
            style={{
              fontSize: "1.5rem",
              fontWeight: 700,
              color: "#1e293b",
              marginBottom: "8px",
            }}
          >
            📊 Global Data Management Hub
          </h3>
          <p
            style={{
              color: "#64748b",
              fontSize: "0.95rem",
              margin: 0,
            }}
          >
            Centralized data storage for Excel files, AMC calculations, and
            Warranty calculations
          </p>
        </div>
        {(hasData || hasAMCData || hasWarrantyData) && (
          <button
            onClick={handleClearAllData}
            style={{
              padding: "8px 16px",
              backgroundColor: "#ef4444",
              color: "white",
              border: "none",
              borderRadius: "8px",
              fontSize: "0.875rem",
              cursor: "pointer",
              transition: "all 0.2s ease",
            }}
          >
            Clear All Data
          </button>
        )}
      </div>

      {/* Data Status Grid */}
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))",
          gap: "16px",
          marginBottom: "24px",
        }}
      >
        {/* Excel Data Status */}
        <div
          style={{
            padding: "16px",
            backgroundColor: hasData ? "#f0fdf4" : "#f8fafc",
            border: `1px solid ${hasData ? "#bbf7d0" : "#e2e8f0"}`,
            borderRadius: "12px",
          }}
        >
          <div style={{ fontSize: "1.5rem", marginBottom: "8px" }}>
            {hasData ? "📊" : "📄"}
          </div>
          <h4
            style={{
              fontSize: "0.9rem",
              fontWeight: 600,
              marginBottom: "4px",
              color: "#1e293b",
            }}
          >
            Excel Files
          </h4>
          <p
            style={{
              fontSize: "0.8rem",
              color: "#64748b",
              margin: 0,
            }}
          >
            {hasData
              ? `${sheets.length} sheets, ${totalRows.toLocaleString()} rows`
              : "No files uploaded"}
          </p>
        </div>

        {/* AMC Data Status */}
        <div
          style={{
            padding: "16px",
            backgroundColor: hasAMCData ? "#eff6ff" : "#f8fafc",
            border: `1px solid ${hasAMCData ? "#bfdbfe" : "#e2e8f0"}`,
            borderRadius: "12px",
          }}
        >
          <div style={{ fontSize: "1.5rem", marginBottom: "8px" }}>
            {hasAMCData ? "🔧" : "⚙️"}
          </div>
          <h4
            style={{
              fontSize: "0.9rem",
              fontWeight: 600,
              marginBottom: "4px",
              color: "#1e293b",
            }}
          >
            AMC Calculations
          </h4>
          <p style={{ fontSize: "0.8rem", color: "#64748b", margin: 0 }}>
            {hasAMCData
              ? `${totalAMCAssets} assets, ₹${totalAMCValue.toLocaleString()}`
              : "No calculations stored"}
          </p>
        </div>

        {/* Warranty Data Status */}
        <div
          style={{
            padding: "16px",
            backgroundColor: hasWarrantyData ? "#fef3c7" : "#f8fafc",
            border: `1px solid ${hasWarrantyData ? "#fcd34d" : "#e2e8f0"}`,
            borderRadius: "12px",
          }}
        >
          <div style={{ fontSize: "1.5rem", marginBottom: "8px" }}>
            {hasWarrantyData ? "🛡️" : "🔒"}
          </div>
          <h4
            style={{
              fontSize: "0.9rem",
              fontWeight: 600,
              marginBottom: "4px",
              color: "#1e293b",
            }}
          >
            Warranty Calculations
          </h4>
          <p style={{ fontSize: "0.8rem", color: "#64748b", margin: 0 }}>
            {hasWarrantyData
              ? `${totalWarrantyAssets} assets, ₹${totalWarrantyValue.toLocaleString()}`
              : "No calculations stored"}
          </p>
        </div>
      </div>

      {/* Excel Upload Area */}
      <div
        style={{
          border: `2px dashed ${dragActive ? "#3b82f6" : "#cbd5e1"}`,
          borderRadius: "12px",
          padding: "32px",
          textAlign: "center",
          backgroundColor: dragActive
            ? "#f0f9ff"
            : hasData
            ? "#f0fdf4"
            : "#f8fafc",
          transition: "all 0.3s ease",
          position: "relative",
        }}
        onDragEnter={handleDrag}
        onDragLeave={handleDrag}
        onDragOver={handleDrag}
        onDrop={handleDrop}
      >
        {isLoading ? (
          <div>
            <div style={{ fontSize: "2rem", marginBottom: "16px" }}>⏳</div>
            <p style={{ color: "#3b82f6", fontWeight: 600 }}>
              Processing Excel file...
            </p>
            <p style={{ fontSize: "0.8rem", color: "#64748b" }}>
              Please wait while we process your file
            </p>
          </div>
        ) : hasData ? (
          <div>
            <div style={{ fontSize: "2rem", marginBottom: "16px" }}>✅</div>
            <p
              style={{ color: "#059669", fontWeight: 600, marginBottom: "8px" }}
            >
              Excel File Loaded Successfully
            </p>
            <p
              style={{
                color: "#6b7280",
                fontSize: "0.9rem",
                marginBottom: "12px",
              }}
            >
              {fileName}
            </p>
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                gap: "24px",
                marginBottom: "16px",
                flexWrap: "wrap",
              }}
            >
              <div
                style={{
                  backgroundColor: "white",
                  padding: "12px 20px",
                  borderRadius: "8px",
                  border: "1px solid #e2e8f0",
                }}
              >
                <div
                  style={{
                    fontSize: "1.25rem",
                    fontWeight: 700,
                    color: "#3b82f6",
                  }}
                >
                  {sheets.length}
                </div>
                <div style={{ fontSize: "0.8rem", color: "#64748b" }}>
                  Sheets
                </div>
              </div>
              <div
                style={{
                  backgroundColor: "white",
                  padding: "12px 20px",
                  borderRadius: "8px",
                  border: "1px solid #e2e8f0",
                }}
              >
                <div
                  style={{
                    fontSize: "1.25rem",
                    fontWeight: 700,
                    color: "#059669",
                  }}
                >
                  {totalRows.toLocaleString()}
                </div>
                <div style={{ fontSize: "0.8rem", color: "#64748b" }}>
                  Total Rows
                </div>
              </div>
            </div>
            <button
              onClick={() => {
                document.getElementById("excel-file-input").click();
              }}
              style={{
                padding: "8px 16px",
                backgroundColor: "#3b82f6",
                color: "white",
                border: "none",
                borderRadius: "6px",
                cursor: "pointer",
                fontSize: "0.875rem",
              }}
            >
              Upload Different File
            </button>
          </div>
        ) : (
          <div>
            <div style={{ fontSize: "2rem", marginBottom: "16px" }}>📊</div>
            <p
              style={{
                fontSize: "1.1rem",
                fontWeight: 600,
                marginBottom: "8px",
                color: "#374151",
              }}
            >
              Upload Excel File
            </p>
            <p
              style={{
                color: "#6b7280",
                marginBottom: "16px",
                fontSize: "0.9rem",
              }}
            >
              Drag and drop your Excel file here, or click to browse
            </p>
            <button
              onClick={() => {
                document.getElementById("excel-file-input").click();
              }}
              style={{
                padding: "12px 24px",
                backgroundColor: "#3b82f6",
                color: "white",
                border: "none",
                borderRadius: "8px",
                cursor: "pointer",
                fontSize: "0.9rem",
                fontWeight: 600,
                transition: "all 0.2s ease",
              }}
            >
              Choose File
            </button>
          </div>
        )}

        <input
          id="excel-file-input"
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileInput}
          style={{ display: "none" }}
        />
      </div>

      {error && (
        <div
          style={{
            marginTop: "16px",
            padding: "12px",
            backgroundColor: "#fef2f2",
            border: "1px solid #fecaca",
            borderRadius: "8px",
            color: "#dc2626",
            fontSize: "0.875rem",
          }}
        >
          <strong>Error:</strong> {error}
        </div>
      )}

      {/* Data Flow Information */}
      {(hasData || hasAMCData || hasWarrantyData) && (
        <div
          style={{
            marginTop: "24px",
            padding: "16px",
            backgroundColor: "#f0f9ff",
            border: "1px solid #bfdbfe",
            borderRadius: "12px",
          }}
        >
          <h4
            style={{
              fontSize: "0.9rem",
              fontWeight: 600,
              marginBottom: "12px",
              color: "#1e293b",
            }}
          >
            🔄 Data Flow Status
          </h4>
          <div
            style={{ fontSize: "0.8rem", color: "#64748b", lineHeight: 1.6 }}
          >
            <p style={{ margin: "4px 0" }}>
              • <strong>Excel Files:</strong>{" "}
              {hasData
                ? "Available for AMC/Warranty calculations"
                : "Upload required"}
            </p>
            <p style={{ margin: "4px 0" }}>
              • <strong>AMC Calculations:</strong>{" "}
              {hasData
                ? "Available for AMC/Warranty calculations"
                : "Upload required"}
            </p>
            <p style={{ margin: "4px 0" }}>
              • <strong>Warranty Calculations:</strong>{" "}
              {hasData
                ? "Available for AMC/Warranty calculations"
                : "Upload required"}
            </p>
          </div>
        </div>
      )}
    </div>
  );
};

function EnhancedUIDAIDashboard() {
  const dispatch = useDispatch();
  const navigate = useNavigate();

  // Dashboard selectors
  const hoveredCard = useSelector(selectHoveredCard);
  const isMobile = useSelector(selectIsMobile);
  const currentTime = useSelector(selectCurrentTime);
  const searchTerm = useSelector(selectSearchTerm);
  const selectedCategory = useSelector(selectSelectedCategory);
  const stats = useSelector(selectStats);
  const showSearch = useSelector(selectShowSearch);
  const recentActivity = useSelector(selectRecentActivity);

  // Data selectors for global context
  const hasExcelData = useSelector(selectHasData);

  useEffect(() => {
    const savedExcel = localStorage.getItem("excelData");
    const savedAMC = localStorage.getItem("amcData");
    const savedWarranty = localStorage.getItem("warrantyData");

    if (savedExcel) dispatch(loadExcelFromMemory());
    if (savedAMC) dispatch(loadAMCFromMemory());
    if (savedWarranty) dispatch(loadWarrantyFromMemory());
  }, [dispatch]);

  // Initialize responsive behavior and real-time clock
  useEffect(() => {
    const handleResize = () => {
      dispatch(setIsMobile(window.innerWidth <= 768));
    };

    const updateTime = () => {
      const time = new Date().toLocaleTimeString("en-IN", {
        timeZone: "Asia/Kolkata",
        hour12: true,
        hour: "2-digit",
        minute: "2-digit",
      });
      dispatch(setCurrentTime(time));
    };

    // Initialize
    handleResize();
    updateTime();

    // Set up event listeners
    window.addEventListener("resize", handleResize);
    const timeInterval = setInterval(updateTime, 1000);

    // Simulate real-time stats updates
    const statsInterval = setInterval(() => {
      dispatch(
        updateStats({
          usersOnline: stats.usersOnline + Math.floor(Math.random() * 10) - 5,
          tasksCompleted: stats.tasksCompleted + Math.floor(Math.random() * 3),
        })
      );
    }, 30000);

    // Auto-rotate tips
    const tipsInterval = setInterval(() => {
      dispatch(nextTip());
    }, 5000);

    // Cleanup
    return () => {
      window.removeEventListener("resize", handleResize);
      clearInterval(timeInterval);
      clearInterval(statsInterval);
      clearInterval(tipsInterval);
    };
  }, [dispatch, stats.usersOnline, stats.tasksCompleted]);

  // Keyboard shortcuts
  useEffect(() => {
    const handleKeyPress = (e) => {
      if (e.ctrlKey && e.key === "k") {
        e.preventDefault();
        dispatch(setShowSearch(true));
      }
      if (e.key === "Escape") {
        dispatch(setShowSearch(false));
        dispatch(toggleNotifications());
      }
    };

    window.addEventListener("keydown", handleKeyPress);
    return () => window.removeEventListener("keydown", handleKeyPress);
  }, [dispatch]);

  // Navigation handler for tool cards
  const handleNavigation = useCallback(
    (path, toolName) => {
      try {
        navigate(path);
      } catch (error) {
        console.error(`❌ Navigation failed for ${toolName}:`, error);
      }
    },
    [navigate]
  );

  // Tool data for filtering - Enhanced with data availability indicators  
  const tools = [
    {
      id: "amc",
      title: "AMC Calculator",
      category: "financial",
      description:
        "No more manual AMC calculations — just upload your sheet, and let the tool handle everything!",
      icon: "📊",
      path: "/enhanced-amc-calculator",
      tag: "Financial Planning",
      dataStatus: hasExcelData ? "ready" : "needs-data",
      statusText: hasExcelData
        ? "Excel data available"
        : "Upload Excel file first",
    },
    {
      id: "warranty",
      title: "Warranty Estimator",
      category: "risk",
      description:
        "Your digital assistant for warranty tracking — smart calculator, export-ready, and easy to use.",
      icon: "🛡️",
      path: "/warranty-estimator",
      tag: "Risk Management",
      dataStatus: hasExcelData ? "ready" : "needs-data",
      statusText: hasExcelData
        ? "Excel data available"
        : "Upload Excel file first",
    },
    {
      id: "amc-payment-tracker",
      title: "AMC Payment Tracker",
      category: "financial",
      icon: "💳",
      description: "Stay ahead of deadlines with our AMC Payment Tracker — your smart, simple solution for hassle-free contract payments!",
      path: "/amc-payment-tracker",
      tag: "Financial Monitoring",
      dataStatus: hasExcelData ? "ready" : "needs-data",
      statusText: hasExcelData
        ? "Excel data available"
        : "Upload Excel file first",
    },
    {
      id: "warranty-payment-tracker",
      title: "Warranty Payment Tracker",
      category: "financial",
      icon: "💳",
      description: "Keep every warranty under control — track, manage, and never lose coverage with ease!",
      path: "/warranty-payment-tracker",
      tag: "Financial Monitoring",
      dataStatus: hasExcelData ? "ready" : "needs-data",
      statusText: hasExcelData
        ? "Excel data available"
        : "Upload Excel file first",
    },
  ];

  // Filter tools based on search and category using selector
  const filteredTools = useSelector((state) =>
    selectFilteredTools(state, tools)
  );

  // Event handlers using dispatch
  const handleMouseEnterCard = useCallback(
    (cardId) => {
      dispatch(setHoveredCard(cardId));
    },
    [dispatch]
  );

  const handleMouseLeaveCard = useCallback(() => {
    dispatch(setHoveredCard(null));
  }, [dispatch]);

  const handleSearchChange = useCallback(
      (e) => {
        dispatch(setSearchTerm(e.target.value));
      },
      [dispatch]
  );

  const handleCategoryChange = useCallback(
    (category) => {
      dispatch(setSelectedCategory(category));
    },
    [dispatch]
  );

  const containerStyle = {
    minHeight: "100vh",
    background: "linear-gradient(135deg, #f8fafc 0%, #e2e8f0 50%, #cbd5e1 100%)",
    color: "#1e293b",
    fontFamily:
      '"Inter", "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif',
    display: "flex",
    flexDirection: "column",
    position: "relative",
    transition: "all 0.3s ease",
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

   /* ============== NAVBAR STYLING ============= */

  const navbarStyle = {
    width: "97%",
    backgroundColor: "rgba(255, 255, 255, 0.98)",
    backdropFilter: "blur(20px)",
    borderBottom: "1px solid rgba(203, 213, 225, 0.5)",
    color: "#1e293b",
    padding: "16px 20px",
    position: "relative",
    zIndex: 10,
    boxShadow: "0 4px 25px -12px rgba(0, 0, 0, 0.25)",
  };

  // Search Modal Styles
  const searchModalStyle = {
    position: "fixed",
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: "rgba(0, 0, 0, 0.8)",
    display: showSearch ? "flex" : "none",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 1000,
  };

  const searchBoxStyle = {
    backgroundColor: "white",
    borderRadius: "16px",
    padding: "24px",
    width: "90%",
    maxWidth: "500px",
    boxShadow: "0 25px 50px -12px rgba(0, 0, 0, 0.5)",
  };

  const searchInputStyle = {
    width: "100%",
    padding: "16px",
    fontSize: "1.1rem",
    border: "2px solid #e2e8f0",
    borderRadius: "12px",
    outline: "none",
    marginBottom: "16px",
    backgroundColor: "white",
    color: "#1e293b"
  };

  const mainContentStyle = {
    flex: 1,
    padding: isMobile ? "20px" : "40px",
    overflowY: "auto",
    position: "relative",
    zIndex: 1,
  };

  const toolCardStyle = (tool) => ({
    background: hoveredCard === tool.id
        ? "linear-gradient(135deg, rgba(255, 255, 255, 0.98) 0%, rgba(248, 250, 252, 0.95) 100%)"
        : "rgba(255, 255, 255, 0.95)",
    backdropFilter: "blur(20px)",
    borderRadius: "20px",
    padding: "32px",
    border: `1px solid ${
      hoveredCard === tool.id ? "#3b82f6" : "rgba(255, 255, 255, 0.2)"
    }`,
    boxShadow:
      hoveredCard === tool.id
        ? "0 25px 50px -12px rgba(59, 130, 246, 0.25)"
        : "0 20px 40px -12px rgba(0, 0, 0, 0.15)",
    transform: hoveredCard === tool.id ? "translateY(-8px)" : "translateY(0)",
    transition: "all 0.3s cubic-bezier(0.4, 0, 0.2, 1)",
    cursor: "pointer",
    position: "relative",
    overflow: "hidden",
  });

  const toolIconStyle = {
    fontSize: "3rem",
    marginBottom: "20px",
    display: "block",
  };

  const toolTitleStyle = {
    fontSize: "1.5rem",
    fontWeight: 700,
    color: "#1e293b",
    marginBottom: "12px",
    lineHeight: 1.2,
  };

  const toolDescriptionStyle = {
    color: "#64748b",
    fontSize: "0.95rem",
    lineHeight: 1.6,
    marginBottom: "20px",
  };

  const toolTagStyle = {
    display: "inline-block",
    padding: "6px 12px",
    backgroundColor: "#f1f5f9",
    color: "#475569",
    borderRadius: "20px",
    fontSize: "0.8rem",
    fontWeight: 600,
    marginBottom: "16px",
  };

  const statusIndicatorStyle = (status) => ({
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "8px 12px",
    borderRadius: "8px",
    fontSize: "0.8rem",
    fontWeight: 500,
    backgroundColor: status === "ready" ? "#f0fdf4" : "#fef3c7",
    color: status === "ready" ? "#059669" : "#d97706",
    border: `1px solid ${status === "ready" ? "#bbf7d0" : "#fcd34d"}`,
  });

  const categoryFilterStyle = (category, isActive) => ({
    padding: "8px 16px",
    borderRadius: "20px",
    border: "none",
    backgroundColor: isActive ? "#3b82f6" : "rgba(255, 255, 255, 0.8)",
    color: isActive ? "white" : "#64748b",
    fontSize: "0.875rem",
    fontWeight: 600,
    cursor: "pointer",
    transition: "all 0.2s ease",
    backdropFilter: "blur(10px)",
  });

// RENDER COMPONENT 
  return (
    <div style={containerStyle}>
      <div style={backgroundOverlayStyle}></div>

      {/* Search Modal */}
      <div
        style={searchModalStyle}
        onClick={() => dispatch(setShowSearch(false))}
      >
        <div style={searchBoxStyle} onClick={(e) => e.stopPropagation()}>
          <input
            type="text"
            placeholder="Search tools and features..."
            value={searchTerm}
            onChange={handleSearchChange}
            style={searchInputStyle}
            autoFocus
          />
          <div
            style={{
              fontSize: "0.875rem",
              color: "#64748b",
            }}
          >
            Press <kbd>Escape</kbd> to close
          </div>
        </div>
      </div>

      {/* Navbar - Fixed for mobile responsiveness */}
      <div style={navbarStyle}>
        <div
          style={{
            display: "flex",
            flexDirection: isMobile ? "column" : "row",
            alignItems: isMobile ? "center" : "center",
            justifyContent: "space-between",
            gap: isMobile ? "12px" : "16px",
            width: "96%",
          }}
        >
          {/* Left Section - Brand with tri-color */}
          <div
            style={{
              display: "flex",
              flexDirection: isMobile ? "column" : "row",
              alignItems: "center",
              gap: isMobile ? "8px" : "24px",
            }}
          >
            <div style={{ textAlign: isMobile ? "center" : "left" }}>
              <h1 style={{ fontSize: isMobile ? "1.5rem" : "2rem", fontWeight: 800, margin: 0 }}>
              <span
                style={{
                  background: "linear-gradient(45deg, #FF6B35, #FFA500)",
                  WebkitBackgroundClip: "text",
                  WebkitTextFillColor: "transparent",
                  backgroundClip: "text",
                  color: "transparent",
                }}
              >
                SUVIDHA{" "}
              </span>
              <span
                style={{
                  background: "linear-gradient(45deg, #28A745, #28A745)", 
                  WebkitBackgroundClip: "text",
                  WebkitTextFillColor: "transparent",
                  backgroundClip: "text",
                  color: "transparent",
                }}
              >
                SETU
              </span>
            </h1>
              {!isMobile && (
                <p
                  style={{
                    color: "#64748b",
                    fontSize: "0.85rem",
                    margin: 0,
                    lineHeight: 1.4,
                    maxWidth: "300px",
                    marginTop: "4px",
                  }}
                >
                  Advanced financial calculation suite with intelligent automation
                </p>
              )}
            </div>
          </div>

          {/* Right Section - Controls in responsive layout */}
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: isMobile ? "8px" : "12px",
              flexWrap: "wrap",
              justifyContent: isMobile ? "center" : "flex-end",
            }}
          >
            {/* Current Time */}
            <div
              style={{
                fontSize: "0.75rem",
                color: "#64748b",
                fontWeight: 600,
                padding: "6px 10px",
                backgroundColor: "rgba(248, 250, 252, 0.8)",
                borderRadius: "6px",
                border: "1px solid rgba(203, 213, 225, 0.3)",
                whiteSpace: "nowrap",
              }}
            >
              🕒 {currentTime}
            </div>
          </div>
        </div>
      </div>  

      {/* Main Content */}
      <div style={mainContentStyle}>
        {/* Global Data Management Hub */}
        <ExcelUploadSection />

        {/* Category Filters */}
        <div
          style={{
            display: "flex",
            gap: "12px",
            marginBottom: "32px",
            flexWrap: "wrap",
          }}
        >
          {["all", "financial", "risk", "monitoring"].map((category) => (
            <button
              key={category}
              onClick={() => handleCategoryChange(category)}
              style={categoryFilterStyle(
                category,
                selectedCategory === category
              )}
            >
              {category === "all"
                ? "🔍 All Tools"
                : category === "financial"
                ? "💰 Financial"
                : category === "risk"
                ? "🛡️ Risk Management"
                : "📊 Monitoring"}
            </button>
          ))}
        </div>

        {/* Tools Grid */}
        <div
          style={{
            display: "grid",
            gridTemplateColumns: isMobile
              ? "1fr"
              : "repeat(auto-fit, minmax(400px, 1fr))",
            gap: "32px",
            marginBottom: "40px",
          }}
        >
          {filteredTools.map((tool) => (
            <div
              key={tool.id}
              style={toolCardStyle(tool)}
              onMouseEnter={() => handleMouseEnterCard(tool.id)}
              onMouseLeave={handleMouseLeaveCard}
              onClick={() => handleNavigation(tool.path, tool.title)}
            >
              {/* Tool Icon */}
              <div style={toolIconStyle}>{tool.icon}</div>

              {/* Tool Tag */}
              <div style={toolTagStyle}>{tool.tag}</div>

              {/* Tool Title */}
              <h3 style={toolTitleStyle}>{tool.title}</h3>

              {/* Tool Description */}
              <p style={toolDescriptionStyle}>{tool.description}</p>

              {/* Data Status */}
              <div style={statusIndicatorStyle(tool.dataStatus)}>
                <div
                  style={{
                    width: "8px",
                    height: "8px",
                    borderRadius: "50%",
                    backgroundColor: tool.dataStatus === "ready" ? "#10b981" : "#f59e0b",
                  }}
                ></div>
                {tool.statusText}
              </div>

              {/* Hover Effect Overlay */}
              {hoveredCard === tool.id && (
                <div
                  style={{
                    position: "absolute",
                    top: 0,
                    left: 0,
                    right: 0,
                    bottom: 0,
                    background: "linear-gradient(135deg, rgba(59, 130, 246, 0.05) 0%, rgba(139, 92, 246, 0.05) 100%)",
                    borderRadius: "20px",
                    pointerEvents: "none",
                  }}
                ></div>
              )}
            </div>
          ))}
        </div>

        {/* Recent Activity */}
        {recentActivity.length > 0 && (
          <div
            style={{
              background: "rgba(255, 255, 255, 0.95)",
              backdropFilter: "blur(20px)",
              borderRadius: "20px",
              padding: "32px",
              border: "1px solid rgba(255, 255, 255, 0.2)",
              boxShadow: "0 20px 40px -12px rgba(0, 0, 0, 0.15)",
            }}
          >
            <h3
              style={{
                fontSize: "1.5rem",
                fontWeight: 700,
                color: "#1e293b",
                marginBottom: "24px",
              }}
            >
              📈 Recent Activity
            </h3>
            <div style={{ space: "16px" }}>
              {recentActivity.map((activity, index) => (
                <div
                  key={index}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: "16px",
                    padding: "16px",
                    backgroundColor: "#f8fafc",
                    borderRadius: "12px",
                    marginBottom: "12px",
                  }}
                >
                  <div style={{ fontSize: "1.5rem" }}>{activity.icon}</div>
                  <div style={{ flex: 1 }}>
                    <div
                      style={{
                        fontSize: "0.9rem",
                        fontWeight: 600,
                        color: "#374151",
                        marginBottom: "4px",
                      }}
                    >
                      {activity.title}
                    </div>
                    <div
                      style={{
                        fontSize: "0.8rem",
                        color: "#64748b",
                      }}
                    >
                      {activity.timestamp}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
        {/* Footer */}
        <div
          style={{
            textAlign: "center",
            marginTop: "40px",
            padding: "20px",
            borderTop: "1px solid rgba(203, 213, 225, 0.3)",
          }}
        >
          <div
            style={{
              fontSize: "0.75rem",
              color: "#9ca3af",
            }}
          >
            Suvidha Setu v1.0 - Advanced Financial Calculation Suite
          </div>
        </div>
      </div>
    </div>
  );
}

export default EnhancedUIDAIDashboard;