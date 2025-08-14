import React, { useMemo, useState } from "react";
import { Card, Select, Row, Col, Statistic, InputNumber } from "antd";

const { Option } = Select;

// Utility function to format numbers in Indian numbering system
const formatIndianCurrency = (value, showDecimal = true) => {
  if (value < 100000) {
    // Less than 1 lakh - show in thousands
    return `₹${(value / 1000).toFixed(showDecimal ? 1 : 0)}K`;
  } else if (value < 10000000) {
    // 1 lakh to 1 crore - show in lakhs
    return `₹${(value / 100000).toFixed(showDecimal ? 2 : 0)}L`;
  } else {
    // Above 1 crore - show in crores
    return `₹${(value / 10000000).toFixed(showDecimal ? 2 : 0)}Cr`;
  }
};

// Enhanced Bar Chart Component with Interactive Hover Tooltips
const BarChart = ({ data, title, width = 800, height = 400 }) => {
  const [hoveredBar, setHoveredBar] = useState(null);
  const [mousePosition, setMousePosition] = useState({ x: 0, y: 0 });

  const margin = { top: 20, right: 30, bottom: 60, left: 80 };
  const chartWidth = width - margin.left - margin.right;
  const chartHeight = height - margin.top - margin.bottom;

  const maxValue = Math.max(...data.map((d) => d.value));
  const barWidth = (chartWidth / data.length) * 0.8;
  const barSpacing = (chartWidth / data.length) * 0.2;

  const handleBarMouseEnter = (index, event) => {
    const rect = event.currentTarget.closest('svg').getBoundingClientRect();
    setHoveredBar(index);
    setMousePosition({
      x: event.clientX - rect.left,
      y: event.clientY - rect.top
    });
  };

  const handleBarMouseMove = (event) => {
    if (hoveredBar !== null) {
      const rect = event.currentTarget.closest('svg').getBoundingClientRect();
      setMousePosition({
        x: event.clientX - rect.left,
        y: event.clientY - rect.top
      });
    }
  };

  const handleBarMouseLeave = () => {
    setHoveredBar(null);
  };

  return (
    <Card title={title} style={{ marginBottom: 24 }}>
      <div style={{ position: 'relative' }}>
        <svg 
          width={width} 
          height={height}
          onMouseMove={handleBarMouseMove}
          style={{ cursor: hoveredBar !== null ? 'pointer' : 'default' }}
        >
          {/* Y-axis */}
          <line
            x1={margin.left}
            y1={margin.top}
            x2={margin.left}
            y2={height - margin.bottom}
            stroke="#64748b"
            strokeWidth="1"
          />

          {/* X-axis */}
          <line
            x1={margin.left}
            y1={height - margin.bottom}
            x2={width - margin.right}
            y2={height - margin.bottom}
            stroke="#64748b"
            strokeWidth="1"
          />

          {/* Bars */}
          {data.map((item, index) => {
            const barHeight = (item.value / maxValue) * chartHeight;
            const x = margin.left + index * (barWidth + barSpacing) + barSpacing / 2;
            const y = height - margin.bottom - barHeight;
            const isHovered = hoveredBar === index;

            return (
              <g key={index}>
                <rect
                  x={x}
                  y={y}
                  width={barWidth}
                  height={barHeight}
                  fill={isHovered ? "#065f46" : "#10b981"}
                  stroke={isHovered ? "#022c22" : "#059669"}
                  strokeWidth={isHovered ? "2" : "1"}
                  opacity={isHovered ? "1" : "0.8"}
                  style={{
                    transition: 'all 0.2s ease-in-out',
                    filter: isHovered ? 'drop-shadow(0 4px 8px rgba(0,0,0,0.2))' : 'none'
                  }}
                  onMouseEnter={(e) => handleBarMouseEnter(index, e)}
                  onMouseLeave={handleBarMouseLeave}
                />

                {/* Hover overlay for better interaction */}
                <rect
                  x={x - 5}
                  y={margin.top}
                  width={barWidth + 10}
                  height={chartHeight}
                  fill="transparent"
                  style={{ cursor: 'pointer' }}
                  onMouseEnter={(e) => handleBarMouseEnter(index, e)}
                  onMouseLeave={handleBarMouseLeave}
                />

                {/* Value labels on bars */}
                <text
                  x={x + barWidth / 2}
                  y={y - 5}
                  textAnchor="middle"
                  fontSize="10"
                  fill="#374151"
                  fontWeight={isHovered ? "600" : "400"}
                >
                  {formatIndianCurrency(item.value, false)}
                </text>

                {/* X-axis labels */}
                <text
                  x={x + barWidth / 2}
                  y={height - margin.bottom + 15}
                  textAnchor="middle"
                  fontSize="11"
                  fill={isHovered ? "#10b981" : "#64748b"}
                  fontWeight={isHovered ? "600" : "400"}
                  transform={`rotate(-45, ${x + barWidth / 2}, ${height - margin.bottom + 15})`}
                >
                  {item.label}
                </text>
              </g>
            );
          })}

          {/* Y-axis labels */}
          {[0, 0.25, 0.5, 0.75, 1].map((ratio, index) => {
            const y = height - margin.bottom - ratio * chartHeight;
            const value = ratio * maxValue;

            return (
              <g key={index}>
                <line
                  x1={margin.left - 5}
                  y1={y}
                  x2={margin.left}
                  y2={y}
                  stroke="#64748b"
                  strokeWidth="1"
                />
                <text
                  x={margin.left - 10}
                  y={y + 3}
                  textAnchor="end"
                  fontSize="10"
                  fill="#64748b"
                >
                  {formatIndianCurrency(value, false)}
                </text>
              </g>
            );
          })}
        </svg>

        {/* Enhanced Tooltip */}
        {hoveredBar !== null && (
          <div
            style={{
              position: 'absolute',
              left: Math.min(mousePosition.x + 10, width - 220),
              top: Math.max(mousePosition.y - 80, 10),
              backgroundColor: '#064e3b',
              color: 'white',
              padding: '12px 16px',
              borderRadius: '8px',
              boxShadow: '0 10px 25px rgba(0,0,0,0.3)',
              fontSize: '14px',
              fontWeight: '500',
              zIndex: 1000,
              pointerEvents: 'none',
              minWidth: '200px',
              border: '1px solid #065f46'
            }}
          >
            <div style={{ marginBottom: '8px', fontWeight: '600', color: '#34d399' }}>
              🛡️ {data[hoveredBar].label}
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <span style={{ color: '#d1fae5' }}>Warranty Value:</span>
              <span style={{ fontWeight: '600', color: '#10b981' }}>
                {formatIndianCurrency(data[hoveredBar].value, true)}
              </span>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '4px' }}>
              <span style={{ color: '#d1fae5' }}>Exact Amount:</span>
              <span style={{ fontWeight: '600', color: '#34d399' }}>
                ₹{data[hoveredBar].value.toLocaleString('en-IN', { maximumFractionDigits: 2 })}
              </span>
            </div>
            {(() => {
              const totalValue = data.reduce((sum, item) => sum + item.value, 0);
              const percentage = ((data[hoveredBar].value / totalValue) * 100).toFixed(1);
              return (
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '4px' }}>
                  <span style={{ color: '#d1fae5' }}>% of Total:</span>
                  <span style={{ fontWeight: '600', color: '#6ee7b7' }}>
                    {percentage}%
                  </span>
                </div>
              );
            })()}
          </div>
        )}
      </div>
    </Card>
  );
};

// Enhanced Line Chart Component with Interactive Hover Tooltips
const LineChart = ({ data, title, width = 800, height = 400 }) => {
  const [hoveredPoint, setHoveredPoint] = useState(null);
  const [mousePosition, setMousePosition] = useState({ x: 0, y: 0 });

  const margin = { top: 20, right: 30, bottom: 60, left: 80 };
  const chartWidth = width - margin.left - margin.right;
  const chartHeight = height - margin.top - margin.bottom;

  const maxValue = Math.max(...data.map((d) => d.value));
  const minValue = Math.min(...data.map((d) => d.value));
  const valueRange = maxValue - minValue;

  const pathData = data
    .map((item, index) => {
      const x = margin.left + (index / (data.length - 1)) * chartWidth;
      const y = height - margin.bottom - ((item.value - minValue) / valueRange) * chartHeight;
      return `${index === 0 ? "M" : "L"} ${x} ${y}`;
    })
    .join(" ");

  const handlePointMouseEnter = (index, event) => {
    const rect = event.currentTarget.closest('svg').getBoundingClientRect();
    setHoveredPoint(index);
    setMousePosition({
      x: event.clientX - rect.left,
      y: event.clientY - rect.top
    });
  };

  const handlePointMouseMove = (event) => {
    if (hoveredPoint !== null) {
      const rect = event.currentTarget.closest('svg').getBoundingClientRect();
      setMousePosition({
        x: event.clientX - rect.left,
        y: event.clientY - rect.top
      });
    }
  };

  const handlePointMouseLeave = () => {
    setHoveredPoint(null);
  };

  return (
    <Card title={title} style={{ marginBottom: 24 }}>
      <div style={{ position: 'relative' }}>
        <svg 
          width={width} 
          height={height}
          onMouseMove={handlePointMouseMove}
        >
          {/* Y-axis */}
          <line
            x1={margin.left}
            y1={margin.top}
            x2={margin.left}
            y2={height - margin.bottom}
            stroke="#64748b"
            strokeWidth="1"
          />

          {/* X-axis */}
          <line
            x1={margin.left}
            y1={height - margin.bottom}
            x2={width - margin.right}
            y2={height - margin.bottom}
            stroke="#64748b"
            strokeWidth="1"
          />

          {/* Line */}
          <path 
            d={pathData} 
            fill="none" 
            stroke="#10b981" 
            strokeWidth="3"
            style={{
              filter: hoveredPoint !== null ? 'drop-shadow(0 2px 4px rgba(16,185,129,0.3))' : 'none',
              transition: 'all 0.2s ease-in-out'
            }}
          />

          {/* Data points */}
          {data.map((item, index) => {
            const x = margin.left + (index / (data.length - 1)) * chartWidth;
            const y = height - margin.bottom - ((item.value - minValue) / valueRange) * chartHeight;
            const isHovered = hoveredPoint === index;

            return (
              <g key={index}>
                <circle
                  cx={x}
                  cy={y}
                  r={isHovered ? "6" : "4"}
                  fill={isHovered ? "#065f46" : "#059669"}
                  stroke="white"
                  strokeWidth={isHovered ? "3" : "2"}
                  style={{
                    cursor: 'pointer',
                    transition: 'all 0.2s ease-in-out',
                    filter: isHovered ? 'drop-shadow(0 4px 8px rgba(0,0,0,0.2))' : 'none'
                  }}
                  onMouseEnter={(e) => handlePointMouseEnter(index, e)}
                  onMouseLeave={handlePointMouseLeave}
                />

                {/* Invisible larger circle for better hover area */}
                <circle
                  cx={x}
                  cy={y}
                  r="12"
                  fill="transparent"
                  style={{ cursor: 'pointer' }}
                  onMouseEnter={(e) => handlePointMouseEnter(index, e)}
                  onMouseLeave={handlePointMouseLeave}
                />

                {/* X-axis labels */}
                <text
                  x={x}
                  y={height - margin.bottom + 15}
                  textAnchor="middle"
                  fontSize="11"
                  fill={isHovered ? "#059669" : "#64748b"}
                  fontWeight={isHovered ? "600" : "400"}
                  transform={`rotate(-45, ${x}, ${height - margin.bottom + 15})`}
                >
                  {item.label}
                </text>
              </g>
            );
          })}

          {/* Y-axis labels */}
          {[0, 0.25, 0.5, 0.75, 1].map((ratio, index) => {
            const y = height - margin.bottom - ratio * chartHeight;
            const value = minValue + ratio * valueRange;

            return (
              <g key={index}>
                <line
                  x1={margin.left - 5}
                  y1={y}
                  x2={margin.left}
                  y2={y}
                  stroke="#64748b"
                  strokeWidth="1"
                />
                <text
                  x={margin.left - 10}
                  y={y + 3}
                  textAnchor="end"
                  fontSize="10"
                  fill="#64748b"
                >
                  {formatIndianCurrency(value, false)}
                </text>
              </g>
            );
          })}
        </svg>

        {/* Tooltip for Line Chart */}
        {hoveredPoint !== null && (
          <div
            style={{
              position: 'absolute',
              left: Math.min(mousePosition.x + 10, width - 220),
              top: Math.max(mousePosition.y - 80, 10),
              backgroundColor: '#064e3b',
              color: 'white',
              padding: '12px 16px',
              borderRadius: '8px',
              boxShadow: '0 10px 25px rgba(0,0,0,0.3)',
              fontSize: '14px',
              fontWeight: '500',
              zIndex: 1000,
              pointerEvents: 'none',
              minWidth: '200px',
              border: '1px solid #065f46'
            }}
          >
            <div style={{ marginBottom: '8px', fontWeight: '600', color: '#34d399' }}>
              📈 {data[hoveredPoint].label}
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <span style={{ color: '#d1fae5' }}>Warranty Value:</span>
              <span style={{ fontWeight: '600', color: '#10b981' }}>
                {formatIndianCurrency(data[hoveredPoint].value, true)}
              </span>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '4px' }}>
              <span style={{ color: '#d1fae5' }}>Exact Amount:</span>
              <span style={{ fontWeight: '600', color: '#34d399' }}>
                ₹{data[hoveredPoint].value.toLocaleString('en-IN', { maximumFractionDigits: 2 })}
              </span>
            </div>
            {/* Calculate trend if not first point */}
            {hoveredPoint > 0 && (() => {
              const prevValue = data[hoveredPoint - 1].value;
              const currentValue = data[hoveredPoint].value;
              const change = currentValue - prevValue;
              const changePercent = ((change / prevValue) * 100).toFixed(1);
              const isPositive = change >= 0;
              
              return (
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '4px' }}>
                  <span style={{ color: '#d1fae5' }}>Trend:</span>
                  <span style={{ 
                    fontWeight: '600', 
                    color: isPositive ? '#34d399' : '#fca5a5' 
                  }}>
                    {isPositive ? '↗' : '↘'} {changePercent}%
                  </span>
                </div>
              );
            })()}
          </div>
        )}
      </div>
    </Card>
  );
};

// Enhanced Pie Chart Component with Interactive Hover Tooltips
const PieChart = ({ data, title, width = 400, height = 400 }) => {
  const [hoveredSegment, setHoveredSegment] = useState(null);
  const [mousePosition, setMousePosition] = useState({ x: 0, y: 0 });

  const radius = Math.min(width, height) / 2 - 40;
  const centerX = width / 2;
  const centerY = height / 2;

  const total = data.reduce((sum, item) => sum + item.value, 0);
  let currentAngle = -Math.PI / 2; // Start at top

  const colors = [
    "#10b981", "#059669", "#047857", "#065f46", "#064e3b", 
    "#34d399", "#6ee7b7", "#a7f3d0", "#d1fae5", "#ecfdf5"
  ];

  const handleSegmentMouseEnter = (index, event) => {
    const rect = event.currentTarget.closest('svg').getBoundingClientRect();
    setHoveredSegment(index);
    setMousePosition({
      x: event.clientX - rect.left,
      y: event.clientY - rect.top
    });
  };

  const handleSegmentMouseMove = (event) => {
    if (hoveredSegment !== null) {
      const rect = event.currentTarget.closest('svg').getBoundingClientRect();
      setMousePosition({
        x: event.clientX - rect.left,
        y: event.clientY - rect.top
      });
    }
  };

  const handleSegmentMouseLeave = () => {
    setHoveredSegment(null);
  };

  return (
    <Card title={title} style={{ marginBottom: 24 }}>
      <div style={{ display: "flex", alignItems: "center", position: 'relative', justifyContent: 'center' }}>
        <svg 
          width={width} 
          height={height}
          onMouseMove={handleSegmentMouseMove}
        >
          {data.map((item, index) => {
            const angle = (item.value / total) * 2 * Math.PI;
            const startAngle = currentAngle;
            const endAngle = currentAngle + angle;
            const isHovered = hoveredSegment === index;

            const x1 = centerX + radius * Math.cos(startAngle);
            const y1 = centerY + radius * Math.sin(startAngle);
            const x2 = centerX + radius * Math.cos(endAngle);
            const y2 = centerY + radius * Math.sin(endAngle);

            const largeArcFlag = angle > Math.PI ? 1 : 0;

            const pathData = [
              `M ${centerX} ${centerY}`,
              `L ${x1} ${y1}`,
              `A ${radius} ${radius} 0 ${largeArcFlag} 1 ${x2} ${y2}`,
              "Z",
            ].join(" ");

            currentAngle += angle;

            // Label position
            const labelAngle = startAngle + angle / 2;
            const labelX = centerX + radius * 0.7 * Math.cos(labelAngle);
            const labelY = centerY + radius * 0.7 * Math.sin(labelAngle);

            return (
              <g key={index}>
                <path
                  d={pathData}
                  fill={colors[index % colors.length]}
                  stroke="white"
                  strokeWidth="2"
                  opacity={isHovered ? "1" : "0.8"}
                  style={{
                    cursor: 'pointer',
                    transition: 'all 0.2s ease-in-out',
                    filter: isHovered ? 'drop-shadow(0 4px 8px rgba(0,0,0,0.2))' : 'none',
                    transform: isHovered ? 'scale(1.05)' : 'scale(1)',
                    transformOrigin: `${centerX}px ${centerY}px`
                  }}
                  onMouseEnter={(e) => handleSegmentMouseEnter(index, e)}
                  onMouseLeave={handleSegmentMouseLeave}
                />
                {parseFloat(((item.value / total) * 100).toFixed(1)) > 5 && (
                  <text
                    x={labelX}
                    y={labelY}
                    textAnchor="middle"
                    fontSize="12"
                    fill="white"
                    fontWeight="600"
                    style={{ pointerEvents: 'none' }}
                  >
                    {((item.value / total) * 100).toFixed(1)}%
                  </text>
                )}
              </g>
            );
          })}
        </svg>

        {/* Legend */}
        <div style={{ marginLeft: 20 }}>
          {data.map((item, index) => (
            <div
              key={index}
              style={{ 
                display: "flex", 
                alignItems: "center", 
                marginBottom: 8,
                padding: '4px 8px',
                borderRadius: '4px',
                backgroundColor: hoveredSegment === index ? '#f3f4f6' : 'transparent',
                transition: 'all 0.2s ease-in-out',
                cursor: 'pointer'
              }}
              onMouseEnter={() => setHoveredSegment(index)}
              onMouseLeave={() => setHoveredSegment(null)}
            >
              <div
                style={{
                  width: 16,
                  height: 16,
                  backgroundColor: colors[index % colors.length],
                  marginRight: 8,
                  borderRadius: 2,
                  transition: 'all 0.2s ease-in-out',
                  transform: hoveredSegment === index ? 'scale(1.1)' : 'scale(1)'
                }}
              />
              <span style={{ 
                fontSize: "0.875rem",
                fontWeight: hoveredSegment === index ? '600' : '400'
              }}>
                {item.label}: {formatIndianCurrency(item.value)}
              </span>
            </div>
          ))}
        </div>

        {/* Pie Chart Tooltip */}
        {hoveredSegment !== null && (
          <div
            style={{
              position: 'absolute',
              left: Math.min(mousePosition.x + 10, width + 200 - 220),
              top: Math.max(mousePosition.y - 80, 10),
              backgroundColor: colors[hoveredSegment % colors.length],
              color: 'white',
              padding: '12px 16px',
              borderRadius: '8px',
              boxShadow: '0 10px 25px rgba(0,0,0,0.3)',
              fontSize: '14px',
              fontWeight: '500',
              zIndex: 1000,
              pointerEvents: 'none',
              minWidth: '200px',
              border: '2px solid white'
            }}
          >
            <div style={{ marginBottom: '8px', fontWeight: '600' }}>
              🛡️ {data[hoveredSegment].label}
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <span>Warranty Value:</span>
              <span style={{ fontWeight: '600' }}>
                {formatIndianCurrency(data[hoveredSegment].value, true)}
              </span>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '4px' }}>
              <span>Exact Amount:</span>
              <span style={{ fontWeight: '600' }}>
                ₹{data[hoveredSegment].value.toLocaleString('en-IN', { maximumFractionDigits: 2 })}
              </span>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '4px' }}>
              <span>Percentage:</span>
              <span style={{ fontWeight: '600' }}>
                {((data[hoveredSegment].value / total) * 100).toFixed(2)}%
              </span>
            </div>
          </div>
        )}
      </div>
    </Card>
  );
};

// Main Warranty Charts View Component
const WarrantyChartsView = ({ data, rawResults }) => {
  const [selectedChart, setSelectedChart] = useState("overview");
  const [selectedYear, setSelectedYear] = useState("all");
  const [quartersToShow, setQuartersToShow] = useState(8); // New state for number of quarters

  // Process data for charts
  const chartData = useMemo(() => {
    if (!data || data.length === 0) return null;

    // Extract quarterly totals from the data
    const quarterTotals = {};
    const yearTotals = {};

    // Process all products except Grand Total
    const productRows = data.filter((row) => row.itemName !== "Grand Total");

    productRows.forEach((product) => {
      Object.keys(product).forEach((key) => {
        // Process quarter columns (e.g., "JFM 2024", "AMJ 2024")
        if (key.match(/^(JFM|AMJ|JAS|OND) \d{4}$/)) {
          const value = product[key] || 0;
          if (!quarterTotals[key]) {
            quarterTotals[key] = 0;
          }
          quarterTotals[key] += value;

          // Extract year for yearly totals
          const year = key.split(" ")[1];
          if (!yearTotals[year]) {
            yearTotals[year] = 0;
          }
          yearTotals[year] += value;
        }
      });
    });

    // Quarterly totals chart data
    const allQuarterlyData = Object.entries(quarterTotals)
      .sort(([a], [b]) => {
        const [qA, yearA] = a.split(" ");
        const [qB, yearB] = b.split(" ");
        if (parseInt(yearA) !== parseInt(yearB)) {
          return parseInt(yearA) - parseInt(yearB);
        }
        const quarterOrder = { JFM: 0, AMJ: 1, JAS: 2, OND: 3 };
        return quarterOrder[qA] - quarterOrder[qB];
      })
      .map(([quarter, value]) => ({
        label: quarter,
        value: value || 0,
      }));

    // Filter quarterly data based on selected year
    const quarterlyData =
      selectedYear === "all"
        ? allQuarterlyData
        : allQuarterlyData.filter((q) => q.label.includes(selectedYear));

    // Yearly totals chart data
    const yearlyData = Object.entries(yearTotals)
      .sort(([a], [b]) => parseInt(a) - parseInt(b))
      .map(([year, value]) => ({
        label: year,
        value: value || 0,
      }));

    // Top products by warranty value
    const productData = productRows
      .map((product) => {
        // Find the total column for each product
        const totalColumn = Object.keys(product).find(
          (key) => key.includes("Total (") && key.includes(" Years)")
        );
        return {
          label: product.itemName,
          value: totalColumn ? product[totalColumn] || 0 : 0,
        };
      })
      .sort((a, b) => b.value - a.value)
      .slice(0, 10); // Top 10 products

    // Quarter-over-quarter growth (based on filtered data)
    const growthData = quarterlyData.map((quarter, index) => {
      if (index === 0) return { ...quarter, growth: 0 };
      const prevValue = quarterlyData[index - 1].value;
      const growth =
        prevValue > 0 ? ((quarter.value - prevValue) / prevValue) * 100 : 0;
      return { ...quarter, growth };
    });

    // Extract available years
    const availableYears = ["all", ...yearlyData.map((year) => year.label)];

    return {
      quarterly: quarterlyData,
      yearly: yearlyData,
      products: productData,
      growth: growthData,
      availableYears: availableYears,
    };
  }, [data, selectedYear]);

  if (!chartData) {
    return (
      <div style={{ textAlign: "center", padding: "60px 20px" }}>
        <div style={{ fontSize: "3rem", marginBottom: "16px" }}>📊</div>
        <h3 style={{ color: "#64748b", marginBottom: "8px" }}>
          No Data Available
        </h3>
        <p style={{ color: "#94a3b8" }}>
          Please calculate warranty schedules first to view charts and
          analytics.
        </p>
      </div>
    );
  }

  // Summary statistics (based on filtered quarterly data and quarters to show)
  const displayedQuarters = chartData.quarterly.slice(0, quartersToShow);
  const totalWarranty = displayedQuarters.reduce((sum, q) => sum + q.value, 0);
  const avgQuarterly = displayedQuarters.length > 0 ? totalWarranty / displayedQuarters.length : 0;
  const highestQuarter = displayedQuarters.length > 0
    ? displayedQuarters.reduce(
        (max, q) => (q.value > max.value ? q : max),
        displayedQuarters[0]
      )
    : { label: "N/A", value: 0 };

  return (
    <div style={{ padding: "0 20px" }}>
      {/* Chart Selection */}
      <div
        style={{
          marginBottom: 24,
          display: "flex",
          alignItems: "center",
          gap: 16,
          flexWrap: "wrap",
        }}
      >
        <span style={{ fontWeight: 600, color: "#374151" }}>View:</span>
        <Select
          value={selectedChart}
          onChange={setSelectedChart}
          style={{ width: 200 }}
        >
          <Option value="overview">📊 Overview</Option>
          <Option value="quarterly">📈 Quarterly Trends</Option>
          <Option value="yearly">📅 Yearly Distribution</Option>
          <Option value="products">🏆 Top Products</Option>
        </Select>

        {/* Year Filter - Show only for quarterly charts */}
        {(selectedChart === "quarterly" || selectedChart === "overview") && (
          <>
            <span style={{ fontWeight: 600, color: "#374151" }}>Year:</span>
            <Select
              value={selectedYear}
              onChange={setSelectedYear}
              style={{ width: 150 }}
            >
              <Option value="all">🗓️ All Years</Option>
              {chartData.availableYears
                .filter((year) => year !== "all")
                .map((year) => (
                  <Option key={year} value={year}>
                    📅 {year}
                  </Option>
                ))}
            </Select>
          </>
        )}

        {/* Quarters to Show Filter - Show only for quarterly and overview charts */}
        {(selectedChart === "quarterly" || selectedChart === "overview") && (
          <>
            <span style={{ fontWeight: 600, color: "#374151" }}>Show Quarters:</span>
            <InputNumber
              min={1}
              max={chartData.quarterly.length}
              value={quartersToShow}
              onChange={setQuartersToShow}
              style={{ width: 100 }}
              placeholder="Quarters"
            />
            <span style={{ color: "#64748b", fontSize: "0.875rem" }}>
              (Max: {chartData.quarterly.length})
            </span>
          </>
        )}
      </div>

      {/* Summary Statistics */}
      <Row gutter={16} style={{ marginBottom: 24 }}>
        <Col span={6}>
          <Card>
            <Statistic
              title={
                selectedYear === "all"
                  ? `Total Warranty Value (${displayedQuarters.length} quarters)`
                  : `Total Warranty Value (${selectedYear} - ${displayedQuarters.length} quarters)`
              }
              value={totalWarranty}
              formatter={(value) => formatIndianCurrency(value)}
              valueStyle={{ color: "#10b981" }}
            />
          </Card>
        </Col>
        <Col span={6}>
          <Card>
            <Statistic
              title={
                selectedYear === "all"
                  ? "Average per Quarter"
                  : `Avg per Quarter (${selectedYear})`
              }
              value={avgQuarterly}
              formatter={(value) => formatIndianCurrency(value)}
              valueStyle={{ color: "#059669" }}
            />
          </Card>
        </Col>
        <Col span={6}>
          <Card>
            <Statistic
              title={
                selectedYear === "all"
                  ? "Highest Quarter"
                  : `Highest Quarter (${selectedYear})`
              }
              value={highestQuarter.label}
              suffix={formatIndianCurrency(highestQuarter.value)}
              valueStyle={{ color: "#047857" }}
            />
          </Card>
        </Col>
        <Col span={6}>
          <Card>
            <Statistic
              title={
                selectedYear === "all"
                  ? "Displaying Quarters"
                  : `Quarters in ${selectedYear}`
              }
              value={displayedQuarters.length}
              suffix="quarters"
              valueStyle={{ color: "#065f46" }}
            />
          </Card>
        </Col>
      </Row>

      {/* Charts */}
      {selectedChart === "overview" && (
        <div>
          {/* Bar Chart - Full Width */}
          <Row gutter={16}>
            <Col span={24}>
              <BarChart
                data={displayedQuarters}
                title={
                  selectedYear === "all"
                    ? `Quarterly Warranty Revenue (Showing ${displayedQuarters.length} quarters)`
                    : `Quarterly Warranty Revenue - ${selectedYear} (Showing ${displayedQuarters.length} quarters)`
                }
                width={1000}
                height={400}
              />
            </Col>
          </Row>
          
          {/* Pie Chart - Full Width */}
          <Row gutter={16}>
            <Col span={24}>
              <div style={{ display: 'flex', justifyContent: 'center' }}>
                <PieChart
                  data={chartData.yearly}
                  title="Warranty Distribution by Year"
                  width={600}
                  height={400}
                />
              </div>
            </Col>
          </Row>
        </div>
      )}

      {selectedChart === "quarterly" && (
        <Row gutter={16}>
          <Col span={24}>
            <BarChart
              data={displayedQuarters}
              title={
                selectedYear === "all"
                  ? `Quarterly Warranty Schedule (Showing ${displayedQuarters.length} of ${chartData.quarterly.length} quarters)`
                  : `Quarterly Warranty Schedule - ${selectedYear} (Showing ${displayedQuarters.length} quarters)`
              }
              width={1000}
              height={400}
            />
          </Col>
          <Col span={24}>
            <LineChart
              data={displayedQuarters}
              title={
                selectedYear === "all"
                  ? `Quarterly Warranty Trends (Showing ${displayedQuarters.length} quarters)`
                  : `Quarterly Warranty Trends - ${selectedYear} (Showing ${displayedQuarters.length} quarters)`
              }
              width={1000}
              height={350}
            />
          </Col>
        </Row>
      )}

      {selectedChart === "yearly" && (
        <div>
          {/* Bar Chart - Full Width */}
          <Row gutter={16}>
            <Col span={24}>
              <BarChart
                data={chartData.yearly}
                title="Yearly Warranty Totals"
                width={1000}
                height={400}
              />
            </Col>
          </Row>
          
          {/* Pie Chart - Full Width */}
          <Row gutter={16}>
            <Col span={24}>
              <div style={{ display: 'flex', justifyContent: 'center' }}>
                <PieChart
                  data={chartData.yearly}
                  title="Year-wise Warranty Distribution"
                  width={600}
                  height={400}
                />
              </div>
            </Col>
          </Row>
        </div>
      )}

      {selectedChart === "products" && (
        <Row gutter={16}>
          <Col span={24}>
            <BarChart
              data={chartData.products}
              title="Top 10 Products by Warranty Value"
              width={1000}
              height={450}
            />
          </Col>
        </Row>
      )}
    </div>
  );
};

export default WarrantyChartsView;