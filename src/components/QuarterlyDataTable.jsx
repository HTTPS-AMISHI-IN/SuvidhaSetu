// Quarterly Data Table for AMC Schedule - Shows quarters as columns
import React, { useMemo, useState, useCallback } from "react";
import { Card, Button, Space, Typography, Tag, Input, Select } from "antd";
import { ExportOutlined, SearchOutlined } from "@ant-design/icons";

const { Text, Title } = Typography;
const { Option } = Select;

// Generate quarter headers for the next few years
const generateQuarterHeaders = (startYear, numberOfYears = 6) => {
  const quarters = ["JFM", "AMJ", "JAS", "OND"];
  const headers = [];

  for (let year = startYear; year < startYear + numberOfYears; year++) {
    quarters.forEach((quarter) => {
      headers.push(`${quarter}-${year}`);
    });
  }

  return headers;
};

// Transform AMC data to quarterly format
const transformToQuarterlyFormat = (amcResults) => {
  if (!amcResults || amcResults.length === 0) return [];

  return amcResults.map((product) => {
    const row = {
      id: product.id,
      uatDate: product.uatDate,
      productName: product.productName,
      location: product.location,
      invoiceValue: product.invoiceValue,
      quantity: product.quantity,
      totalAmcValue: product.totalAmcValue,
      totalAmountWithGST: product.totalAmountWithGST,
    };

    // Add quarterly data as individual columns
    if (product.quarters && Array.isArray(product.quarters)) {
      product.quarters.forEach((quarter) => {
        const quarterKey = `${quarter.quarter}-${quarter.year}`;
        row[quarterKey] = quarter.totalAmount || 0;
      });
    }

    return row;
  });
};

// Get all unique quarters from the data
const getUniqueQuarters = (transformedData) => {
  const quarterSet = new Set();

  transformedData.forEach((row) => {
    Object.keys(row).forEach((key) => {
      if (key.match(/^(JFM|AMJ|JAS|OND)-\d{4}$/)) {
        quarterSet.add(key);
      }
    });
  });

  // Sort quarters chronologically
  return Array.from(quarterSet).sort((a, b) => {
    const [quarterA, yearA] = a.split("-");
    const [quarterB, yearB] = b.split("-");

    if (yearA !== yearB) {
      return parseInt(yearA) - parseInt(yearB);
    }

    const quarterOrder = { JFM: 0, AMJ: 1, JAS: 2, OND: 3 };
    return quarterOrder[quarterA] - quarterOrder[quarterB];
  });
};

const QuarterlyDataTable = ({
  data = [],
  title = "Quarterly AMC Schedule",
  onExport,
  searchable = true,
  filterable = true,
  showWithoutGST = false,
  selectedLocation = "All Locations",
  height = 600,
}) => {
  const [searchTerm, setSearchTerm] = useState("");
  const [selectedLocationFilter, setSelectedLocationFilter] =
    useState("All Locations");

  // Transform data to quarterly format
  const transformedData = useMemo(() => {
    return transformToQuarterlyFormat(data);
  }, [data]);

  // Get unique quarters from data
  const quarters = useMemo(() => {
    return getUniqueQuarters(transformedData);
  }, [transformedData]);

  // Filter data based on search and location
  const filteredData = useMemo(() => {
    let filtered = transformedData;

    if (searchTerm) {
      const searchLower = searchTerm.toLowerCase();
      filtered = filtered.filter(
        (row) =>
          row.productName?.toLowerCase().includes(searchLower) ||
          row.location?.toLowerCase().includes(searchLower) ||
          row.uatDate?.toLowerCase().includes(searchLower)
      );
    }

    if (selectedLocationFilter !== "All Locations") {
      filtered = filtered.filter(
        (row) => row.location === selectedLocationFilter
      );
    }

    return filtered;
  }, [transformedData, searchTerm, selectedLocationFilter]);

  // Get unique locations for filter
  const locations = useMemo(() => {
    const locationSet = new Set(
      transformedData.map((row) => row.location).filter(Boolean)
    );
    return ["All Locations", ...Array.from(locationSet)];
  }, [transformedData]);

  // Calculate column totals
  const columnTotals = useMemo(() => {
    const totals = {};

    quarters.forEach((quarter) => {
      totals[quarter] = filteredData.reduce((sum, row) => {
        return sum + (row[quarter] || 0);
      }, 0);
    });

    return totals;
  }, [filteredData, quarters]);

  // Handle export
  const handleExport = useCallback(async () => {
    if (onExport) {
      // Create export data in the format shown in the image
      const exportData = filteredData.map((row) => {
        const exportRow = {
          "UAT Date": row.uatDate,
          "Product Name": row.productName,
          Location: row.location,
          "Invoice Value": row.invoiceValue,
          Quantity: row.quantity,
        };

        // Add quarterly columns
        quarters.forEach((quarter) => {
          exportRow[quarter] = row[quarter] || 0;
        });

        return exportRow;
      });

      // Add totals row
      const totalsRow = {
        "UAT Date": "TOTAL",
        "Product Name": "",
        Location: "",
        "Invoice Value": "",
        Quantity: "",
      };

      quarters.forEach((quarter) => {
        totalsRow[quarter] = columnTotals[quarter];
      });

      exportData.push(totalsRow);

      onExport(exportData, "Quarterly_AMC_Schedule");
    }
  }, [filteredData, quarters, columnTotals, onExport]);

  // Format currency values
  const formatCurrency = (value) => {
    if (!value || value === 0) return "";
    return `₹${Math.round(value).toLocaleString()}`;
  };

  // Define fixed column widths
  const fixedColumns = [
    { key: "uatDate", title: "UAT Date", width: 120 },
    { key: "productName", title: "Product Name", width: 200 },
    { key: "location", title: "Location", width: 130 },
    { key: "invoiceValue", title: "Invoice Value", width: 130 },
    { key: "quantity", title: "Quantity", width: 100 },
  ];

  return (
    <Card className="w-full">
      {/* Header */}
      <div className="mb-4">
        <div className="flex items-center justify-between mb-4">
          <div className="flex items-center gap-3">
            <Title level={4} className="mb-0">
              {title}
            </Title>
            <Tag color="blue">
              {filteredData.length} products × {quarters.length} quarters
            </Tag>
          </div>
          <Space>
            <Button
              icon={<ExportOutlined />}
              onClick={handleExport}
              type="primary"
              disabled={filteredData.length === 0}
            >
              Export to Excel
            </Button>
          </Space>
        </div>

        {/* Search and filters */}
        {(searchable || filterable) && (
          <div className="flex flex-wrap gap-4 mb-4">
            {searchable && (
              <Input
                placeholder="Search products, locations, dates..."
                prefix={<SearchOutlined />}
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                style={{ width: 300 }}
                allowClear
              />
            )}

            {filterable && (
              <Select
                placeholder="Filter by Location"
                style={{ width: 200 }}
                value={selectedLocationFilter}
                onChange={setSelectedLocationFilter}
              >
                {locations.map((location) => (
                  <Option key={location} value={location}>
                    {location}
                  </Option>
                ))}
              </Select>
            )}
          </div>
        )}

        {/* Summary */}
        <div className="flex flex-wrap gap-4 mb-4">
          <Tag color="blue">
            Products: {filteredData.length.toLocaleString()}
          </Tag>
          <Tag color="green">
            Total Value:{" "}
            {formatCurrency(
              filteredData.reduce(
                (sum, row) => sum + (row.totalAmountWithGST || 0),
                0
              )
            )}
          </Tag>
          <Tag color="purple">Location: {selectedLocationFilter}</Tag>
        </div>
      </div>

      {/* Quarterly Table */}
      <div
        style={{
          border: "1px solid #d1d5db",
          borderRadius: "8px",
          overflow: "auto",
          backgroundColor: "white",
          height: height,
        }}
      >
        <table
          style={{
            width: "100%",
            borderCollapse: "collapse",
            fontSize: "0.875rem",
            minWidth:
              fixedColumns.reduce((sum, col) => sum + col.width, 0) +
              quarters.length * 120,
          }}
        >
          {/* Table Header */}
          <thead style={{ position: "sticky", top: 0, zIndex: 10 }}>
            <tr style={{ backgroundColor: "#f3f4f6" }}>
              {/* Fixed columns */}
              {fixedColumns.map((column) => (
                <th
                  key={column.key}
                  style={{
                    padding: "12px 8px",
                    border: "1px solid #e5e7eb",
                    textAlign: "left",
                    fontWeight: 600,
                    position: "sticky",
                    left:
                      column.key === "uatDate"
                        ? 0
                        : column.key === "productName"
                        ? 120
                        : column.key === "location"
                        ? 320
                        : column.key === "invoiceValue"
                        ? 470
                        : column.key === "quantity"
                        ? 600
                        : "auto",
                    backgroundColor: "#f3f4f6",
                    zIndex: 9,
                    width: column.width,
                    minWidth: column.width,
                    maxWidth: column.width,
                  }}
                >
                  {column.title}
                </th>
              ))}

              {/* Quarter columns */}
              {quarters.map((quarter) => (
                <th
                  key={quarter}
                  style={{
                    padding: "12px 8px",
                    border: "1px solid #e5e7eb",
                    textAlign: "center",
                    fontWeight: 600,
                    backgroundColor: "#f3f4f6",
                    width: 120,
                    minWidth: 120,
                    maxWidth: 120,
                    transform: "rotate(-45deg)",
                    transformOrigin: "center",
                    height: "60px",
                    verticalAlign: "bottom",
                  }}
                >
                  <div
                    style={{
                      transform: "rotate(45deg)",
                      whiteSpace: "nowrap",
                      fontSize: "0.75rem",
                    }}
                  >
                    {quarter}
                  </div>
                </th>
              ))}
            </tr>
          </thead>

          {/* Table Body */}
          <tbody>
            {filteredData.map((row, index) => (
              <tr
                key={row.id || index}
                style={{
                  backgroundColor: index % 2 === 0 ? "white" : "#f9fafb",
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.backgroundColor = "#eff6ff";
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.backgroundColor =
                    index % 2 === 0 ? "white" : "#f9fafb";
                }}
              >
                {/* Fixed columns */}
                <td
                  style={{
                    padding: "8px",
                    border: "1px solid #e5e7eb",
                    position: "sticky",
                    left: 0,
                    backgroundColor: "inherit",
                    zIndex: 5,
                    width: 120,
                    fontSize: "0.8rem",
                  }}
                >
                  {row.uatDate}
                </td>
                <td
                  style={{
                    padding: "8px",
                    border: "1px solid #e5e7eb",
                    position: "sticky",
                    left: 120,
                    backgroundColor: "inherit",
                    zIndex: 5,
                    width: 200,
                    fontSize: "0.8rem",
                    overflow: "hidden",
                    textOverflow: "ellipsis",
                    whiteSpace: "nowrap",
                  }}
                  title={row.productName}
                >
                  {row.productName}
                </td>
                <td
                  style={{
                    padding: "8px",
                    border: "1px solid #e5e7eb",
                    position: "sticky",
                    left: 320,
                    backgroundColor: "inherit",
                    zIndex: 5,
                    width: 150,
                    fontSize: "0.8rem",
                  }}
                >
                  {row.location}
                </td>
                <td
                  style={{
                    padding: "8px",
                    border: "1px solid #e5e7eb",
                    position: "sticky",
                    left: 470,
                    backgroundColor: "inherit",
                    zIndex: 5,
                    width: 130,
                    textAlign: "right",
                    fontSize: "0.8rem",
                  }}
                >
                  {formatCurrency(row.invoiceValue)}
                </td>
                <td
                  style={{
                    padding: "8px",
                    border: "1px solid #e5e7eb",
                    position: "sticky",
                    left: 600,
                    backgroundColor: "inherit",
                    zIndex: 5,
                    width: 100,
                    textAlign: "center",
                    fontSize: "0.8rem",
                  }}
                >
                  {row.quantity}
                </td>

                {/* Quarter columns */}
                {quarters.map((quarter) => (
                  <td
                    key={quarter}
                    style={{
                      padding: "8px",
                      border: "1px solid #e5e7eb",
                      textAlign: "right",
                      width: 120,
                      fontSize: "0.75rem",
                      color: row[quarter] > 0 ? "#059669" : "#9ca3af",
                    }}
                  >
                    {row[quarter] > 0 ? formatCurrency(row[quarter]) : ""}
                  </td>
                ))}
              </tr>
            ))}

            {/* Totals Row */}
            <tr
              style={{
                backgroundColor: "#f0f9ff",
                fontWeight: 600,
                borderTop: "2px solid #3b82f6",
              }}
            >
              <td
                style={{
                  padding: "8px",
                  border: "1px solid #e5e7eb",
                  position: "sticky",
                  left: 0,
                  backgroundColor: "#f0f9ff",
                  zIndex: 5,
                  fontWeight: 600,
                }}
              >
                TOTAL
              </td>
              <td
                style={{
                  padding: "8px",
                  border: "1px solid #e5e7eb",
                  position: "sticky",
                  left: 120,
                  backgroundColor: "#f0f9ff",
                  zIndex: 5,
                }}
              ></td>
              <td
                style={{
                  padding: "8px",
                  border: "1px solid #e5e7eb",
                  position: "sticky",
                  left: 320,
                  backgroundColor: "#f0f9ff",
                  zIndex: 5,
                }}
              ></td>
              <td
                style={{
                  padding: "8px",
                  border: "1px solid #e5e7eb",
                  position: "sticky",
                  left: 470,
                  backgroundColor: "#f0f9ff",
                  zIndex: 5,
                }}
              ></td>
              <td
                style={{
                  padding: "8px",
                  border: "1px solid #e5e7eb",
                  position: "sticky",
                  left: 600,
                  backgroundColor: "#f0f9ff",
                  zIndex: 5,
                }}
              ></td>

              {quarters.map((quarter) => (
                <td
                  key={quarter}
                  style={{
                    padding: "8px",
                    border: "1px solid #e5e7eb",
                    textAlign: "right",
                    backgroundColor: "#f0f9ff",
                    fontWeight: 600,
                    color: "#2563eb",
                    fontSize: "0.75rem",
                  }}
                >
                  {formatCurrency(columnTotals[quarter])}
                </td>
              ))}
            </tr>
          </tbody>
        </table>
      </div>

      {/* Footer */}
      <div
        style={{
          marginTop: "16px",
          fontSize: "0.875rem",
          color: "#6b7280",
          textAlign: "center",
        }}
      >
        Showing {filteredData.length.toLocaleString()} products across{" "}
        {quarters.length} quarters
        {searchTerm && ` (filtered by "${searchTerm}")`}
      </div>
    </Card>
  );
};

export default QuarterlyDataTable;
