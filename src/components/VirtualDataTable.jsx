// Virtual Data Table for handling large datasets (100k+ rows)
import React, { useMemo, useCallback, useState, useEffect } from "react";
import { FixedSizeList as List } from "react-window";
import InfiniteLoader from "react-window-infinite-loader";
import { Card, Input, Select, Button, Space, Typography, Tag } from "antd";
import { SearchOutlined, ExportOutlined } from "@ant-design/icons";

const { Text } = Typography;
const { Option } = Select;

// Row component for AMC data with fixed grid layout
const AMCRow = React.memo(({ index, style, data }) => {
  const { items, onRowClick, visibleColumns, formatters } = data;
  const item = items[index];

  if (!item) {
    return (
      <div
        style={{
          ...style,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          borderBottom: "1px solid #e5e7eb",
          backgroundColor: "#f9fafb",
        }}
      >
        <Text type="secondary">Loading...</Text>
      </div>
    );
  }
  const handleRowClick = () => {
    if (onRowClick) {
      onRowClick(item, index);
    }
  };

  // Calculate total width to match header exactly
  const totalWidth = visibleColumns.reduce(
    (sum, col) => sum + (col.width || 150),
    0
  );

  return (
    <div
      style={{
        ...style,
        display: "flex",
        width: `${totalWidth}px`,
        minWidth: `${totalWidth}px`,
        borderBottom: "1px solid #e5e7eb",
        cursor: "pointer",
        backgroundColor: item?.id === 'TOTAL_ROW' ? "#f0f9ff" : "white",
        borderTop: item?.id === 'TOTAL_ROW' ? "2px solid #3b82f6" : "none",
      }}
      onClick={handleRowClick}
      onMouseEnter={(e) => {
        e.currentTarget.style.backgroundColor = 
          item?.id === 'TOTAL_ROW' ? "#dbeafe" : "#eff6ff";
      }}
      onMouseLeave={(e) => {
        e.currentTarget.style.backgroundColor = 
          item?.id === 'TOTAL_ROW' ? "#f0f9ff" : "white";
      }}
    >
      {visibleColumns.map((column) => {
        const value = item[column.key];
        const formattedValue = formatters[column.key]
          ? formatters[column.key](value, item)
          : value || "-";

        const isProductTotalColumn = column.key === 'rowTotal';
        const isTotalRow = item?.id === 'TOTAL_ROW';

        return (
          <div
            key={column.key}
            style={{
              width: `${column.width}px`,
              minWidth: `${column.width}px`,
              maxWidth: `${column.width}px`,
              padding: "8px 12px",
              display: "flex",
              alignItems: "center",
              textAlign: column.className?.includes("text-right")
                ? "right"
                : column.className?.includes("text-center")
                ? "center"
                : "left",
              justifyContent: column.className?.includes("text-right")
                ? "flex-end"
                : column.className?.includes("text-center")
                ? "center"
                : "flex-start",
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap",
              borderRight: "1px solid #f3f4f6",
              boxSizing: "border-box",
              backgroundColor: isProductTotalColumn ? 
                (isTotalRow ? "#bcfd9eff" : "#fff5e6ff") :
                (column.className?.includes("text-right") ? "inherit" : "inherit"),
              borderLeft: isProductTotalColumn ? "2px solid #f59e0b" : "none", // Orange border
              fontWeight: isProductTotalColumn ? "600" : "normal",
            }}
            title={String(formattedValue)} // Tooltip for long text
          >
            <span
              style={{
                overflow: "hidden",
                textOverflow: "ellipsis",
                whiteSpace: "nowrap",
                width: "100%",
              }}
            >
              {formattedValue}
            </span>
          </div>
        );
      })}
    </div>
  );
});

// Header component with fixed grid layout
const TableHeader = React.memo(({ columns, onSort, sortConfig }) => {
  // Calculate total width to match data rows exactly
  const totalWidth = columns.reduce((sum, col) => sum + (col.width || 150), 0);

  return (
    <div
      style={{
        display: "flex",
        width: `${totalWidth}px`,
        minWidth: `${totalWidth}px`,
        backgroundColor: "#f3f4f6",
        borderBottom: "2px solid #d1d5db",
        fontWeight: 600,
      }}
    >
      {columns.map((column) => {
        const isProductTotalColumn = column.key === 'rowTotal';
        
        return (
          <div
            key={column.key}
            style={{
              width: `${column.width}px`,
              minWidth: `${column.width}px`,
              maxWidth: `${column.width}px`,
              padding: "12px",
              display: "flex",
              alignItems: "center",
              justifyContent: column.className?.includes("text-left")
                ? "flex-end"
                : column.className?.includes("text-center")
                ? "center"
                : "flex-start",
              cursor: "pointer",
              userSelect: "none",
              borderRight: "1px solid #e5e7eb",
              transition: "background-color 0.2s ease",
              boxSizing: "border-box",
              backgroundColor: isProductTotalColumn ? "#fed7aa" : "#f3f4f6", // Orange background
              borderLeft: isProductTotalColumn ? "2px solid #f59e0b" : "none", // Orange border
            }}
            onClick={() => onSort && onSort(column.key)}
            onMouseEnter={(e) => {
              e.currentTarget.style.backgroundColor = isProductTotalColumn ? "#fdba74" : "#e5e7eb";
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.backgroundColor = isProductTotalColumn ? "#fed7aa" : "#f3f4f6";
            }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                width: "100%",
                overflow: "hidden",
              }}
            >
              <span
                style={{
                  overflow: "hidden",
                  textOverflow: "ellipsis",
                  whiteSpace: "nowrap",
                }}
              >
                {column.title}
              </span>
              {sortConfig.key === column.key && (
                <span
                  style={{
                    marginLeft: "4px",
                    fontSize: "0.75rem",
                    flexShrink: 0,
                  }}
                >
                  {sortConfig.direction === "asc" ? "↑" : "↓"}
                </span>
              )}
            </div>
          </div>
        );
      })}
    </div>
  );
});

// Virtual Data Table Component
const VirtualDataTable = ({
  data = [],
  columns = [],
  height = 600,
  rowHeight = 50,
  loading = false,
  onRowClick,
  onExport,
  title = "Data Table",
  searchable = true,
  filterable = true,
  sortable = true,
  showSummary = true,
  pageSize = 50,
  formatters = {},
  summary = {},
  onLoadMore,
  hasNextPage = false,
  // NEW PROPS FOR TOTALS
  showTotals = false,
  showRowTotals = false,
}) => {
  const [searchTerm, setSearchTerm] = useState("");
  const [debouncedSearchTerm, setDebouncedSearchTerm] = useState("");
  const [filters, setFilters] = useState({});
  const [sortConfig, setSortConfig] = useState({ key: null, direction: "asc" });
  const [isSorting, setIsSorting] = useState(false);

  // Debounce search term for better performance
  useEffect(() => {
    const timer = setTimeout(() => {
      setDebouncedSearchTerm(searchTerm);
    }, 300);

    return () => clearTimeout(timer);
  }, [searchTerm]);

  // Optimized comparison function for different data types
  const compareValues = useCallback((a, b, key, direction) => {
    let aVal = a[key];
    let bVal = b[key];

    // Handle null/undefined values
    if (aVal == null && bVal == null) return 0;
    if (aVal == null) return direction === "asc" ? -1 : 1;
    if (bVal == null) return direction === "asc" ? 1 : -1;

    // Convert to appropriate types for comparison
    const aStr = String(aVal);
    const bStr = String(bVal);

    // Check if both are numbers
    const aNum = parseFloat(aStr.replace(/[₹,\s]/g, ""));
    const bNum = parseFloat(bStr.replace(/[₹,\s]/g, ""));

    if (!isNaN(aNum) && !isNaN(bNum)) {
      // Numeric comparison
      const result = aNum - bNum;
      return direction === "asc" ? result : -result;
    }

    // Check if both are dates
    const aDate = new Date(aVal);
    const bDate = new Date(bVal);

    if (!isNaN(aDate.getTime()) && !isNaN(bDate.getTime())) {
      // Date comparison
      const result = aDate.getTime() - bDate.getTime();
      return direction === "asc" ? result : -result;
    }

    // String comparison (case-insensitive)
    const result = aStr.toLowerCase().localeCompare(bStr.toLowerCase());
    return direction === "asc" ? result : -result;
  }, []);

  // Filtered data (without sorting for better performance)
  const filteredData = useMemo(() => {
    let filtered = data;

    // Apply search filter with debounced term
    if (debouncedSearchTerm) {
      const searchLower = debouncedSearchTerm.toLowerCase();
      filtered = filtered.filter((item) =>
        Object.values(item).some((value) =>
          String(value).toLowerCase().includes(searchLower)
        )
      );
    }

    // Apply column filters
    Object.entries(filters).forEach(([key, value]) => {
      if (value && value !== "all") {
        filtered = filtered.filter((item) => item[key] === value);
      }
    });

    return filtered;
  }, [data, debouncedSearchTerm, filters]);

  // Sorted data (separate memo for better performance)
  const processedData = useMemo(() => {
    if (!sortConfig.key) {
      return filteredData;
    }

    const startTime = performance.now();
    setIsSorting(true);

    // Use a more efficient sorting algorithm for large datasets
    const sorted = [...filteredData];

    // Optimized sort with performance logging
    sorted.sort((a, b) =>
      compareValues(a, b, sortConfig.key, sortConfig.direction)
    );

    const endTime = performance.now();
    const sortDuration = endTime - startTime;

    if (sortDuration > 100) {
      console.log(
        `⚡ Sorting ${sorted.length} items by ${
          sortConfig.key
        } took ${sortDuration.toFixed(2)}ms`
      );
    }

    // Reset sorting indicator after a brief delay to show the loading state
    setTimeout(() => setIsSorting(false), 100);

    return sorted;
  }, [filteredData, sortConfig, compareValues]);

  // TOTALS CALCULATION
  const finalData = useMemo(() => {
    if (!showTotals || !processedData || processedData.length === 0) {
      return showRowTotals ? processedData.map(row => ({
        ...row,
        rowTotal: Object.keys(row)
          .filter(key => key.match(/^[A-Z]{3}-\d{4}$/))
          .reduce((sum, quarter) => sum + (row[quarter] || 0), 0)
      })) : processedData;
    }

    // Calculate total row
    const totalRow = {
      id: 'TOTAL_ROW',
      productName: 'TOTAL',
      location: `${processedData.length} Products`,
      amcStartDate: '',
      uatDate: '',
    };

    // Sum all numeric columns
    processedData.forEach(row => {
      Object.keys(row).forEach(key => {
        const value = row[key];
        if (typeof value === 'number') {
          if (!totalRow[key]) totalRow[key] = 0;
          totalRow[key] += value;
        }
      });
    });

    // Add row totals to all rows if requested
    const dataWithRowTotals = showRowTotals ? 
      [...processedData, totalRow].map(row => ({
        ...row,
        rowTotal: Object.keys(row)
          .filter(key => key.match(/^[A-Z]{3}-\d{4}$/))
          .reduce((sum, quarter) => sum + (row[quarter] || 0), 0)
      })) :
      [...processedData, totalRow];

    return dataWithRowTotals;
  }, [processedData, showTotals, showRowTotals]);

  // COLUMNS WITH ROW TOTALS
  const finalColumns = useMemo(() => {
    if (!showRowTotals) return columns;
    return [...columns, {
      key: "rowTotal",
      title: " PRODUCT TOTAL",
      width: 150,
      className: "text-left",
    }];
  }, [columns, showRowTotals]);

  // Handle sorting
  const handleSort = useCallback(
    (key) => {
      if (!sortable) return;

      setSortConfig((prev) => ({
        key,
        direction:
          prev.key === key && prev.direction === "asc" ? "desc" : "asc",
      }));
    },
    [sortable]
  );

  // Filter options for dropdowns (optimized)
  const filterOptions = useMemo(() => {
    const options = {};

    columns.forEach((column) => {
      if (column.filterable) {
        // Use Set for O(1) lookups and avoid re-sorting every time
        const valueSet = new Set();
        data.forEach((item) => {
          const value = item[column.key];
          if (value != null) {
            valueSet.add(value);
          }
        });

        // Convert to array and sort only once
        options[column.key] = Array.from(valueSet).sort((a, b) => {
          return String(a).localeCompare(String(b));
        });
      }
    });

    return options;
  }, [data, columns]);

  // Check if item is loaded for infinite scrolling
  const isItemLoaded = useCallback(
    (index) => {
      return !!finalData[index];
    },
    [finalData]
  );

  // Load more items for infinite scrolling
  const loadMoreItems = useCallback(
    (startIndex, stopIndex) => {
      if (onLoadMore && hasNextPage) {
        return onLoadMore(startIndex, stopIndex);
      }
      return Promise.resolve();
    },
    [onLoadMore, hasNextPage]
  );

  // Data for the list component
  const listData = useMemo(
    () => ({
      items: finalData,
      onRowClick,
      visibleColumns: finalColumns,
      formatters,
    }),
    [finalData, onRowClick, finalColumns, formatters]
  );

  return (
    <Card className="w-full">
      {/* Header with title and controls */}
      <div className="mb-4">
        <div className="flex items-center justify-between mb-4">
          <div className="flex items-center gap-3">
            <Typography.Title level={4} className="mb-0">
              {title}
            </Typography.Title>
            {sortConfig.key && (
              <Tag
                color={isSorting ? "orange" : "blue"}
                style={{ fontSize: "11px" }}
              >
                {isSorting
                  ? "⏳ Sorting..."
                  : `Sorted by ${
                      columns.find((c) => c.key === sortConfig.key)?.title ||
                      sortConfig.key
                    }`}
                {!isSorting && (sortConfig.direction === "asc" ? " ↑" : " ↓")}
              </Tag>
            )}
            {searchTerm && (
              <Tag color="orange" style={{ fontSize: "11px" }}>
                Filtered: "{searchTerm}"
              </Tag>
            )}
          </div>
          <Space>
            {onExport && (
              <Button
                icon={<ExportOutlined />}
                onClick={() => onExport(finalData)}
                type="primary"
              >
                Export ({finalData.length} items)
              </Button>
            )}
          </Space>
        </div>

        {/* Search and filters */}
        {(searchable || filterable) && (
          <div className="flex flex-wrap gap-4 mb-4">
            {searchable && (
              <Input
                placeholder="Search all columns..."
                prefix={<SearchOutlined />}
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                style={{ width: 300 }}
                allowClear
                suffix={
                  searchTerm !== debouncedSearchTerm && (
                    <span style={{ fontSize: "10px", color: "#999" }}>
                      searching...
                    </span>
                  )
                }
              />
            )}

            {filterable &&
              Object.entries(filterOptions).map(([key, options]) => {
                const column = columns.find((col) => col.key === key);
                return (
                  <Select
                    key={key}
                    placeholder={`Filter by ${column?.title || key}`}
                    style={{ width: 200 }}
                    value={filters[key] || "all"}
                    onChange={(value) =>
                      setFilters((prev) => ({ ...prev, [key]: value }))
                    }
                    allowClear
                  >
                    <Option value="all">All {column?.title || key}</Option>
                    {options.map((option) => (
                      <Option key={option} value={option}>
                        {option}
                      </Option>
                    ))}
                  </Select>
                );
              })}
          </div>
        )}

        {/* Summary stats */}
        {showSummary && (
          <div className="flex flex-wrap gap-4 mb-4">
            <Tag color="blue">
              Total: {processedData.length.toLocaleString()}
              {data.length !== processedData.length &&
                ` (of ${data.length.toLocaleString()})`}
            </Tag>
            {summary.totalValue && (
              <Tag color="green">
                Total Value: ₹{summary.totalValue.toLocaleString()}
              </Tag>
            )}
            {summary.successful && (
              <Tag color="success">
                Successful: {summary.successful.toLocaleString()}
              </Tag>
            )}
            {summary.errors > 0 && (
              <Tag color="error">Errors: {summary.errors.toLocaleString()}</Tag>
            )}
            {finalData.length > 10000 && (
              <Tag color="purple" style={{ fontSize: "10px" }}>
                ⚡ Virtual scrolling active
              </Tag>
            )}
          </div>
        )}
      </div>

      {/* Virtual scrolling table */}
      <div
        style={{
          border: "1px solid #d1d5db",
          borderRadius: "8px",
          overflow: "hidden",
          backgroundColor: "white",
        }}
      >
        {/* Scrollable container for both header and data */}
        {finalData.length > 0 ? (
          <div
            style={{
              position: "relative",
              overflow: "auto",
              backgroundColor: "#fafafa",
              height: height - 100, // Account for header and controls
            }}
          >
            {/* Table header - now inside scrollable container */}
            <div
              style={{
                position: "sticky",
                top: 0,
                zIndex: 2,
                backgroundColor: "white",
              }}
            >
              <TableHeader
                columns={finalColumns}
                onSort={handleSort}
                sortConfig={sortConfig}
              />
            </div>

            {/* Virtual list */}
            <InfiniteLoader
              isItemLoaded={isItemLoaded}
              itemCount={
                hasNextPage ? finalData.length + 1 : finalData.length
              }
              loadMoreItems={loadMoreItems}
              threshold={15}
            >
              {({ onItemsRendered, ref }) => {
                // Calculate total width for virtual list to match header/rows
                const totalWidth = finalColumns.reduce(
                  (sum, col) => sum + (col.width || 150),
                  0
                );

                return (
                  <List
                    ref={ref}
                    height={height - 150} // Account for header and controls + header height
                    itemCount={finalData.length}
                    itemSize={rowHeight}
                    itemData={listData}
                    onItemsRendered={onItemsRendered}
                    overscanCount={5}
                    width={totalWidth}
                  >
                    {AMCRow}
                  </List>
                );
              }}
            </InfiniteLoader>
          </div>
        ) : (
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              padding: "32px",
              backgroundColor: "#f9fafb",
            }}
          >
            <Text type="secondary">
              {loading ? "Loading data..." : "No data available"}
            </Text>
          </div>
        )}
      </div>

      {/* Footer with pagination info */}
      <div
        style={{
          marginTop: "16px",
          fontSize: "0.875rem",
          color: "#6b7280",
          textAlign: "center",
        }}
      >
        Showing {finalData.length.toLocaleString()} of{" "}
        {data.length.toLocaleString()} items
        {searchTerm && ` (filtered by "${searchTerm}")`}
      </div>
    </Card>
  );
};

export default VirtualDataTable;