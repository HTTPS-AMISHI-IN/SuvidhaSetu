// Cache Management UI Component
import React, { useState } from "react";
import {
  Card,
  Button,
  Space,
  Typography,
  Progress,
  List,
  Tag,
  Modal,
  Tooltip,
  Alert,
  Row,
  Col,
  Statistic,
  Popconfirm,
  Input,
  Select,
} from "antd";
import {
  DeleteOutlined,
  ClearOutlined,
  ReloadOutlined,
  InfoCircleOutlined,
  FileExcelOutlined,
  CalculatorOutlined,
  CloudOutlined,
  DisconnectOutlined,
  SettingOutlined,
} from "@ant-design/icons";
import { useCalculationCache } from "../hooks/useCalculationCache";

const { Title, Text, Paragraph } = Typography;
const { Option } = Select;

const CacheManager = ({ visible, onClose, compact = false }) => {
  const {
    isSupported,
    cacheStats,
    quotaInfo,
    canUseCache,
    isOnline,
    recentCalculations,
    recentFiles,
    updateStats,
    deleteCalculation,
    deleteFile,
    cleanupCache,
    clearAllCache,
  } = useCalculationCache();

  const [loading, setLoading] = useState(false);
  const [cleanupDays, setCleanupDays] = useState(30);

  // Handle cleanup
  const handleCleanup = async () => {
    setLoading(true);
    try {
      const deletedCount = await cleanupCache(cleanupDays);
      Modal.success({
        title: "Cache Cleanup Complete",
        content: `Successfully deleted ${deletedCount} old cache entries.`,
      });
    } catch (error) {
      Modal.error({
        title: "Cleanup Failed",
        content: "An error occurred during cache cleanup.",
      });
    } finally {
      setLoading(false);
    }
  };

  // Handle clear all
  const handleClearAll = async () => {
    setLoading(true);
    try {
      await clearAllCache();
      Modal.success({
        title: "Cache Cleared",
        content: "All cached data has been successfully removed.",
      });
    } catch (error) {
      Modal.error({
        title: "Clear Failed",
        content: "An error occurred while clearing the cache.",
      });
    } finally {
      setLoading(false);
    }
  };

  // Handle delete calculation
  const handleDeleteCalculation = async (calculationId) => {
    try {
      await deleteCalculation(calculationId);
      await updateStats();
    } catch (error) {
      Modal.error({
        title: "Delete Failed",
        content: "Failed to delete calculation from cache.",
      });
    }
  };

  // Handle delete file
  const handleDeleteFile = async (fileHash) => {
    try {
      await deleteFile(fileHash);
      await updateStats();
    } catch (error) {
      Modal.error({
        title: "Delete Failed",
        content: "Failed to delete file from cache.",
      });
    }
  };

  // Format file size
  const formatFileSize = (bytes) => {
    if (bytes === 0) return "0 B";
    const k = 1024;
    const sizes = ["B", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
  };

  // Format date
  const formatDate = (timestamp) => {
    return new Date(timestamp).toLocaleDateString("en-US", {
      year: "numeric",
      month: "short",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });
  };

  if (!isSupported) {
    return (
      <Modal
        title="Cache Manager"
        open={visible}
        onCancel={onClose}
        footer={null}
        width={600}
      >
        <Alert
          type="warning"
          message="IndexedDB Not Supported"
          description="Your browser doesn't support IndexedDB. Caching features are not available."
          showIcon
        />
      </Modal>
    );
  }

  if (compact) {
    return (
      <Card size="small" className="mb-4">
        <div className="flex items-center justify-between">
          <div className="flex items-center space-x-4">
            <div className="flex items-center">
              {isOnline ? (
                <CloudOutlined className="text-green-500 mr-1" />
              ) : (
                <DisconnectOutlined className="text-orange-500 mr-1" />
              )}
              <Text type="secondary" className="text-sm">
                {isOnline ? "Online" : "Offline"}
              </Text>
            </div>

            <div className="flex items-center">
              <CalculatorOutlined className="text-blue-500 mr-1" />
              <Text type="secondary" className="text-sm">
                {cacheStats.counts.amcCalculations} cached
              </Text>
            </div>

            <div className="flex items-center">
              <FileExcelOutlined className="text-green-500 mr-1" />
              <Text type="secondary" className="text-sm">
                {cacheStats.counts.excelFiles} files
              </Text>
            </div>
          </div>

          <Button
            size="small"
            icon={<SettingOutlined />}
            onClick={() =>
              Modal.info({
                title: "Cache Status",
                content: (
                  <div className="space-y-2">
                    <p>
                      Cached Calculations: {cacheStats.counts.amcCalculations}
                    </p>
                    <p>Cached Files: {cacheStats.counts.excelFiles}</p>
                    <p>Storage Used: {cacheStats.estimatedSizeMB} MB</p>
                    {quotaInfo && (
                      <p>Quota Usage: {quotaInfo.usagePercentage}%</p>
                    )}
                  </div>
                ),
              })
            }
          >
            Cache Info
          </Button>
        </div>
      </Card>
    );
  }

  return (
    <Modal
      title={
        <div className="flex items-center">
          <SettingOutlined className="mr-2" />
          Cache Manager
        </div>
      }
      open={visible}
      onCancel={onClose}
      footer={null}
      width={800}
      className="cache-manager-modal"
    >
      <div className="space-y-6">
        {/* Status Overview */}
        <Card size="small">
          <Row gutter={16}>
            <Col span={6}>
              <Statistic
                title="Cached Calculations"
                value={cacheStats.counts.amcCalculations}
                prefix={<CalculatorOutlined />}
              />
            </Col>
            <Col span={6}>
              <Statistic
                title="Cached Files"
                value={cacheStats.counts.excelFiles}
                prefix={<FileExcelOutlined />}
              />
            </Col>
            <Col span={6}>
              <Statistic
                title="Storage Used"
                value={cacheStats.estimatedSizeMB}
                suffix="MB"
                precision={2}
              />
            </Col>
            <Col span={6}>
              <div className="flex items-center space-x-2">
                {isOnline ? (
                  <Tag color="green" icon={<CloudOutlined />}>
                    Online
                  </Tag>
                ) : (
                  <Tag color="orange" icon={<DisconnectOutlined />}>
                    Offline
                  </Tag>
                )}
                {canUseCache && <Tag color="blue">Cache Available</Tag>}
              </div>
            </Col>
          </Row>

          {quotaInfo && (
            <div className="mt-4">
              <Text strong>Storage Quota</Text>
              <Progress
                percent={quotaInfo.usagePercentage}
                status={quotaInfo.usagePercentage > 80 ? "exception" : "normal"}
                format={(percent) =>
                  `${percent}% (${formatFileSize(
                    quotaInfo.usage
                  )} / ${formatFileSize(quotaInfo.quota)})`
                }
              />
            </div>
          )}
        </Card>

        {/* Cache Management Actions */}
        <Card title="Cache Management" size="small">
          <Space wrap>
            <Button
              icon={<ReloadOutlined />}
              onClick={updateStats}
              loading={loading}
            >
              Refresh Stats
            </Button>

            <div className="flex items-center space-x-2">
              <Select
                value={cleanupDays}
                onChange={setCleanupDays}
                style={{ width: 120 }}
                size="small"
              >
                <Option value={7}>7 days</Option>
                <Option value={30}>30 days</Option>
                <Option value={60}>60 days</Option>
                <Option value={90}>90 days</Option>
              </Select>
              <Button
                icon={<ClearOutlined />}
                onClick={handleCleanup}
                loading={loading}
              >
                Cleanup Old
              </Button>
            </div>

            <Popconfirm
              title="Clear All Cache"
              description="This will delete all cached calculations and files. Are you sure?"
              onConfirm={handleClearAll}
              okText="Yes"
              cancelText="No"
            >
              <Button danger icon={<DeleteOutlined />} loading={loading}>
                Clear All
              </Button>
            </Popconfirm>
          </Space>
        </Card>

        {/* Recent Calculations */}
        {recentCalculations.length > 0 && (
          <Card title="Recent Calculations" size="small">
            <List
              size="small"
              dataSource={recentCalculations.slice(0, 5)}
              renderItem={(item) => (
                <List.Item
                  actions={[
                    <Tooltip title="View Details">
                      <Button
                        size="small"
                        icon={<InfoCircleOutlined />}
                        onClick={() =>
                          Modal.info({
                            title: "Calculation Details",
                            content: (
                              <div className="space-y-2">
                                <p>
                                  <strong>Results:</strong>{" "}
                                  {item.metadata.resultCount} products
                                </p>
                                <p>
                                  <strong>Total Value:</strong> ₹
                                  {item.metadata.totalValue?.toLocaleString()}
                                </p>
                                <p>
                                  <strong>Created:</strong>{" "}
                                  {formatDate(item.metadata.timestamp)}
                                </p>
                                <p>
                                  <strong>Accessed:</strong> {item.accessCount}{" "}
                                  times
                                </p>
                                <p>
                                  <strong>Settings:</strong>{" "}
                                  {JSON.stringify(item.settings, null, 2)}
                                </p>
                              </div>
                            ),
                            width: 600,
                          })
                        }
                      />
                    </Tooltip>,
                    <Popconfirm
                      title="Delete this calculation?"
                      onConfirm={() =>
                        handleDeleteCalculation(item.calculationId)
                      }
                    >
                      <Button size="small" danger icon={<DeleteOutlined />} />
                    </Popconfirm>,
                  ]}
                >
                  <List.Item.Meta
                    avatar={<CalculatorOutlined />}
                    title={
                      <div className="flex items-center space-x-2">
                        <Text strong>AMC Calculation</Text>
                        <Tag color="blue">
                          {item.metadata.resultCount} products
                        </Tag>
                      </div>
                    }
                    description={
                      <div className="text-sm text-gray-500">
                        <div>
                          Value: ₹{item.metadata.totalValue?.toLocaleString()}
                        </div>
                        <div>
                          Created: {formatDate(item.metadata.timestamp)}
                        </div>
                      </div>
                    }
                  />
                </List.Item>
              )}
            />
          </Card>
        )}

        {/* Recent Files */}
        {recentFiles.length > 0 && (
          <Card title="Recent Files" size="small">
            <List
              size="small"
              dataSource={recentFiles.slice(0, 5)}
              renderItem={(item) => (
                <List.Item
                  actions={[
                    <Popconfirm
                      title="Delete this file?"
                      onConfirm={() => handleDeleteFile(item.fileHash)}
                    >
                      <Button size="small" danger icon={<DeleteOutlined />} />
                    </Popconfirm>,
                  ]}
                >
                  <List.Item.Meta
                    avatar={<FileExcelOutlined />}
                    title={
                      <div className="flex items-center space-x-2">
                        <Text strong>{item.fileName}</Text>
                        <Tag color="green">{formatFileSize(item.fileSize)}</Tag>
                      </div>
                    }
                    description={
                      <div className="text-sm text-gray-500">
                        <div>
                          Uploaded:{" "}
                          {formatDate(new Date(item.uploadDate).getTime())}
                        </div>
                        <div>Accessed: {item.accessCount} times</div>
                      </div>
                    }
                  />
                </List.Item>
              )}
            />
          </Card>
        )}

        {/* Help Information */}
        <Alert
          type="info"
          message="Cache Benefits"
          description={
            <div className="text-sm">
              <p>• Faster calculation loading for previously processed files</p>
              <p>• Offline access to calculation results</p>
              <p>• Reduced processing time for large datasets</p>
              <p>• Automatic cleanup of old cache entries</p>
            </div>
          }
          showIcon
        />
      </div>
    </Modal>
  );
};

export default CacheManager;
