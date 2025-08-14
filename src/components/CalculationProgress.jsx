// Progress indicator for large dataset calculations
import React from "react";
import { Card, Progress, Typography, Space, Tag, Button, Divider } from "antd";
import {
  PauseCircleOutlined,
  PlayCircleOutlined,
  StopOutlined,
  InfoCircleOutlined,
} from "@ant-design/icons";

const { Title, Text } = Typography;

const CalculationProgress = ({
  progress,
  summary,
  isCalculating = false,
  onCancel,
  onPause,
  onResume,
  isPaused = false,
  title = "Processing Calculations",
}) => {
  const {
    currentChunk = 0,
    totalChunks = 0,
    processed = 0,
    total = 0,
    timeElapsed = 0,
    estimatedTimeRemaining = 0,
  } = progress;

  // Calculate percentage
  const percentComplete = total > 0 ? Math.round((processed / total) * 100) : 0;
  const chunkPercent =
    totalChunks > 0 ? Math.round((currentChunk / totalChunks) * 100) : 0;

  // Format time
  const formatTime = (milliseconds) => {
    const seconds = Math.floor(milliseconds / 1000);
    const minutes = Math.floor(seconds / 60);
    const hours = Math.floor(minutes / 60);

    if (hours > 0) {
      return `${hours}h ${minutes % 60}m ${seconds % 60}s`;
    } else if (minutes > 0) {
      return `${minutes}m ${seconds % 60}s`;
    } else {
      return `${seconds}s`;
    }
  };

  // Calculate processing rate
  const processingRate =
    timeElapsed > 0 ? Math.round((processed / (timeElapsed / 1000)) * 60) : 0;

  if (!isCalculating && processed === 0) {
    return null;
  }

  return (
    <Card className="mb-6 shadow-lg">
      <div className="space-y-4">
        {/* Header */}
        <div className="flex items-center justify-between">
          <Title level={4} className="mb-0 flex items-center">
            <InfoCircleOutlined className="mr-2 text-blue-500" />
            {title}
          </Title>

          <Space>
            {isCalculating && (
              <>
                {isPaused ? (
                  <Button
                    icon={<PlayCircleOutlined />}
                    onClick={onResume}
                    type="primary"
                    size="small"
                  >
                    Resume
                  </Button>
                ) : (
                  <Button
                    icon={<PauseCircleOutlined />}
                    onClick={onPause}
                    size="small"
                  >
                    Pause
                  </Button>
                )}
                <Button
                  icon={<StopOutlined />}
                  onClick={onCancel}
                  danger
                  size="small"
                >
                  Cancel
                </Button>
              </>
            )}
          </Space>
        </div>

        {/* Overall Progress */}
        <div>
          <div className="flex items-center justify-between mb-2">
            <Text strong>Overall Progress</Text>
            <Text type="secondary">
              {processed.toLocaleString()} / {total.toLocaleString()} products
            </Text>
          </div>
          <Progress
            percent={percentComplete}
            status={
              isPaused ? "exception" : isCalculating ? "active" : "success"
            }
            strokeColor={{
              "0%": "#108ee9",
              "100%": "#87d068",
            }}
            showInfo={true}
            format={(percent) => `${percent}%`}
          />
        </div>

        {/* Chunk Progress */}
        {totalChunks > 1 && (
          <div>
            <div className="flex items-center justify-between mb-2">
              <Text strong>Chunk Progress</Text>
              <Text type="secondary">
                Chunk {currentChunk} / {totalChunks}
              </Text>
            </div>
            <Progress
              percent={chunkPercent}
              size="small"
              status={
                isPaused ? "exception" : isCalculating ? "active" : "success"
              }
            />
          </div>
        )}

        {/* Statistics */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
          <div className="text-center">
            <div className="text-2xl font-bold text-blue-600">
              {formatTime(timeElapsed)}
            </div>
            <Text type="secondary" className="text-sm">
              Time Elapsed
            </Text>
          </div>

          <div className="text-center">
            <div className="text-2xl font-bold text-green-600">
              {estimatedTimeRemaining > 0
                ? formatTime(estimatedTimeRemaining)
                : "-"}
            </div>
            <Text type="secondary" className="text-sm">
              Est. Remaining
            </Text>
          </div>

          <div className="text-center">
            <div className="text-2xl font-bold text-purple-600">
              {processingRate.toLocaleString()}/min
            </div>
            <Text type="secondary" className="text-sm">
              Processing Rate
            </Text>
          </div>

          <div className="text-center">
            <div className="text-2xl font-bold text-orange-600">
              {totalChunks.toLocaleString()}
            </div>
            <Text type="secondary" className="text-sm">
              Total Chunks
            </Text>
          </div>
        </div>

        {/* Status Tags */}
        <div className="flex flex-wrap gap-2">
          <Tag color={isCalculating ? "processing" : "success"}>
            {isPaused ? "Paused" : isCalculating ? "Processing" : "Completed"}
          </Tag>

          {summary.successful > 0 && (
            <Tag color="success">
              Successful: {summary.successful.toLocaleString()}
            </Tag>
          )}

          {summary.errors > 0 && (
            <Tag color="error">Errors: {summary.errors.toLocaleString()}</Tag>
          )}

          {summary.totalValue > 0 && (
            <Tag color="blue">
              Total Value: ₹{Math.round(summary.totalValue).toLocaleString()}
            </Tag>
          )}
        </div>

        {/* Memory and Performance Info */}
        {isCalculating && (
          <>
            <Divider style={{ margin: "12px 0" }} />
            <div className="text-xs text-gray-500 space-y-1">
              <div>
                💡 Large datasets are processed in chunks to maintain browser
                performance
              </div>
              <div>
                🚀 Calculations run in Web Workers to prevent UI blocking
              </div>
              <div>
                📊 Only visible rows are rendered using virtual scrolling
              </div>
            </div>
          </>
        )}
      </div>
    </Card>
  );
};

export default CalculationProgress;
