// React hook for managing Web Workers with Comlink-style communication
import { useRef, useEffect, useCallback, useState } from "react";

export const useWebWorker = (workerPath, options = {}) => {
  const workerRef = useRef(null);
  const [isReady, setIsReady] = useState(false);
  const [error, setError] = useState(null);
  const messageHandlersRef = useRef(new Map());

  // Initialize worker
  useEffect(() => {
    try {
      const worker = new Worker(workerPath, { type: "module" });
      workerRef.current = worker;

      // Handle messages from worker
      worker.onmessage = (event) => {
        const { type, ...data } = event.data;

        if (type === "WORKER_READY") {
          setIsReady(true);
          setError(null);
          return;
        }

        // Call registered handlers
        const handlers = messageHandlersRef.current.get(type) || [];
        handlers.forEach((handler) => {
          try {
            handler(data);
          } catch (err) {
            console.error(`Error in message handler for ${type}:`, err);
          }
        });
      };

      // Handle worker errors
      worker.onerror = (event) => {
        console.error("Worker error:", event);
        setError(event.message || "Worker error");
        setIsReady(false);
      };

      worker.onmessageerror = (event) => {
        console.error("Worker message error:", event);
        setError("Worker communication error");
      };
    } catch (err) {
      console.error("Failed to create worker:", err);
      setError("Failed to initialize worker");
    }

    // Cleanup
    return () => {
      if (workerRef.current) {
        workerRef.current.terminate();
        workerRef.current = null;
      }
      setIsReady(false);
    };
  }, [workerPath]);

  // Register message handler
  const on = useCallback((messageType, handler) => {
    const handlers = messageHandlersRef.current.get(messageType) || [];
    handlers.push(handler);
    messageHandlersRef.current.set(messageType, handlers);

    // Return cleanup function
    return () => {
      const currentHandlers = messageHandlersRef.current.get(messageType) || [];
      const filteredHandlers = currentHandlers.filter((h) => h !== handler);
      if (filteredHandlers.length === 0) {
        messageHandlersRef.current.delete(messageType);
      } else {
        messageHandlersRef.current.set(messageType, filteredHandlers);
      }
    };
  }, []);

  // Send message to worker
  const postMessage = useCallback((type, data = {}) => {
    if (!workerRef.current) {
      console.warn("Worker not ready, cannot send message");
      return false;
    }

    try {
      workerRef.current.postMessage({ type, data });
      return true;
    } catch (err) {
      console.error("Failed to send message to worker:", err);
      setError("Failed to communicate with worker");
      return false;
    }
  }, []);

  // Terminate worker manually
  const terminate = useCallback(() => {
    if (workerRef.current) {
      workerRef.current.terminate();
      workerRef.current = null;
      setIsReady(false);
      messageHandlersRef.current.clear();
    }
  }, []);

  return {
    isReady,
    error,
    postMessage,
    on,
    terminate,
    worker: workerRef.current,
  };
};

// Hook specifically for AMC calculations
export const useAMCCalculationWorker = () => {
  const { isReady, error, postMessage, on, terminate } = useWebWorker(
    "/workers/amcCalculationWorker.js"
  );

  const [progress, setProgress] = useState({
    currentChunk: 0,
    totalChunks: 0,
    processed: 0,
    total: 0,
    isProcessing: false,
    timeElapsed: 0,
    estimatedTimeRemaining: 0,
  });

  const [results, setResults] = useState([]);
  const [summary, setSummary] = useState({
    totalProducts: 0,
    totalValue: 0,
    successful: 0,
    errors: 0,
    chunks: [],
  });

  // Handle progress updates
  useEffect(() => {
    const unsubscribeProgress = on("PROGRESS", (data) => {
      setProgress((prev) => {
        // Calculate absolute progress: (chunks completed * chunk size) + current chunk progress
        const chunksCompleted = data.chunkIndex || 0;
        const currentChunkProgress = data.processed || 0;
        const chunkSize = Math.floor(prev.total / prev.totalChunks) || 1000;
        const absoluteProcessed =
          chunksCompleted * chunkSize + currentChunkProgress;

        console.log("📊 Progress update:", {
          chunkIndex: data.chunkIndex,
          chunkProgress: data.processed,
          totalChunks: prev.totalChunks,
          absoluteProcessed,
          total: prev.total,
        });

        return {
          ...prev,
          processed: Math.min(absoluteProcessed, prev.total), // Don't exceed total
          timeElapsed: data.timeElapsed,
          estimatedTimeRemaining: calculateETA(
            data.timeElapsed,
            absoluteProcessed,
            prev.total
          ),
        };
      });
    });

    const unsubscribeChunkStart = on("CHUNK_STARTED", (data) => {
      setProgress((prev) => ({
        ...prev,
        currentChunk: data.chunkIndex + 1,
        isProcessing: true,
      }));
    });

    const unsubscribeChunkComplete = on("CHUNK_COMPLETE", (data) => {
      setResults((prev) => [...prev, ...data.results]);

      setSummary((prev) => ({
        ...prev,
        totalProducts: prev.totalProducts + data.summary.processed,
        totalValue: prev.totalValue + data.summary.totalValue,
        successful: prev.successful + data.summary.successful,
        errors: prev.errors + data.summary.errors,
        chunks: [...prev.chunks, data.summary],
      }));

      setProgress((prev) => {
        const completedChunks = data.chunkIndex + 1;
        const chunkSize = Math.floor(prev.total / prev.totalChunks) || 1000;
        const processedCount = Math.min(
          completedChunks * chunkSize,
          prev.total
        );

        console.log("✅ Chunk completed:", {
          chunkIndex: data.chunkIndex,
          completedChunks,
          totalChunks: data.totalChunks,
          processedCount,
          isFinished: completedChunks >= data.totalChunks,
        });

        const isFinished = completedChunks >= data.totalChunks;

        return {
          ...prev,
          currentChunk: completedChunks,
          processed: isFinished ? prev.total : processedCount, // Ensure 100% when finished
          isProcessing: !isFinished,
          estimatedTimeRemaining: isFinished ? 0 : prev.estimatedTimeRemaining,
        };
      });
    });

    const unsubscribeError = on("ERROR", (data) => {
      console.error("❌ Worker error:", data.error);
      setProgress((prev) => ({
        ...prev,
        isProcessing: false,
        estimatedTimeRemaining: 0,
      }));
    });

    return () => {
      unsubscribeProgress();
      unsubscribeChunkStart();
      unsubscribeChunkComplete();
      unsubscribeError();
    };
  }, [on]);

  // Calculate chunks and start processing
  const calculateAMCForDataset = useCallback(
    (products, settings = {}, chunkSize = 1000) => {
      if (!isReady || !products || products.length === 0) {
        console.warn(
          "Cannot start calculation: worker not ready or no products"
        );
        return;
      }

      // Reset state
      setResults([]);
      setSummary({
        totalProducts: 0,
        totalValue: 0,
        successful: 0,
        errors: 0,
        chunks: [],
      });

      // Calculate chunks
      const chunks = [];
      for (let i = 0; i < products.length; i += chunkSize) {
        chunks.push(products.slice(i, i + chunkSize));
      }

      setProgress({
        currentChunk: 0,
        totalChunks: chunks.length,
        processed: 0,
        total: products.length,
        isProcessing: true,
        timeElapsed: 0,
        estimatedTimeRemaining: 0,
      });

      // Send chunks to worker
      chunks.forEach((chunk, index) => {
        postMessage("CALCULATE_CHUNK", {
          products: chunk,
          settings,
          chunkIndex: index,
          totalChunks: chunks.length,
        });
      });
    },
    [isReady, postMessage]
  );

  // Calculate single product
  const calculateSingleProduct = useCallback(
    (product, settings = {}) => {
      if (!isReady) {
        console.warn("Worker not ready");
        return;
      }

      return new Promise((resolve, reject) => {
        const cleanup = on("SINGLE_COMPLETE", (data) => {
          cleanup();
          resolve(data.result);
        });

        const errorCleanup = on("ERROR", (data) => {
          errorCleanup();
          reject(new Error(data.error));
        });

        postMessage("CALCULATE_SINGLE", { product, settings });
      });
    },
    [isReady, postMessage, on]
  );

  return {
    isReady,
    error,
    progress,
    results,
    summary,
    calculateAMCForDataset,
    calculateSingleProduct,
    terminate,
  };
};

// Utility function to calculate estimated time remaining
function calculateETA(elapsed, processed, total) {
  if (processed === 0) return 0;
  const rate = processed / elapsed;
  const remaining = total - processed;
  return remaining / rate;
}
