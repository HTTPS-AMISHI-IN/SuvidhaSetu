// React hook for managing calculation cache with IndexedDB
import { useState, useEffect, useCallback } from "react";
import {
  amcCache,
  fileCache,
  cacheManager,
  cacheUtils,
} from "../utils/indexedDBCache";

export const useCalculationCache = () => {
  const [cacheStats, setCacheStats] = useState({
    counts: { amcCalculations: 0, warrantyCalculations: 0, excelFiles: 0 },
    estimatedSize: 0,
    estimatedSizeMB: 0,
  });
  const [isSupported, setIsSupported] = useState(false);
  const [quotaInfo, setQuotaInfo] = useState(null);
  const [recentCalculations, setRecentCalculations] = useState([]);
  const [recentFiles, setRecentFiles] = useState([]);

  // Update cache statistics
  const updateStats = useCallback(async () => {
    try {
      const stats = await cacheManager.getStats();
      setCacheStats(stats);
    } catch (error) {
      console.error("Error updating cache stats:", error);
    }
  }, []);

  // Update quota information
  const updateQuotaInfo = useCallback(async () => {
    try {
      const quota = await cacheUtils.getQuotaInfo();
      setQuotaInfo(quota);
    } catch (error) {
      console.error("Error getting quota info:", error);
    }
  }, []);

  // Load recent calculations and files
  const loadRecentData = useCallback(async () => {
    try {
      const [calculations, files] = await Promise.all([
        amcCache.getAll(),
        fileCache.getAll(),
      ]);

      setRecentCalculations(calculations.slice(0, 10)); // Last 10 calculations
      setRecentFiles(files.slice(0, 10)); // Last 10 files
    } catch (error) {
      console.error("Error loading recent data:", error);
    }
  }, []);

  // Initialize cache support check
  useEffect(() => {
    setIsSupported(cacheUtils.isSupported());
    updateStats();
    loadRecentData();
    updateQuotaInfo();
  }, [updateStats, loadRecentData, updateQuotaInfo]);

  // Check if calculation is cached
  const checkCalculationCache = useCallback(
    async (fileHash, settings) => {
      if (!isSupported) return null;

      try {
        const cached = await amcCache.findByFileAndSettings(fileHash, settings);
        return cached;
      } catch (error) {
        console.error("Error checking calculation cache:", error);
        return null;
      }
    },
    [isSupported]
  );

  // Store calculation results
  const storeCalculation = useCallback(
    async (calculationId, results, metadata) => {
      if (!isSupported) return null;

      try {
        const stored = await amcCache.store(calculationId, results, metadata);
        await updateStats();
        await loadRecentData();
        return stored;
      } catch (error) {
        console.error("Error storing calculation:", error);
        return null;
      }
    },
    [isSupported, updateStats, loadRecentData]
  );

  // Get cached calculation
  const getCalculation = useCallback(
    async (calculationId) => {
      if (!isSupported) return null;

      try {
        const result = await amcCache.get(calculationId);
        if (result) {
          await loadRecentData(); // Update access tracking
        }
        return result;
      } catch (error) {
        console.error("Error getting calculation:", error);
        return null;
      }
    },
    [isSupported, loadRecentData]
  );

  // Store Excel file data
  const storeFile = useCallback(
    async (file, processedData) => {
      if (!isSupported) return null;

      try {
        const result = await fileCache.store(file, processedData);
        await updateStats();
        await loadRecentData();
        return result;
      } catch (error) {
        console.error("Error storing file:", error);
        return null;
      }
    },
    [isSupported, updateStats, loadRecentData]
  );

  // Get cached file
  const getFile = useCallback(
    async (fileHash) => {
      if (!isSupported) return null;

      try {
        const result = await fileCache.get(fileHash);
        if (result) {
          await loadRecentData(); // Update access tracking
        }
        return result;
      } catch (error) {
        console.error("Error getting file:", error);
        return null;
      }
    },
    [isSupported, loadRecentData]
  );

  // Delete calculation
  const deleteCalculation = useCallback(
    async (calculationId) => {
      if (!isSupported) return false;

      try {
        await amcCache.delete(calculationId);
        await updateStats();
        await loadRecentData();
        return true;
      } catch (error) {
        console.error("Error deleting calculation:", error);
        return false;
      }
    },
    [isSupported, updateStats, loadRecentData]
  );

  // Delete file
  const deleteFile = useCallback(
    async (fileHash) => {
      if (!isSupported) return false;

      try {
        await fileCache.delete(fileHash);
        await updateStats();
        await loadRecentData();
        return true;
      } catch (error) {
        console.error("Error deleting file:", error);
        return false;
      }
    },
    [isSupported, updateStats, loadRecentData]
  );

  // Cleanup old cache entries
  const cleanupCache = useCallback(
    async (olderThanDays = 30) => {
      if (!isSupported) return 0;

      try {
        const deletedCount = await cacheManager.cleanup(olderThanDays);
        await updateStats();
        await loadRecentData();
        return deletedCount;
      } catch (error) {
        console.error("Error during cache cleanup:", error);
        return 0;
      }
    },
    [isSupported, updateStats, loadRecentData]
  );

  // Clear all cache
  const clearAllCache = useCallback(async () => {
    if (!isSupported) return false;

    try {
      await cacheManager.clearAll();
      await updateStats();
      await loadRecentData();
      return true;
    } catch (error) {
      console.error("Error clearing all cache:", error);
      return false;
    }
  }, [isSupported, updateStats, loadRecentData]);

  // Check if we can use cache (online/offline logic)
  const canUseCache = useCallback(() => {
    return (
      isSupported && (cacheUtils.isOnline() || recentCalculations.length > 0)
    );
  }, [isSupported, recentCalculations.length]);

  // Get cache strategy recommendation
  const getCacheStrategy = useCallback(
    (fileSize, dataLength) => {
      if (!isSupported) return "none";

      // For large datasets, always cache
      if (dataLength > 10000) return "always";

      // For medium datasets, cache if offline or space available
      if (dataLength > 1000) {
        if (!cacheUtils.isOnline()) return "always";
        if (quotaInfo && quotaInfo.usagePercentage < 80) return "conditional";
      }

      // For small datasets, cache only if offline
      if (!cacheUtils.isOnline()) return "offline_only";

      return "optional";
    },
    [isSupported, quotaInfo]
  );

  return {
    // Cache status
    isSupported,
    cacheStats,
    quotaInfo,
    canUseCache: canUseCache(),
    isOnline: cacheUtils.isOnline(),

    // Recent data
    recentCalculations,
    recentFiles,

    // Calculation cache operations
    checkCalculationCache,
    storeCalculation,
    getCalculation,
    deleteCalculation,

    // File cache operations
    storeFile,
    getFile,
    deleteFile,

    // Cache management
    updateStats,
    updateQuotaInfo,
    loadRecentData,
    cleanupCache,
    clearAllCache,

    // Utility functions
    getCacheStrategy,
    generateFileHash: cacheUtils.generateFileHash,
  };
};

// Hook for AMC-specific caching logic
export const useAMCCache = (fileData, settings) => {
  const cache = useCalculationCache();
  const [cacheKey, setCacheKey] = useState(null);
  const [cachedResult, setCachedResult] = useState(null);
  const [cacheStatus, setCacheStatus] = useState("checking"); // checking, hit, miss, storing, error

  // Generate cache key when file data or settings change
  useEffect(() => {
    const generateCacheKey = async () => {
      if (!fileData || !cache.isSupported) {
        setCacheKey(null);
        setCacheStatus("disabled");
        return;
      }

      try {
        setCacheStatus("checking");
        const fileHash = await cache.generateFileHash(fileData);
        const key = `${fileHash}_${JSON.stringify(settings)}`;
        setCacheKey(key);

        // Check if we have a cached result
        const cached = await cache.checkCalculationCache(fileHash, settings);
        if (cached) {
          setCachedResult(cached);
          setCacheStatus("hit");
        } else {
          setCachedResult(null);
          setCacheStatus("miss");
        }
      } catch (error) {
        console.error("Error generating cache key:", error);
        setCacheStatus("error");
      }
    };

    generateCacheKey();
  }, [fileData, settings, cache]);

  // Store calculation result
  const storeCachedResult = useCallback(
    async (results, metadata = {}) => {
      if (!cacheKey || !cache.isSupported) return null;

      try {
        setCacheStatus("storing");
        const fileHash = await cache.generateFileHash(fileData);

        const stored = await cache.storeCalculation(cacheKey, results, {
          ...metadata,
          fileHash,
          settings,
          calculationType: "amc",
        });

        if (stored) {
          setCachedResult(stored);
          setCacheStatus("stored");
        }

        return stored;
      } catch (error) {
        console.error("Error storing cached result:", error);
        setCacheStatus("error");
        return null;
      }
    },
    [cacheKey, cache, fileData, settings]
  );

  return {
    ...cache,
    cacheKey,
    cachedResult,
    cacheStatus,
    hasCachedResult: !!cachedResult,
    storeCachedResult,
  };
};
