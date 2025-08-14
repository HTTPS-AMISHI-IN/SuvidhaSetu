// IndexedDB Cache System for AMC and Warranty Calculations
import { openDB } from "idb";

const DB_NAME = "PPOCalculationCache";
const DB_VERSION = 1;

// Store names
const STORES = {
  AMC_CALCULATIONS: "amcCalculations",
  WARRANTY_CALCULATIONS: "warrantyCalculations",
  EXCEL_FILES: "excelFiles",
  CALCULATION_METADATA: "calculationMetadata",
};

// Database initialization
let dbPromise = null;

const initDB = async () => {
  if (dbPromise) return dbPromise;

  dbPromise = openDB(DB_NAME, DB_VERSION, {
    upgrade(db) {
      // AMC Calculations store
      if (!db.objectStoreNames.contains(STORES.AMC_CALCULATIONS)) {
        const amcStore = db.createObjectStore(STORES.AMC_CALCULATIONS, {
          keyPath: "id",
          autoIncrement: false,
        });
        amcStore.createIndex("calculationId", "calculationId", {
          unique: false,
        });
        amcStore.createIndex("timestamp", "timestamp", { unique: false });
        amcStore.createIndex("fileHash", "fileHash", { unique: false });
      }

      // Warranty Calculations store
      if (!db.objectStoreNames.contains(STORES.WARRANTY_CALCULATIONS)) {
        const warrantyStore = db.createObjectStore(
          STORES.WARRANTY_CALCULATIONS,
          {
            keyPath: "id",
            autoIncrement: false,
          }
        );
        warrantyStore.createIndex("calculationId", "calculationId", {
          unique: false,
        });
        warrantyStore.createIndex("timestamp", "timestamp", { unique: false });
        warrantyStore.createIndex("fileHash", "fileHash", { unique: false });
      }

      // Excel Files store (for caching uploaded files)
      if (!db.objectStoreNames.contains(STORES.EXCEL_FILES)) {
        const excelStore = db.createObjectStore(STORES.EXCEL_FILES, {
          keyPath: "id",
          autoIncrement: false,
        });
        excelStore.createIndex("fileHash", "fileHash", { unique: true });
        excelStore.createIndex("fileName", "fileName", { unique: false });
        excelStore.createIndex("uploadDate", "uploadDate", { unique: false });
      }

      // Calculation Metadata store (for tracking calculations)
      if (!db.objectStoreNames.contains(STORES.CALCULATION_METADATA)) {
        const metadataStore = db.createObjectStore(
          STORES.CALCULATION_METADATA,
          {
            keyPath: "id",
            autoIncrement: false,
          }
        );
        metadataStore.createIndex("type", "type", { unique: false });
        metadataStore.createIndex("status", "status", { unique: false });
        metadataStore.createIndex("createdAt", "createdAt", { unique: false });
      }
    },
  });

  return dbPromise;
};

// Utility functions
const generateFileHash = async (data) => {
  const dataString = typeof data === "string" ? data : JSON.stringify(data);
  const encoder = new TextEncoder();
  const dataBytes = encoder.encode(dataString);
  const hashBuffer = await crypto.subtle.digest("SHA-256", dataBytes);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map((b) => b.toString(16).padStart(2, "0")).join("");
};

// AMC Calculation Cache
export const amcCache = {
  // Store AMC calculation results
  async store(calculationId, results, metadata = {}) {
    try {
      const db = await initDB();
      const tx = db.transaction(STORES.AMC_CALCULATIONS, "readwrite");

      const cacheEntry = {
        id: calculationId,
        calculationId,
        results,
        metadata: {
          ...metadata,
          timestamp: Date.now(),
          resultCount: results.length,
          totalValue: results.reduce(
            (sum, r) => sum + (r.totalAmountWithGST || 0),
            0
          ),
        },
        fileHash: metadata.fileHash,
        settings: metadata.settings || {},
        createdAt: new Date().toISOString(),
        accessCount: 0,
        lastAccessed: Date.now(),
      };

      await tx.store.put(cacheEntry);
      await tx.done;

      console.log(`✅ AMC calculation cached: ${calculationId}`);
      return cacheEntry;
    } catch (error) {
      console.error("Error storing AMC calculation:", error);
      throw error;
    }
  },

  // Retrieve AMC calculation results
  async get(calculationId) {
    try {
      const db = await initDB();
      const tx = db.transaction(STORES.AMC_CALCULATIONS, "readwrite");
      const result = await tx.store.get(calculationId);

      if (result) {
        // Update access tracking
        result.accessCount = (result.accessCount || 0) + 1;
        result.lastAccessed = Date.now();
        await tx.store.put(result);
        await tx.done;

        console.log(
          `📄 AMC calculation retrieved from cache: ${calculationId}`
        );
        return result;
      }

      await tx.done;
      return null;
    } catch (error) {
      console.error("Error retrieving AMC calculation:", error);
      return null;
    }
  },

  // Get all cached AMC calculations
  async getAll() {
    try {
      const db = await initDB();
      const results = await db.getAll(STORES.AMC_CALCULATIONS);
      return results.sort((a, b) => b.lastAccessed - a.lastAccessed);
    } catch (error) {
      console.error("Error getting all AMC calculations:", error);
      return [];
    }
  },

  // Check if calculation exists for file hash and settings
  async findByFileAndSettings(fileHash, settings) {
    try {
      const db = await initDB();
      const tx = db.transaction(STORES.AMC_CALCULATIONS, "readonly");
      const index = tx.store.index("fileHash");
      const results = await index.getAll(fileHash);

      // Find matching settings
      const match = results.find((result) => {
        const cachedSettings = result.settings || {};
        const currentSettings = settings || {};

        return (
          JSON.stringify(cachedSettings) === JSON.stringify(currentSettings)
        );
      });

      await tx.done;
      return match || null;
    } catch (error) {
      console.error("Error finding calculation by file and settings:", error);
      return null;
    }
  },

  // Delete calculation
  async delete(calculationId) {
    try {
      const db = await initDB();
      await db.delete(STORES.AMC_CALCULATIONS, calculationId);
      console.log(`🗑️ AMC calculation deleted: ${calculationId}`);
    } catch (error) {
      console.error("Error deleting AMC calculation:", error);
      throw error;
    }
  },
};

// Excel File Cache
export const fileCache = {
  // Store uploaded Excel file
  async store(file, processedData) {
    try {
      const fileHash = await generateFileHash(processedData);
      const db = await initDB();

      const cacheEntry = {
        id: fileHash,
        fileHash,
        fileName: file.name,
        fileSize: file.size,
        fileType: file.type,
        processedData,
        uploadDate: new Date().toISOString(),
        lastAccessed: Date.now(),
        accessCount: 1,
      };

      await db.put(STORES.EXCEL_FILES, cacheEntry);
      console.log(`📁 Excel file cached: ${file.name}`);
      return { fileHash, cacheEntry };
    } catch (error) {
      console.error("Error storing Excel file:", error);
      throw error;
    }
  },

  // Get cached Excel file
  async get(fileHash) {
    try {
      const db = await initDB();
      const result = await db.get(STORES.EXCEL_FILES, fileHash);

      if (result) {
        // Update access tracking
        result.accessCount = (result.accessCount || 0) + 1;
        result.lastAccessed = Date.now();
        await db.put(STORES.EXCEL_FILES, result);

        console.log(`📄 Excel file retrieved from cache: ${result.fileName}`);
      }

      return result;
    } catch (error) {
      console.error("Error retrieving Excel file:", error);
      return null;
    }
  },

  // Get all cached files
  async getAll() {
    try {
      const db = await initDB();
      const results = await db.getAll(STORES.EXCEL_FILES);
      return results.sort(
        (a, b) => new Date(b.uploadDate) - new Date(a.uploadDate)
      );
    } catch (error) {
      console.error("Error getting all Excel files:", error);
      return [];
    }
  },

  // Delete cached file
  async delete(fileHash) {
    try {
      const db = await initDB();
      await db.delete(STORES.EXCEL_FILES, fileHash);
      console.log(`🗑️ Excel file deleted from cache: ${fileHash}`);
    } catch (error) {
      console.error("Error deleting Excel file:", error);
      throw error;
    }
  },
};

// Cache Management
export const cacheManager = {
  // Get cache statistics
  async getStats() {
    try {
      const db = await initDB();

      const [amcCalculations, warrantyCalculations, excelFiles] =
        await Promise.all([
          db.count(STORES.AMC_CALCULATIONS),
          db.count(STORES.WARRANTY_CALCULATIONS),
          db.count(STORES.EXCEL_FILES),
        ]);

      // Estimate storage usage
      const allData = await Promise.all([
        db.getAll(STORES.AMC_CALCULATIONS),
        db.getAll(STORES.WARRANTY_CALCULATIONS),
        db.getAll(STORES.EXCEL_FILES),
      ]);

      const totalSize = allData.flat().reduce((size, item) => {
        return size + JSON.stringify(item).length;
      }, 0);

      return {
        counts: {
          amcCalculations,
          warrantyCalculations,
          excelFiles,
        },
        estimatedSize: totalSize,
        estimatedSizeMB: Math.round((totalSize / (1024 * 1024)) * 100) / 100,
      };
    } catch (error) {
      console.error("Error getting cache stats:", error);
      return {
        counts: { amcCalculations: 0, warrantyCalculations: 0, excelFiles: 0 },
        estimatedSize: 0,
        estimatedSizeMB: 0,
      };
    }
  },

  // Clean old cache entries
  async cleanup(olderThanDays = 30) {
    try {
      const cutoffTime = Date.now() - olderThanDays * 24 * 60 * 60 * 1000;
      const db = await initDB();

      let deletedCount = 0;

      // Clean AMC calculations
      const amcTx = db.transaction(STORES.AMC_CALCULATIONS, "readwrite");
      const amcIndex = amcTx.store.index("timestamp");
      const amcCursor = await amcIndex.openCursor(
        IDBKeyRange.upperBound(cutoffTime)
      );

      while (amcCursor) {
        await amcCursor.delete();
        deletedCount++;
        await amcCursor.continue();
      }

      // Clean Excel files
      const fileTx = db.transaction(STORES.EXCEL_FILES, "readwrite");
      const allFiles = await fileTx.store.getAll();

      for (const file of allFiles) {
        if (new Date(file.uploadDate).getTime() < cutoffTime) {
          await fileTx.store.delete(file.id);
          deletedCount++;
        }
      }

      console.log(
        `🧹 Cache cleanup completed: ${deletedCount} entries deleted`
      );
      return deletedCount;
    } catch (error) {
      console.error("Error during cache cleanup:", error);
      return 0;
    }
  },

  // Clear all cache
  async clearAll() {
    try {
      const db = await initDB();

      await Promise.all([
        db.clear(STORES.AMC_CALCULATIONS),
        db.clear(STORES.WARRANTY_CALCULATIONS),
        db.clear(STORES.EXCEL_FILES),
        db.clear(STORES.CALCULATION_METADATA),
      ]);

      console.log("🗑️ All cache cleared");
    } catch (error) {
      console.error("Error clearing cache:", error);
      throw error;
    }
  },
};

// Export utility functions
export const cacheUtils = {
  generateFileHash,
  isOnline: () => navigator.onLine,

  // Check if browser supports IndexedDB
  isSupported: () => {
    return "indexedDB" in window;
  },

  // Get cache size estimate
  async getQuotaInfo() {
    if ("storage" in navigator && "estimate" in navigator.storage) {
      const estimate = await navigator.storage.estimate();
      return {
        quota: estimate.quota,
        usage: estimate.usage,
        usagePercentage: Math.round((estimate.usage / estimate.quota) * 100),
      };
    }
    return null;
  },
};

// Initialize database on module load
initDB().catch(console.error);

const cacheSystem = {
  amcCache,
  fileCache,
  cacheManager,
  cacheUtils,
};

export default cacheSystem;
