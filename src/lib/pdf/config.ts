/**
 * PDF Library Configuration
 * Requirements: 1.4, 5.6
 * 
 * Configures PDF processing libraries for client-side processing.
 * Libraries are lazy-loaded to optimize initial page load.
 */

/**
 * PDF.js configuration
 */
export const PDFJS_CONFIG = {
  // Worker source path - loaded from public directory
  workerSrc: '/workers/pdf.worker.min.mjs',
  // CMap URL for character mapping
  cMapUrl: '/pdfjs/cmaps/',
  cMapPacked: true,
  // Standard fonts URL
  standardFontDataUrl: '/pdfjs/standard_fonts/',
} as const;

/**
 * PDF processing configuration
 */
export const PDF_PROCESSING_CONFIG = {
  // Maximum file size for processing (100MB)
  maxFileSize: 100 * 1024 * 1024,
  // Timeout for processing operations (5 minutes)
  processingTimeout: 5 * 60 * 1000,
  // Chunk size for large file processing
  chunkSize: 1024 * 1024, // 1MB
  // Enable Web Worker for processing
  useWebWorker: true,
  // Maximum concurrent operations
  maxConcurrentOperations: 3,
} as const;

/**
 * Supported PDF versions
 */
export const SUPPORTED_PDF_VERSIONS = ['1.0', '1.1', '1.2', '1.3', '1.4', '1.5', '1.6', '1.7', '2.0'] as const;

/**
 * Library loading state
 */
export interface LibraryState {
  pdfLib: boolean;
  pdfjs: boolean;
  loading: boolean;
  error: string | null;
}

/**
 * Initial library state
 */
export const initialLibraryState: LibraryState = {
  pdfLib: false,
  pdfjs: false,
  loading: false,
  error: null,
};
