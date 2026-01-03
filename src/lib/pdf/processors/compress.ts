/**
 * PDF Compress Processor
 * Requirements: 5.1
 * 
 * Implements PDF compression functionality using coherentpdf for better compression.
 * Supports different quality levels and optimization options.
 */

import type {
  ProcessInput,
  ProcessOutput,
  ProgressCallback,
} from '@/types/pdf';
import { PDFErrorCode } from '@/types/pdf';
import { BasePDFProcessor } from '../processor';

/**
 * Compression quality levels
 */
export type CompressionQuality = 'low' | 'medium' | 'high' | 'maximum';

/**
 * Compress PDF options
 */
export interface CompressPDFOptions {
  /** Compression quality level */
  quality: CompressionQuality;
  /** Remove metadata to reduce size */
  removeMetadata: boolean;
  /** Optimize images in the PDF */
  optimizeImages: boolean;
  /** Remove unused objects */
  removeUnusedObjects: boolean;
}

/**
 * Default compress options
 */
const DEFAULT_COMPRESS_OPTIONS: CompressPDFOptions = {
  quality: 'medium',
  removeMetadata: false,
  optimizeImages: true,
  removeUnusedObjects: true,
};

/**
 * Worker message types
 */
interface WorkerProgressMessage {
  status: 'progress';
  progress: number;
}

interface WorkerSuccessMessage {
  status: 'success';
  pdfBytes: ArrayBuffer;
  originalSize: number;
  compressedSize: number;
}

interface WorkerErrorMessage {
  status: 'error';
  message: string;
}

type WorkerMessage = WorkerProgressMessage | WorkerSuccessMessage | WorkerErrorMessage;

/**
 * Compress PDF Processor
 * Compresses PDF files to reduce file size using coherentpdf.
 */
export class CompressPDFProcessor extends BasePDFProcessor {
  private worker: Worker | null = null;

  /**
   * Process PDF file and compress it
   */
  async process(
    input: ProcessInput,
    onProgress?: ProgressCallback
  ): Promise<ProcessOutput> {
    this.reset();
    this.onProgress = onProgress;

    const { files, options } = input;
    const compressOptions: CompressPDFOptions = {
      ...DEFAULT_COMPRESS_OPTIONS,
      ...(options as Partial<CompressPDFOptions>),
    };

    // Validate we have exactly 1 file
    if (files.length !== 1) {
      return this.createErrorOutput(
        PDFErrorCode.INVALID_OPTIONS,
        'Please provide exactly one PDF file to compress.',
        `Received ${files.length} file(s).`
      );
    }

    const file = files[0];

    try {
      this.updateProgress(5, 'Loading PDF file...');

      // Read file as ArrayBuffer
      const arrayBuffer = await file.arrayBuffer();
      const originalSize = arrayBuffer.byteLength;

      if (this.checkCancelled()) {
        return this.createErrorOutput(
          PDFErrorCode.PROCESSING_CANCELLED,
          'Processing was cancelled.'
        );
      }

      this.updateProgress(10, 'Starting compression...');

      // Process using worker
      const result = await this.compressWithWorker(arrayBuffer, compressOptions);

      if (this.checkCancelled()) {
        return this.createErrorOutput(
          PDFErrorCode.PROCESSING_CANCELLED,
          'Processing was cancelled.'
        );
      }

      const compressedSize = result.compressedSize;
      const compressionRatio = ((originalSize - compressedSize) / originalSize * 100).toFixed(1);

      // Create blob from the result
      const blob = new Blob([result.pdfBytes], { type: 'application/pdf' });

      this.updateProgress(100, 'Complete!');

      // Generate output filename
      const outputFilename = generateCompressedFilename(file.name);

      return this.createSuccessOutput(blob, outputFilename, {
        originalSize,
        compressedSize,
        compressionRatio: `${compressionRatio}%`,
        quality: compressOptions.quality,
      });

    } catch (error) {
      // Clean up worker on error
      this.terminateWorker();
      
      if (error instanceof Error && error.message.includes('encrypt')) {
        return this.createErrorOutput(
          PDFErrorCode.PDF_ENCRYPTED,
          'The PDF file is encrypted.',
          'Please decrypt the file before compressing.'
        );
      }
      
      return this.createErrorOutput(
        PDFErrorCode.PROCESSING_FAILED,
        'Failed to compress PDF file.',
        error instanceof Error ? error.message : 'Unknown error'
      );
    }
  }

  /**
   * Compress PDF using web worker
   */
  private compressWithWorker(
    pdfData: ArrayBuffer,
    options: CompressPDFOptions
  ): Promise<{ pdfBytes: ArrayBuffer; compressedSize: number }> {
    return new Promise((resolve, reject) => {
      try {
        this.worker = new Worker('/workers/compress.worker.js');

        this.worker.onmessage = (e: MessageEvent<WorkerMessage>) => {
          const data = e.data;

          if (data.status === 'progress') {
            // Map worker progress (20-100) to overall progress (10-95)
            const mappedProgress = 10 + (data.progress / 100) * 85;
            this.updateProgress(mappedProgress, 'Compressing PDF...');
          } else if (data.status === 'success') {
            this.terminateWorker();
            resolve({
              pdfBytes: data.pdfBytes,
              compressedSize: data.compressedSize,
            });
          } else if (data.status === 'error') {
            this.terminateWorker();
            reject(new Error(data.message));
          }
        };

        this.worker.onerror = (error) => {
          this.terminateWorker();
          reject(new Error(`Worker error: ${error.message}`));
        };

        // Send data to worker
        this.worker.postMessage(
          {
            command: 'compress',
            pdfData: pdfData,
            options: {
              quality: options.quality,
              removeMetadata: options.removeMetadata,
            },
          },
          [pdfData]
        );
      } catch (error) {
        this.terminateWorker();
        reject(error);
      }
    });
  }

  /**
   * Terminate the worker
   */
  private terminateWorker(): void {
    if (this.worker) {
      this.worker.terminate();
      this.worker = null;
    }
  }

  /**
   * Cancel processing
   */
  cancel(): void {
    super.cancel();
    this.terminateWorker();
  }

  /**
   * Get accepted file types for compress processor
   */
  protected getAcceptedTypes(): string[] {
    return ['application/pdf'];
  }
}

/**
 * Generate a filename for the compressed PDF
 */
function generateCompressedFilename(originalName: string): string {
  const lastDot = originalName.lastIndexOf('.');
  const baseName = lastDot === -1 ? originalName : originalName.slice(0, lastDot);
  return `${baseName}_compressed.pdf`;
}

/**
 * Create a new instance of the compress processor
 */
export function createCompressProcessor(): CompressPDFProcessor {
  return new CompressPDFProcessor();
}

/**
 * Compress a PDF file (convenience function)
 */
export async function compressPDF(
  file: File,
  options?: Partial<CompressPDFOptions>,
  onProgress?: ProgressCallback
): Promise<ProcessOutput> {
  const processor = createCompressProcessor();
  return processor.process(
    {
      files: [file],
      options: options || {},
    },
    onProgress
  );
}
