/**
 * PDF Table of Contents Processor
 * Uses coherentpdf worker to generate clickable TOC from PDF bookmarks
 * Requirements: 5.1
 */

import type { ProcessInput, ProcessOutput, ProgressCallback } from '@/types/pdf';
import { PDFErrorCode } from '@/types/pdf';
import { BasePDFProcessor } from '../processor';

export interface TOCOptions {
  title?: string;
  fontSize?: number;
  fontFamily?: number; // 0=Times, 1=Helvetica, 2=Courier
  addBookmark?: boolean;
}

export class TableOfContentsProcessor extends BasePDFProcessor {
  private worker: Worker | null = null;

  async process(input: ProcessInput, onProgress?: ProgressCallback): Promise<ProcessOutput> {
    this.reset();
    this.onProgress = onProgress;

    const { files, options } = input;
    const tocOptions = options as TOCOptions;

    if (files.length !== 1) {
      return this.createErrorOutput(PDFErrorCode.INVALID_OPTIONS, 'Exactly 1 PDF file is required.');
    }

    try {
      this.updateProgress(10, 'Loading PDF...');
      const file = files[0];
      const arrayBuffer = await file.arrayBuffer();

      this.updateProgress(30, 'Generating table of contents...');

      const result = await this.generateTOCWithWorker(
        arrayBuffer,
        tocOptions.title || 'Table of Contents',
        tocOptions.fontSize || 12,
        tocOptions.fontFamily ?? 1, // Default to Helvetica
        tocOptions.addBookmark ?? true
      );

      if (result.status === 'error') {
        return this.createErrorOutput(
          PDFErrorCode.PROCESSING_FAILED,
          result.message || 'Failed to generate table of contents.'
        );
      }

      this.updateProgress(90, 'Saving PDF...');
      const blob = new Blob([new Uint8Array(result.pdfBytes)], { type: 'application/pdf' });

      this.updateProgress(100, 'Complete!');
      return this.createSuccessOutput(blob, file.name.replace('.pdf', '_with_toc.pdf'), {});

    } catch (error) {
      return this.createErrorOutput(
        PDFErrorCode.PROCESSING_FAILED,
        'Failed to generate table of contents.',
        error instanceof Error ? error.message : 'Unknown error'
      );
    } finally {
      this.terminateWorker();
    }
  }

  private generateTOCWithWorker(
    pdfData: ArrayBuffer,
    title: string,
    fontSize: number,
    fontFamily: number,
    addBookmark: boolean
  ): Promise<{ status: 'success'; pdfBytes: ArrayBuffer } | { status: 'error'; message: string }> {
    return new Promise((resolve) => {
      this.worker = new Worker('/workers/table-of-contents.worker.js');

      this.worker.onmessage = (e) => {
        resolve(e.data);
      };

      this.worker.onerror = (error) => {
        resolve({
          status: 'error',
          message: error.message || 'Worker error occurred',
        });
      };

      this.worker.postMessage(
        {
          command: 'generate-toc',
          pdfData,
          title,
          fontSize,
          fontFamily,
          addBookmark,
        },
        [pdfData]
      );
    });
  }

  private terminateWorker(): void {
    if (this.worker) {
      this.worker.terminate();
      this.worker = null;
    }
  }

  protected getAcceptedTypes(): string[] {
    return ['application/pdf'];
  }
}

export function createTableOfContentsProcessor(): TableOfContentsProcessor {
  return new TableOfContentsProcessor();
}

export async function generateTableOfContents(
  file: File,
  options: TOCOptions,
  onProgress?: ProgressCallback
): Promise<ProcessOutput> {
  const processor = createTableOfContentsProcessor();
  return processor.process({ files: [file], options: options as unknown as Record<string, unknown> }, onProgress);
}
