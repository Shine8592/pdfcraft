/**
 * PDF Sign Processor
 * Requirements: 5.1
 */

import type { ProcessInput, ProcessOutput, ProgressCallback } from '@/types/pdf';
import { PDFErrorCode } from '@/types/pdf';
import { BasePDFProcessor } from '../processor';
import { loadPdfLib } from '../loader';

export interface SignatureItem {
  type: 'draw' | 'type' | 'image';
  data: string; // Base64 image data or text
  pageNumber: number;
  x: number;
  y: number;
  width?: number;
  height?: number;
  fontFamily?: string;
  fontSize?: number;
}

export interface SignOptions {
  signatures: SignatureItem[];
}

export class SignProcessor extends BasePDFProcessor {
  async process(input: ProcessInput, onProgress?: ProgressCallback): Promise<ProcessOutput> {
    this.reset();
    this.onProgress = onProgress;

    const { files, options } = input;
    const signOptions = options as unknown as SignOptions;

    if (files.length !== 1) {
      return this.createErrorOutput(PDFErrorCode.INVALID_OPTIONS, 'Exactly 1 PDF file is required.');
    }

    if (!signOptions.signatures || signOptions.signatures.length === 0) {
      return this.createErrorOutput(PDFErrorCode.INVALID_OPTIONS, 'At least one signature is required.');
    }

    try {
      this.updateProgress(10, 'Loading PDF library...');
      const pdfLib = await loadPdfLib();

      this.updateProgress(20, 'Loading PDF...');
      const file = files[0];
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await pdfLib.PDFDocument.load(arrayBuffer, { ignoreEncryption: true });

      const totalPages = pdf.getPageCount();

      this.updateProgress(30, 'Adding signatures...');

      for (let i = 0; i < signOptions.signatures.length; i++) {
        if (this.checkCancelled()) {
          return this.createErrorOutput(PDFErrorCode.PROCESSING_CANCELLED, 'Processing was cancelled.');
        }

        const sig = signOptions.signatures[i];
        const pageIndex = sig.pageNumber - 1;

        if (pageIndex < 0 || pageIndex >= totalPages) {
          continue;
        }

        const page = pdf.getPage(pageIndex);

        if (sig.type === 'type') {
          const font = await pdf.embedFont(pdfLib.StandardFonts.Courier);
          page.drawText(sig.data, {
            x: sig.x,
            y: sig.y,
            size: sig.fontSize || 24,
            font,
          });
        } else if (sig.type === 'draw' || sig.type === 'image') {
          // Embed image signature
          try {
            const imageData = sig.data.replace(/^data:image\/\w+;base64,/, '');
            const imageBytes = Uint8Array.from(atob(imageData), c => c.charCodeAt(0));
            
            let image;
            if (sig.data.includes('image/png')) {
              image = await pdf.embedPng(imageBytes);
            } else {
              image = await pdf.embedJpg(imageBytes);
            }

            const width = sig.width || 150;
            const height = sig.height || 50;

            page.drawImage(image, {
              x: sig.x,
              y: sig.y,
              width,
              height,
            });
          } catch (imgError) {
            console.error('Failed to embed signature image:', imgError);
          }
        }

        this.updateProgress(30 + (60 * (i + 1) / signOptions.signatures.length), `Adding signature ${i + 1}...`);
      }

      this.updateProgress(95, 'Saving PDF...');
      const pdfBytes = await pdf.save();
      const blob = new Blob([new Uint8Array(pdfBytes)], { type: 'application/pdf' });

      this.updateProgress(100, 'Complete!');
      return this.createSuccessOutput(blob, file.name.replace('.pdf', '_signed.pdf'), { 
        signatureCount: signOptions.signatures.length 
      });

    } catch (error) {
      return this.createErrorOutput(PDFErrorCode.PROCESSING_FAILED, 'Failed to sign PDF.', error instanceof Error ? error.message : 'Unknown error');
    }
  }

  protected getAcceptedTypes(): string[] {
    return ['application/pdf'];
  }
}

export function createSignProcessor(): SignProcessor {
  return new SignProcessor();
}

export async function signPDF(file: File, options: SignOptions, onProgress?: ProgressCallback): Promise<ProcessOutput> {
  const processor = createSignProcessor();
  return processor.process({ files: [file], options: options as unknown as Record<string, unknown> }, onProgress);
}
