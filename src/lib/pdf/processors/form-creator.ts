/**
 * PDF Form Creator Processor
 * Requirements: 5.1
 */

import type { ProcessInput, ProcessOutput, ProgressCallback } from '@/types/pdf';
import { PDFErrorCode } from '@/types/pdf';
import { BasePDFProcessor } from '../processor';
import { loadPdfLib } from '../loader';

export interface FormField {
  type: 'text' | 'checkbox' | 'dropdown' | 'radio';
  name: string;
  pageNumber: number;
  x: number;
  y: number;
  width: number;
  height: number;
  options?: string[]; // For dropdown/radio
  defaultValue?: string | boolean;
  required?: boolean;
  multiline?: boolean;
}

export interface FormCreatorOptions {
  fields: FormField[];
}

export class FormCreatorProcessor extends BasePDFProcessor {
  async process(input: ProcessInput, onProgress?: ProgressCallback): Promise<ProcessOutput> {
    this.reset();
    this.onProgress = onProgress;

    const { files, options } = input;
    const formOptions = options as unknown as FormCreatorOptions;

    if (files.length !== 1) {
      return this.createErrorOutput(PDFErrorCode.INVALID_OPTIONS, 'Exactly 1 PDF file is required.');
    }

    if (!formOptions.fields || formOptions.fields.length === 0) {
      return this.createErrorOutput(PDFErrorCode.INVALID_OPTIONS, 'At least one form field is required.');
    }

    try {
      this.updateProgress(10, 'Loading PDF library...');
      const pdfLib = await loadPdfLib();

      this.updateProgress(20, 'Loading PDF...');
      const file = files[0];
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await pdfLib.PDFDocument.load(arrayBuffer, { ignoreEncryption: true });

      const form = pdf.getForm();
      const totalPages = pdf.getPageCount();

      this.updateProgress(30, 'Creating form fields...');

      let createdCount = 0;
      for (let i = 0; i < formOptions.fields.length; i++) {
        if (this.checkCancelled()) {
          return this.createErrorOutput(PDFErrorCode.PROCESSING_CANCELLED, 'Processing was cancelled.');
        }

        const fieldDef = formOptions.fields[i];
        const pageIndex = fieldDef.pageNumber - 1;

        if (pageIndex < 0 || pageIndex >= totalPages) continue;

        const page = pdf.getPage(pageIndex);

        try {
          switch (fieldDef.type) {
            case 'text': {
              const textField = form.createTextField(fieldDef.name);
              textField.addToPage(page, {
                x: fieldDef.x,
                y: fieldDef.y,
                width: fieldDef.width,
                height: fieldDef.height,
              });
              if (fieldDef.defaultValue) {
                textField.setText(String(fieldDef.defaultValue));
              }
              if (fieldDef.multiline) {
                textField.enableMultiline();
              }
              createdCount++;
              break;
            }
            case 'checkbox': {
              const checkBox = form.createCheckBox(fieldDef.name);
              checkBox.addToPage(page, {
                x: fieldDef.x,
                y: fieldDef.y,
                width: fieldDef.width,
                height: fieldDef.height,
              });
              if (fieldDef.defaultValue) {
                checkBox.check();
              }
              createdCount++;
              break;
            }
            case 'dropdown': {
              const dropdown = form.createDropdown(fieldDef.name);
              dropdown.addToPage(page, {
                x: fieldDef.x,
                y: fieldDef.y,
                width: fieldDef.width,
                height: fieldDef.height,
              });
              if (fieldDef.options) {
                dropdown.addOptions(fieldDef.options);
                if (fieldDef.defaultValue) {
                  dropdown.select(String(fieldDef.defaultValue));
                }
              }
              createdCount++;
              break;
            }
            case 'radio': {
              const radioGroup = form.createRadioGroup(fieldDef.name);
              if (fieldDef.options) {
                fieldDef.options.forEach((option, idx) => {
                  radioGroup.addOptionToPage(option, page, {
                    x: fieldDef.x,
                    y: fieldDef.y - (idx * (fieldDef.height + 5)),
                    width: fieldDef.width,
                    height: fieldDef.height,
                  });
                });
                if (fieldDef.defaultValue) {
                  radioGroup.select(String(fieldDef.defaultValue));
                }
              }
              createdCount++;
              break;
            }
          }
        } catch (fieldError) {
          console.warn(`Failed to create field ${fieldDef.name}:`, fieldError);
        }

        this.updateProgress(30 + (60 * (i + 1) / formOptions.fields.length), `Creating field ${i + 1}...`);
      }

      this.updateProgress(95, 'Saving PDF...');
      const pdfBytes = await pdf.save();
      const blob = new Blob([new Uint8Array(pdfBytes)], { type: 'application/pdf' });

      this.updateProgress(100, 'Complete!');
      return this.createSuccessOutput(blob, file.name.replace('.pdf', '_form.pdf'), { 
        fieldsCreated: createdCount,
      });

    } catch (error) {
      return this.createErrorOutput(PDFErrorCode.PROCESSING_FAILED, 'Failed to create form.', error instanceof Error ? error.message : 'Unknown error');
    }
  }

  protected getAcceptedTypes(): string[] {
    return ['application/pdf'];
  }
}

export function createFormCreatorProcessor(): FormCreatorProcessor {
  return new FormCreatorProcessor();
}

export async function createForm(file: File, options: FormCreatorOptions, onProgress?: ProgressCallback): Promise<ProcessOutput> {
  const processor = createFormCreatorProcessor();
  return processor.process({ files: [file], options: options as unknown as Record<string, unknown> }, onProgress);
}
