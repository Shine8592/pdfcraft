'use client';

import React, { useState, useCallback, useRef } from 'react';
import { useTranslations } from 'next-intl';
import { FileUploader } from '../FileUploader';
import { ProcessingProgress, ProcessingStatus } from '../ProcessingProgress';
import { DownloadButton } from '../DownloadButton';
import { Button } from '@/components/ui/Button';
import { Card } from '@/components/ui/Card';
import { createNUpPDF, type NUpOptions } from '@/lib/pdf/processors/n-up';
import { configurePdfjsWorker } from '@/lib/pdf/loader';
import type { ProcessOutput } from '@/types/pdf';

export interface NUpPDFToolProps {
  /** Custom class name */
  className?: string;
}

/**
 * NUpPDFTool Component
 * Requirements: 5.1, 5.2
 * 
 * Provides the UI for creating N-Up PDF layouts.
 */
export function NUpPDFTool({ className = '' }: NUpPDFToolProps) {
  const t = useTranslations('common');
  const tTools = useTranslations('tools');
  
  // State
  const [file, setFile] = useState<File | null>(null);
  const [totalPages, setTotalPages] = useState<number>(0);
  const [status, setStatus] = useState<ProcessingStatus>('idle');
  const [progress, setProgress] = useState(0);
  const [progressMessage, setProgressMessage] = useState('');
  const [result, setResult] = useState<Blob | null>(null);
  const [error, setError] = useState<string | null>(null);
  
  // Options
  const [pagesPerSheet, setPagesPerSheet] = useState<2 | 4 | 9 | 16>(4);
  const [pageSize, setPageSize] = useState<'A4' | 'Letter' | 'Legal' | 'A3'>('A4');
  const [orientation, setOrientation] = useState<'portrait' | 'landscape' | 'auto'>('auto');
  const [useMargins, setUseMargins] = useState(true);
  const [addBorder, setAddBorder] = useState(false);
  const [borderColor, setBorderColor] = useState('#000000');
  
  // Ref for cancellation
  const cancelledRef = useRef(false);

  /**
   * Load PDF to get page count
   */
  const loadPdfInfo = useCallback(async (pdfFile: File) => {
    try {
      const pdfjsLib = await import('pdfjs-dist');
      configurePdfjsWorker(pdfjsLib);
      
      const arrayBuffer = await pdfFile.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      
      setTotalPages(pdf.numPages);
    } catch (err) {
      console.error('Failed to load PDF info:', err);
      setError('Failed to load PDF. The file may be corrupted or encrypted.');
    }
  }, []);

  /**
   * Handle file selected from uploader
   */
  const handleFilesSelected = useCallback((files: File[]) => {
    if (files.length > 0) {
      const selectedFile = files[0];
      setFile(selectedFile);
      setError(null);
      setResult(null);
      loadPdfInfo(selectedFile);
    }
  }, [loadPdfInfo]);

  /**
   * Handle file upload error
   */
  const handleUploadError = useCallback((errorMessage: string) => {
    setError(errorMessage);
  }, []);

  /**
   * Clear file and reset state
   */
  const handleClearFile = useCallback(() => {
    setFile(null);
    setTotalPages(0);
    setResult(null);
    setError(null);
    setStatus('idle');
    setProgress(0);
  }, []);

  /**
   * Handle N-Up operation
   */
  const handleCreateNUp = useCallback(async () => {
    if (!file) {
      setError('Please upload a PDF file first.');
      return;
    }

    cancelledRef.current = false;
    setStatus('processing');
    setProgress(0);
    setError(null);
    setResult(null);

    const options: Partial<NUpOptions> = {
      pagesPerSheet,
      pageSize,
      orientation,
      useMargins,
      addBorder,
      borderColor,
    };

    try {
      const output: ProcessOutput = await createNUpPDF(
        file,
        options,
        (prog, message) => {
          if (!cancelledRef.current) {
            setProgress(prog);
            setProgressMessage(message || '');
          }
        }
      );

      if (cancelledRef.current) {
        setStatus('idle');
        return;
      }

      if (output.success && output.result) {
        setResult(output.result as Blob);
        setStatus('complete');
      } else {
        setError(output.error?.message || 'Failed to create N-Up PDF.');
        setStatus('error');
      }
    } catch (err) {
      if (!cancelledRef.current) {
        setError(err instanceof Error ? err.message : 'An unexpected error occurred.');
        setStatus('error');
      }
    }
  }, [file, pagesPerSheet, pageSize, orientation, useMargins, addBorder, borderColor]);

  /**
   * Handle cancel operation
   */
  const handleCancel = useCallback(() => {
    cancelledRef.current = true;
    setStatus('idle');
    setProgress(0);
  }, []);

  /**
   * Format file size
   */
  const formatSize = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  const isProcessing = status === 'processing' || status === 'uploading';
  const canProcess = file && totalPages > 0 && !isProcessing;
  const outputSheets = Math.ceil(totalPages / pagesPerSheet);

  return (
    <div className={`space-y-6 ${className}`.trim()}>
      {/* File Upload Area */}
      {!file && (
        <FileUploader
          accept={['application/pdf', '.pdf']}
          multiple={false}
          maxFiles={1}
          maxSize={500 * 1024 * 1024}
          onFilesSelected={handleFilesSelected}
          onError={handleUploadError}
          disabled={isProcessing}
          label={tTools('nUpPdf.uploadLabel') || 'Upload PDF File'}
          description={tTools('nUpPdf.uploadDescription') || 'Drag and drop a PDF file here, or click to browse.'}
        />
      )}

      {/* Error Message */}
      {error && (
        <div 
          className="p-4 rounded-[var(--radius-md)] bg-red-50 border border-red-200 text-red-700"
          role="alert"
        >
          <p className="text-sm">{error}</p>
        </div>
      )}

      {/* File Info */}
      {file && (
        <Card variant="outlined">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <svg className="w-10 h-10 text-red-500" viewBox="0 0 24 24" fill="currentColor">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z" />
                <path d="M14 2v6h6" fill="white" />
                <text x="7" y="17" fontSize="6" fill="white" fontWeight="bold">PDF</text>
              </svg>
              <div>
                <p className="font-medium text-[hsl(var(--color-foreground))]">{file.name}</p>
                <p className="text-sm text-[hsl(var(--color-muted-foreground))]">
                  {formatSize(file.size)} • {totalPages} {totalPages === 1 ? 'page' : 'pages'}
                </p>
              </div>
            </div>
            <Button
              variant="ghost"
              size="sm"
              onClick={handleClearFile}
              disabled={isProcessing}
            >
              {t('buttons.remove') || 'Remove'}
            </Button>
          </div>
        </Card>
      )}

      {/* Options */}
      {file && totalPages > 0 && (
        <Card variant="outlined" size="lg">
          <h3 className="text-lg font-medium text-[hsl(var(--color-foreground))] mb-4">
            {tTools('nUpPdf.optionsTitle') || 'N-Up Options'}
          </h3>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {/* Pages per sheet */}
            <div>
              <label htmlFor="pagesPerSheet" className="block text-sm font-medium text-[hsl(var(--color-foreground))] mb-1">
                {tTools('nUpPdf.pagesPerSheet') || 'Pages per Sheet'}
              </label>
              <select
                id="pagesPerSheet"
                value={pagesPerSheet}
                onChange={(e) => setPagesPerSheet(parseInt(e.target.value) as 2 | 4 | 9 | 16)}
                disabled={isProcessing}
                className="w-full px-3 py-2 rounded-[var(--radius-md)] border border-[hsl(var(--color-border))] bg-[hsl(var(--color-background))] text-[hsl(var(--color-foreground))] focus:outline-none focus:ring-2 focus:ring-[hsl(var(--color-primary))]"
              >
                <option value={2}>2-up (2×1)</option>
                <option value={4}>4-up (2×2)</option>
                <option value={9}>9-up (3×3)</option>
                <option value={16}>16-up (4×4)</option>
              </select>
            </div>

            {/* Page size */}
            <div>
              <label htmlFor="pageSize" className="block text-sm font-medium text-[hsl(var(--color-foreground))] mb-1">
                {tTools('nUpPdf.pageSize') || 'Output Page Size'}
              </label>
              <select
                id="pageSize"
                value={pageSize}
                onChange={(e) => setPageSize(e.target.value as 'A4' | 'Letter' | 'Legal' | 'A3')}
                disabled={isProcessing}
                className="w-full px-3 py-2 rounded-[var(--radius-md)] border border-[hsl(var(--color-border))] bg-[hsl(var(--color-background))] text-[hsl(var(--color-foreground))] focus:outline-none focus:ring-2 focus:ring-[hsl(var(--color-primary))]"
              >
                <option value="A4">A4</option>
                <option value="Letter">Letter</option>
                <option value="Legal">Legal</option>
                <option value="A3">A3</option>
              </select>
            </div>

            {/* Orientation */}
            <div>
              <label htmlFor="orientation" className="block text-sm font-medium text-[hsl(var(--color-foreground))] mb-1">
                {tTools('nUpPdf.orientation') || 'Orientation'}
              </label>
              <select
                id="orientation"
                value={orientation}
                onChange={(e) => setOrientation(e.target.value as 'portrait' | 'landscape' | 'auto')}
                disabled={isProcessing}
                className="w-full px-3 py-2 rounded-[var(--radius-md)] border border-[hsl(var(--color-border))] bg-[hsl(var(--color-background))] text-[hsl(var(--color-foreground))] focus:outline-none focus:ring-2 focus:ring-[hsl(var(--color-primary))]"
              >
                <option value="auto">{t('options.auto') || 'Auto'}</option>
                <option value="portrait">{t('options.portrait') || 'Portrait'}</option>
                <option value="landscape">{t('options.landscape') || 'Landscape'}</option>
              </select>
            </div>

            {/* Margins */}
            <div className="flex items-center gap-2">
              <input
                type="checkbox"
                id="useMargins"
                checked={useMargins}
                onChange={(e) => setUseMargins(e.target.checked)}
                disabled={isProcessing}
                className="w-4 h-4 rounded border-[hsl(var(--color-border))]"
              />
              <label htmlFor="useMargins" className="text-sm text-[hsl(var(--color-foreground))]">
                {tTools('nUpPdf.useMargins') || 'Add margins and gutters'}
              </label>
            </div>

            {/* Border */}
            <div className="flex items-center gap-2">
              <input
                type="checkbox"
                id="addBorder"
                checked={addBorder}
                onChange={(e) => setAddBorder(e.target.checked)}
                disabled={isProcessing}
                className="w-4 h-4 rounded border-[hsl(var(--color-border))]"
              />
              <label htmlFor="addBorder" className="text-sm text-[hsl(var(--color-foreground))]">
                {tTools('nUpPdf.addBorder') || 'Add border around pages'}
              </label>
            </div>

            {/* Border color */}
            {addBorder && (
              <div>
                <label htmlFor="borderColor" className="block text-sm font-medium text-[hsl(var(--color-foreground))] mb-1">
                  {tTools('nUpPdf.borderColor') || 'Border Color'}
                </label>
                <input
                  type="color"
                  id="borderColor"
                  value={borderColor}
                  onChange={(e) => setBorderColor(e.target.value)}
                  disabled={isProcessing}
                  className="w-full h-10 rounded-[var(--radius-md)] border border-[hsl(var(--color-border))] cursor-pointer"
                />
              </div>
            )}
          </div>

          {/* Preview info */}
          <div className="mt-4 p-3 rounded-[var(--radius-md)] bg-[hsl(var(--color-muted))]">
            <p className="text-sm text-[hsl(var(--color-muted-foreground))]">
              {tTools('nUpPdf.previewInfo', { pages: totalPages, sheets: outputSheets, perSheet: pagesPerSheet }) || 
                `${totalPages} pages will be arranged into ${outputSheets} sheet${outputSheets !== 1 ? 's' : ''} with ${pagesPerSheet} pages per sheet.`
              }
            </p>
          </div>
        </Card>
      )}

      {/* Processing Progress */}
      {isProcessing && (
        <ProcessingProgress
          progress={progress}
          status={status}
          message={progressMessage}
          onCancel={handleCancel}
          showPercentage
        />
      )}

      {/* Action Buttons */}
      {file && (
        <div className="flex flex-wrap items-center gap-4">
          <Button
            variant="primary"
            size="lg"
            onClick={handleCreateNUp}
            disabled={!canProcess}
            loading={isProcessing}
          >
            {isProcessing 
              ? (t('status.processing') || 'Processing...') 
              : (tTools('nUpPdf.createButton') || `Create ${pagesPerSheet}-Up PDF`)
            }
          </Button>

          {result && (
            <DownloadButton
              file={result}
              filename={file.name.replace('.pdf', `_${pagesPerSheet}-up.pdf`)}
              variant="secondary"
              size="lg"
              showFileSize
            />
          )}
        </div>
      )}

      {/* Success Message */}
      {status === 'complete' && result && (
        <div 
          className="p-4 rounded-[var(--radius-md)] bg-green-50 border border-green-200 text-green-700"
          role="status"
        >
          <p className="text-sm font-medium">
            {tTools('nUpPdf.successMessage') || 'N-Up PDF created successfully! Click the download button to save your file.'}
          </p>
        </div>
      )}
    </div>
  );
}

export default NUpPDFTool;
