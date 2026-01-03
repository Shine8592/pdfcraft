'use client';

import React, { useState, useCallback, useRef, useEffect } from 'react';
import { useTranslations } from 'next-intl';
import { FileUploader } from '../FileUploader';
import { ProcessingProgress, ProcessingStatus } from '../ProcessingProgress';
import { DownloadButton } from '../DownloadButton';
import { Button } from '@/components/ui/Button';
import { Card } from '@/components/ui/Card';
import { 
  splitPDF, 
  parsePageRanges, 
  createSplitEveryPage,
  createSplitEveryNPages 
} from '@/lib/pdf';
import { configurePdfjsWorker } from '@/lib/pdf/loader';
import type { SplitOptions, PageRange, ProcessOutput } from '@/types/pdf';

export interface SplitPDFToolProps {
  /** Custom class name */
  className?: string;
}

type SplitMode = 'ranges' | 'every-page' | 'every-n-pages';

interface PagePreview {
  pageNumber: number;
  thumbnail?: string;
}

/**
 * SplitPDFTool Component
 * Requirements: 5.1, 5.2
 * 
 * Provides the UI for splitting PDF files with page range input and preview.
 */
export function SplitPDFTool({ className = '' }: SplitPDFToolProps) {
  const t = useTranslations('common');
  const tTools = useTranslations('tools');
  
  // State
  const [file, setFile] = useState<File | null>(null);
  const [totalPages, setTotalPages] = useState<number>(0);
  const [status, setStatus] = useState<ProcessingStatus>('idle');
  const [progress, setProgress] = useState(0);
  const [progressMessage, setProgressMessage] = useState('');
  const [results, setResults] = useState<{ blob: Blob; filename: string }[]>([]);
  const [error, setError] = useState<string | null>(null);
  
  // Split options
  const [splitMode, setSplitMode] = useState<SplitMode>('ranges');
  const [rangeInput, setRangeInput] = useState('');
  const [pagesPerSplit, setPagesPerSplit] = useState(1);
  
  // Page previews
  const [pagePreviews, setPagePreviews] = useState<PagePreview[]>([]);
  const [selectedPages, setSelectedPages] = useState<Set<number>>(new Set());
  const [isLoadingPreviews, setIsLoadingPreviews] = useState(false);
  
  // Ref for cancellation
  const cancelledRef = useRef(false);
  const pdfDocRef = useRef<any>(null);

  /**
   * Load PDF and generate page previews
   */
  const loadPdfPreviews = useCallback(async (pdfFile: File) => {
    setIsLoadingPreviews(true);
    setPagePreviews([]);
    
    try {
      const pdfjsLib = await import('pdfjs-dist');
      configurePdfjsWorker(pdfjsLib);
      
      const arrayBuffer = await pdfFile.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      
      pdfDocRef.current = pdf;
      setTotalPages(pdf.numPages);
      
      // Generate thumbnails for each page
      const previews: PagePreview[] = [];
      const maxPreviewPages = Math.min(pdf.numPages, 50); // Limit previews for performance
      
      for (let i = 1; i <= maxPreviewPages; i++) {
        const page = await pdf.getPage(i);
        const viewport = page.getViewport({ scale: 0.2 });
        
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        
        if (context) {
          canvas.height = viewport.height;
          canvas.width = viewport.width;
          
          await page.render({
            canvasContext: context,
            viewport: viewport,
          }).promise;
          
          previews.push({
            pageNumber: i,
            thumbnail: canvas.toDataURL('image/jpeg', 0.7),
          });
        }
      }
      
      // Add remaining pages without thumbnails
      for (let i = maxPreviewPages + 1; i <= pdf.numPages; i++) {
        previews.push({ pageNumber: i });
      }
      
      setPagePreviews(previews);
    } catch (err) {
      console.error('Failed to load PDF previews:', err);
      setError('Failed to load PDF preview. The file may be corrupted or encrypted.');
    } finally {
      setIsLoadingPreviews(false);
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
      setResults([]);
      setSelectedPages(new Set());
      setRangeInput('');
      loadPdfPreviews(selectedFile);
    }
  }, [loadPdfPreviews]);

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
    setPagePreviews([]);
    setSelectedPages(new Set());
    setResults([]);
    setError(null);
    setStatus('idle');
    setProgress(0);
    setRangeInput('');
    pdfDocRef.current = null;
  }, []);

  /**
   * Toggle page selection
   */
  const handleTogglePage = useCallback((pageNumber: number) => {
    setSelectedPages(prev => {
      const newSet = new Set(prev);
      if (newSet.has(pageNumber)) {
        newSet.delete(pageNumber);
      } else {
        newSet.add(pageNumber);
      }
      return newSet;
    });
  }, []);

  /**
   * Select all pages
   */
  const handleSelectAll = useCallback(() => {
    setSelectedPages(new Set(pagePreviews.map(p => p.pageNumber)));
  }, [pagePreviews]);

  /**
   * Deselect all pages
   */
  const handleDeselectAll = useCallback(() => {
    setSelectedPages(new Set());
  }, []);

  /**
   * Get page ranges from current selection/input
   */
  const getPageRanges = useCallback((): PageRange[] => {
    switch (splitMode) {
      case 'ranges':
        if (rangeInput.trim()) {
          return parsePageRanges(rangeInput, totalPages);
        }
        // If no input but pages selected, create ranges from selection
        if (selectedPages.size > 0) {
          const sortedPages = Array.from(selectedPages).sort((a, b) => a - b);
          const ranges: PageRange[] = [];
          let start = sortedPages[0];
          let end = sortedPages[0];
          
          for (let i = 1; i < sortedPages.length; i++) {
            if (sortedPages[i] === end + 1) {
              end = sortedPages[i];
            } else {
              ranges.push({ start, end });
              start = sortedPages[i];
              end = sortedPages[i];
            }
          }
          ranges.push({ start, end });
          return ranges;
        }
        return [];
        
      case 'every-page':
        return createSplitEveryPage(totalPages);
        
      case 'every-n-pages':
        return createSplitEveryNPages(totalPages, pagesPerSplit);
        
      default:
        return [];
    }
  }, [splitMode, rangeInput, selectedPages, totalPages, pagesPerSplit]);

  /**
   * Handle split operation
   */
  const handleSplit = useCallback(async () => {
    if (!file) {
      setError('Please upload a PDF file first.');
      return;
    }

    const ranges = getPageRanges();
    if (ranges.length === 0) {
      setError('Please specify page ranges or select pages to extract.');
      return;
    }

    cancelledRef.current = false;
    setStatus('processing');
    setProgress(0);
    setError(null);
    setResults([]);

    const options: SplitOptions = {
      ranges,
      outputFormat: 'multiple',
    };

    try {
      const output: ProcessOutput = await splitPDF(
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
        const blobs = Array.isArray(output.result) ? output.result : [output.result];
        const filenames = output.metadata?.outputFiles as string[] || 
          blobs.map((_, i) => `split_${i + 1}.pdf`);
        
        const resultFiles = blobs.map((blob, i) => ({
          blob,
          filename: filenames[i] || `split_${i + 1}.pdf`,
        }));
        
        setResults(resultFiles);
        setStatus('complete');
      } else {
        setError(output.error?.message || 'Failed to split PDF file.');
        setStatus('error');
      }
    } catch (err) {
      if (!cancelledRef.current) {
        setError(err instanceof Error ? err.message : 'An unexpected error occurred.');
        setStatus('error');
      }
    }
  }, [file, getPageRanges]);

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
  const canSplit = file && totalPages > 0 && !isProcessing && (
    (splitMode === 'ranges' && (rangeInput.trim() || selectedPages.size > 0)) ||
    splitMode === 'every-page' ||
    splitMode === 'every-n-pages'
  );

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
          label={tTools('splitPdf.uploadLabel') || 'Upload PDF File'}
          description={tTools('splitPdf.uploadDescription') || 'Drag and drop a PDF file here, or click to browse.'}
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

      {/* Split Mode Selection */}
      {file && totalPages > 0 && (
        <Card variant="outlined" size="lg">
          <h3 className="text-lg font-medium text-[hsl(var(--color-foreground))] mb-4">
            {tTools('splitPdf.splitModeTitle') || 'Split Method'}
          </h3>
          
          <div className="space-y-4">
            {/* Mode Selection */}
            <div 
              className="flex flex-wrap gap-2"
              role="radiogroup"
              aria-label="Split method selection"
            >
              <button
                type="button"
                role="radio"
                aria-checked={splitMode === 'ranges'}
                onClick={() => setSplitMode('ranges')}
                disabled={isProcessing}
                className={`px-4 py-2 rounded-[var(--radius-md)] text-sm font-medium transition-colors ${
                  splitMode === 'ranges'
                    ? 'bg-[hsl(var(--color-primary))] text-[hsl(var(--color-primary-foreground))]'
                    : 'bg-[hsl(var(--color-muted))] text-[hsl(var(--color-foreground))] hover:bg-[hsl(var(--color-muted)/0.8)]'
                }`}
              >
                {tTools('splitPdf.modeRanges') || 'By Page Ranges'}
              </button>
              <button
                type="button"
                role="radio"
                aria-checked={splitMode === 'every-page'}
                onClick={() => setSplitMode('every-page')}
                disabled={isProcessing}
                className={`px-4 py-2 rounded-[var(--radius-md)] text-sm font-medium transition-colors ${
                  splitMode === 'every-page'
                    ? 'bg-[hsl(var(--color-primary))] text-[hsl(var(--color-primary-foreground))]'
                    : 'bg-[hsl(var(--color-muted))] text-[hsl(var(--color-foreground))] hover:bg-[hsl(var(--color-muted)/0.8)]'
                }`}
              >
                {tTools('splitPdf.modeEveryPage') || 'Split Every Page'}
              </button>
              <button
                type="button"
                role="radio"
                aria-checked={splitMode === 'every-n-pages'}
                onClick={() => setSplitMode('every-n-pages')}
                disabled={isProcessing}
                className={`px-4 py-2 rounded-[var(--radius-md)] text-sm font-medium transition-colors ${
                  splitMode === 'every-n-pages'
                    ? 'bg-[hsl(var(--color-primary))] text-[hsl(var(--color-primary-foreground))]'
                    : 'bg-[hsl(var(--color-muted))] text-[hsl(var(--color-foreground))] hover:bg-[hsl(var(--color-muted)/0.8)]'
                }`}
              >
                {tTools('splitPdf.modeEveryN') || 'Split Every N Pages'}
              </button>
            </div>

            {/* Mode-specific options */}
            {splitMode === 'ranges' && (
              <div className="space-y-3">
                <div>
                  <label 
                    htmlFor="page-ranges" 
                    className="block text-sm font-medium text-[hsl(var(--color-foreground))] mb-1"
                  >
                    {tTools('splitPdf.rangeInputLabel') || 'Page Ranges'}
                  </label>
                  <input
                    id="page-ranges"
                    type="text"
                    value={rangeInput}
                    onChange={(e) => setRangeInput(e.target.value)}
                    placeholder="e.g., 1-5, 8, 10-15"
                    disabled={isProcessing}
                    className="w-full px-3 py-2 rounded-[var(--radius-md)] border border-[hsl(var(--color-border))] bg-[hsl(var(--color-background))] text-[hsl(var(--color-foreground))] placeholder:text-[hsl(var(--color-muted-foreground))] focus:outline-none focus:ring-2 focus:ring-[hsl(var(--color-primary))]"
                  />
                  <p className="mt-1 text-xs text-[hsl(var(--color-muted-foreground))]">
                    {tTools('splitPdf.rangeInputHint') || 'Enter page numbers or ranges separated by commas. Or click pages below to select.'}
                  </p>
                </div>
              </div>
            )}

            {splitMode === 'every-n-pages' && (
              <div>
                <label 
                  htmlFor="pages-per-split" 
                  className="block text-sm font-medium text-[hsl(var(--color-foreground))] mb-1"
                >
                  {tTools('splitPdf.pagesPerSplitLabel') || 'Pages per file'}
                </label>
                <input
                  id="pages-per-split"
                  type="number"
                  min={1}
                  max={totalPages}
                  value={pagesPerSplit}
                  onChange={(e) => setPagesPerSplit(Math.max(1, parseInt(e.target.value) || 1))}
                  disabled={isProcessing}
                  className="w-24 px-3 py-2 rounded-[var(--radius-md)] border border-[hsl(var(--color-border))] bg-[hsl(var(--color-background))] text-[hsl(var(--color-foreground))] focus:outline-none focus:ring-2 focus:ring-[hsl(var(--color-primary))]"
                />
                <p className="mt-1 text-xs text-[hsl(var(--color-muted-foreground))]">
                  {tTools('splitPdf.pagesPerSplitHint', { count: Math.ceil(totalPages / pagesPerSplit) }) || `Will create ${Math.ceil(totalPages / pagesPerSplit)} files`}
                </p>
              </div>
            )}
          </div>
        </Card>
      )}

      {/* Page Preview Grid */}
      {file && pagePreviews.length > 0 && splitMode === 'ranges' && (
        <Card variant="outlined" size="lg">
          <div className="flex items-center justify-between mb-4">
            <h3 className="text-lg font-medium text-[hsl(var(--color-foreground))]">
              {tTools('splitPdf.pagePreviewTitle') || 'Select Pages'} 
              {selectedPages.size > 0 && ` (${selectedPages.size} selected)`}
            </h3>
            <div className="flex gap-2">
              <Button variant="ghost" size="sm" onClick={handleSelectAll} disabled={isProcessing}>
                {t('buttons.selectAll') || 'Select All'}
              </Button>
              <Button variant="ghost" size="sm" onClick={handleDeselectAll} disabled={isProcessing}>
                {t('buttons.deselectAll') || 'Deselect All'}
              </Button>
            </div>
          </div>

          {isLoadingPreviews ? (
            <div className="flex items-center justify-center py-12">
              <div className="flex flex-col items-center gap-3">
                <div className="w-8 h-8 border-2 border-[hsl(var(--color-primary))] border-t-transparent rounded-full animate-spin" />
                <p className="text-sm text-[hsl(var(--color-muted-foreground))]">
                  {t('status.loading') || 'Loading previews...'}
                </p>
              </div>
            </div>
          ) : (
            <div className="grid grid-cols-4 sm:grid-cols-6 md:grid-cols-8 lg:grid-cols-10 gap-3 max-h-[400px] overflow-y-auto p-1">
              {pagePreviews.map((preview) => (
                <button
                  key={preview.pageNumber}
                  type="button"
                  onClick={() => handleTogglePage(preview.pageNumber)}
                  disabled={isProcessing}
                  className={`relative aspect-[3/4] rounded-[var(--radius-md)] border-2 overflow-hidden transition-all ${
                    selectedPages.has(preview.pageNumber)
                      ? 'border-[hsl(var(--color-primary))] ring-2 ring-[hsl(var(--color-primary)/0.3)]'
                      : 'border-[hsl(var(--color-border))] hover:border-[hsl(var(--color-primary)/0.5)]'
                  }`}
                  aria-label={`Page ${preview.pageNumber}${selectedPages.has(preview.pageNumber) ? ' (selected)' : ''}`}
                >
                  {preview.thumbnail ? (
                    <img
                      src={preview.thumbnail}
                      alt={`Page ${preview.pageNumber}`}
                      className="w-full h-full object-cover"
                    />
                  ) : (
                    <div className="w-full h-full bg-[hsl(var(--color-muted))] flex items-center justify-center">
                      <span className="text-xs text-[hsl(var(--color-muted-foreground))]">
                        {preview.pageNumber}
                      </span>
                    </div>
                  )}
                  <div className="absolute bottom-0 left-0 right-0 bg-black/60 text-white text-xs py-0.5 text-center">
                    {preview.pageNumber}
                  </div>
                  {selectedPages.has(preview.pageNumber) && (
                    <div className="absolute top-1 right-1 w-5 h-5 bg-[hsl(var(--color-primary))] rounded-full flex items-center justify-center">
                      <svg className="w-3 h-3 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={3} d="M5 13l4 4L19 7" />
                      </svg>
                    </div>
                  )}
                </button>
              ))}
            </div>
          )}
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
            onClick={handleSplit}
            disabled={!canSplit}
            loading={isProcessing}
          >
            {isProcessing 
              ? (t('status.processing') || 'Processing...') 
              : (tTools('splitPdf.splitButton') || 'Split PDF')
            }
          </Button>
        </div>
      )}

      {/* Results */}
      {status === 'complete' && results.length > 0 && (
        <Card variant="outlined" size="lg">
          <h3 className="text-lg font-medium text-[hsl(var(--color-foreground))] mb-4">
            {tTools('splitPdf.resultsTitle') || 'Split Results'} ({results.length} {results.length === 1 ? 'file' : 'files'})
          </h3>
          
          <div className="space-y-2">
            {results.map((result, index) => (
              <div 
                key={index}
                className="flex items-center justify-between p-3 rounded-[var(--radius-md)] bg-[hsl(var(--color-muted)/0.3)]"
              >
                <div className="flex items-center gap-3">
                  <svg className="w-8 h-8 text-red-500" viewBox="0 0 24 24" fill="currentColor">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z" />
                    <path d="M14 2v6h6" fill="white" />
                    <text x="7" y="17" fontSize="6" fill="white" fontWeight="bold">PDF</text>
                  </svg>
                  <div>
                    <p className="text-sm font-medium text-[hsl(var(--color-foreground))]">
                      {result.filename}
                    </p>
                    <p className="text-xs text-[hsl(var(--color-muted-foreground))]">
                      {formatSize(result.blob.size)}
                    </p>
                  </div>
                </div>
                <DownloadButton
                  file={result.blob}
                  filename={result.filename}
                  variant="secondary"
                  size="sm"
                />
              </div>
            ))}
          </div>
        </Card>
      )}

      {/* Success Message */}
      {status === 'complete' && results.length > 0 && (
        <div 
          className="p-4 rounded-[var(--radius-md)] bg-green-50 border border-green-200 text-green-700"
          role="status"
        >
          <p className="text-sm font-medium">
            {tTools('splitPdf.successMessage') || `PDF split successfully into ${results.length} file(s)! Click the download buttons to save your files.`}
          </p>
        </div>
      )}
    </div>
  );
}

export default SplitPDFTool;
