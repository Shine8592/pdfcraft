'use client';

import React, { useState, useCallback, useRef } from 'react';
import { useTranslations } from 'next-intl';
import { FileUploader } from '../FileUploader';
import { ProcessingProgress, ProcessingStatus } from '../ProcessingProgress';
import { DownloadButton } from '../DownloadButton';
import { Button } from '@/components/ui/Button';
import { Card } from '@/components/ui/Card';
import { Tabs } from '@/components/ui/Tabs';
import { configurePdfjsWorker } from '@/lib/pdf/loader';
import type { ProcessOutput } from '@/types/pdf';

export interface PDFMultiToolProps {
  className?: string;
}

type Operation = 'organize' | 'delete' | 'rotate' | 'extract' | 'add-blank' | 'duplicate';

interface PagePreview {
  pageNumber: number;
  thumbnail?: string;
  rotation: number;
  selected: boolean;
  markedForDelete: boolean;
}

/**
 * PDFMultiTool Component
 * Requirements: 5.1
 * 
 * All-in-one PDF editor combining multiple operations:
 * organize, delete, rotate, add-blank, extract, duplicate
 */
export function PDFMultiTool({ className = '' }: PDFMultiToolProps) {
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
  
  // Page management
  const [pagePreviews, setPagePreviews] = useState<PagePreview[]>([]);
  const [isLoadingPreviews, setIsLoadingPreviews] = useState(false);
  const [currentOperation, setCurrentOperation] = useState<Operation>('organize');

  // Drag state
  const [draggedIndex, setDraggedIndex] = useState<number | null>(null);
  const [dragOverIndex, setDragOverIndex] = useState<number | null>(null);
  
  // Rotation state
  const [globalRotation, setGlobalRotation] = useState<number>(0);
  
  // Add blank page state
  const [blankPagePosition, setBlankPagePosition] = useState<number>(1);
  const [blankPageCount, setBlankPageCount] = useState<number>(1);
  
  // Ref for cancellation
  const cancelledRef = useRef(false);

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
      
      setTotalPages(pdf.numPages);
      
      const previews: PagePreview[] = [];
      const maxPreviewPages = Math.min(pdf.numPages, 100);
      
      for (let i = 1; i <= maxPreviewPages; i++) {
        const page = await pdf.getPage(i);
        const viewport = page.getViewport({ scale: 0.15 });
        
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
            thumbnail: canvas.toDataURL('image/jpeg', 0.6),
            rotation: 0,
            selected: false,
            markedForDelete: false,
          });
        }
      }

      // Add remaining pages without thumbnails
      for (let i = maxPreviewPages + 1; i <= pdf.numPages; i++) {
        previews.push({
          pageNumber: i,
          rotation: 0,
          selected: false,
          markedForDelete: false,
        });
      }
      
      setPagePreviews(previews);
    } catch (err) {
      console.error('Failed to load PDF previews:', err);
      setError('Failed to load PDF preview. The file may be corrupted or encrypted.');
    } finally {
      setIsLoadingPreviews(false);
    }
  }, []);

  const handleFilesSelected = useCallback((files: File[]) => {
    if (files.length > 0) {
      const selectedFile = files[0];
      setFile(selectedFile);
      setError(null);
      setResult(null);
      loadPdfPreviews(selectedFile);
    }
  }, [loadPdfPreviews]);

  const handleUploadError = useCallback((errorMessage: string) => {
    setError(errorMessage);
  }, []);

  const handleClearFile = useCallback(() => {
    setFile(null);
    setTotalPages(0);
    setPagePreviews([]);
    setResult(null);
    setError(null);
    setStatus('idle');
    setProgress(0);
  }, []);

  // Drag handlers for organize
  const handleDragStart = useCallback((index: number) => {
    setDraggedIndex(index);
  }, []);

  const handleDragOver = useCallback((e: React.DragEvent, index: number) => {
    e.preventDefault();
    if (draggedIndex !== null && draggedIndex !== index) {
      setDragOverIndex(index);
    }
  }, [draggedIndex]);

  const handleDragEnd = useCallback(() => {
    if (draggedIndex !== null && dragOverIndex !== null && draggedIndex !== dragOverIndex) {
      setPagePreviews(prev => {
        const newPreviews = [...prev];
        const [draggedPage] = newPreviews.splice(draggedIndex, 1);
        newPreviews.splice(dragOverIndex, 0, draggedPage);
        return newPreviews;
      });
    }
    setDraggedIndex(null);
    setDragOverIndex(null);
  }, [draggedIndex, dragOverIndex]);

  const handleMovePage = useCallback((fromIndex: number, toIndex: number) => {
    if (toIndex < 0 || toIndex >= pagePreviews.length) return;
    setPagePreviews(prev => {
      const newPreviews = [...prev];
      const [movedPage] = newPreviews.splice(fromIndex, 1);
      newPreviews.splice(toIndex, 0, movedPage);
      return newPreviews;
    });
  }, [pagePreviews.length]);

  // Page selection handlers
  const handleToggleSelect = useCallback((index: number) => {
    setPagePreviews(prev => prev.map((p, i) => 
      i === index ? { ...p, selected: !p.selected } : p
    ));
  }, []);

  const handleSelectAll = useCallback(() => {
    setPagePreviews(prev => prev.map(p => ({ ...p, selected: true })));
  }, []);

  const handleDeselectAll = useCallback(() => {
    setPagePreviews(prev => prev.map(p => ({ ...p, selected: false })));
  }, []);

  // Delete handlers
  const handleToggleDelete = useCallback((index: number) => {
    setPagePreviews(prev => prev.map((p, i) => 
      i === index ? { ...p, markedForDelete: !p.markedForDelete } : p
    ));
  }, []);

  const handleDeleteSelected = useCallback(() => {
    setPagePreviews(prev => prev.map(p => 
      p.selected ? { ...p, markedForDelete: true, selected: false } : p
    ));
  }, []);

  // Rotation handlers
  const handleRotatePage = useCallback((index: number, degrees: number) => {
    setPagePreviews(prev => prev.map((p, i) => 
      i === index ? { ...p, rotation: (p.rotation + degrees) % 360 } : p
    ));
  }, []);

  const handleRotateSelected = useCallback((degrees: number) => {
    setPagePreviews(prev => prev.map(p => 
      p.selected ? { ...p, rotation: (p.rotation + degrees) % 360 } : p
    ));
  }, []);

  const handleRotateAll = useCallback((degrees: number) => {
    setGlobalRotation((prev) => (prev + degrees) % 360);
    setPagePreviews(prev => prev.map(p => ({ ...p, rotation: (p.rotation + degrees) % 360 })));
  }, []);

  // Duplicate handler
  const handleDuplicatePage = useCallback((index: number) => {
    setPagePreviews(prev => {
      const newPreviews = [...prev];
      const pageToDuplicate = { ...prev[index] };
      newPreviews.splice(index + 1, 0, pageToDuplicate);
      return newPreviews;
    });
  }, []);

  const handleDuplicateSelected = useCallback(() => {
    setPagePreviews(prev => {
      const newPreviews: PagePreview[] = [];
      prev.forEach(p => {
        newPreviews.push(p);
        if (p.selected) {
          newPreviews.push({ ...p, selected: false });
        }
      });
      return newPreviews;
    });
  }, []);

  // Add blank page handler
  const handleAddBlankPage = useCallback(() => {
    const insertIndex = Math.max(0, Math.min(blankPagePosition - 1, pagePreviews.length));
    const newPages: PagePreview[] = [];
    
    for (let i = 0; i < blankPageCount; i++) {
      newPages.push({
        pageNumber: -1, // -1 indicates blank page
        rotation: 0,
        selected: false,
        markedForDelete: false,
      });
    }
    
    setPagePreviews(prev => {
      const newPreviews = [...prev];
      newPreviews.splice(insertIndex, 0, ...newPages);
      return newPreviews;
    });
  }, [blankPagePosition, blankPageCount, pagePreviews.length]);

  // Reset handler
  const handleReset = useCallback(() => {
    if (file) {
      loadPdfPreviews(file);
    }
    setResult(null);
    setGlobalRotation(0);
  }, [file, loadPdfPreviews]);

  // Process handler
  const handleProcess = useCallback(async () => {
    if (!file || pagePreviews.length === 0) {
      setError('Please upload a PDF file first.');
      return;
    }

    // Filter out deleted pages
    const activePages = pagePreviews.filter(p => !p.markedForDelete);
    
    if (activePages.length === 0) {
      setError('Cannot create a PDF with no pages.');
      return;
    }

    cancelledRef.current = false;
    setStatus('processing');
    setProgress(0);
    setError(null);
    setResult(null);

    try {
      setProgress(5);
      setProgressMessage('Loading PDF library...');

      const pdfLib = await import('pdf-lib');
      
      if (cancelledRef.current) {
        setStatus('idle');
        return;
      }

      setProgress(10);
      setProgressMessage('Loading source PDF...');

      const arrayBuffer = await file.arrayBuffer();
      const sourcePdf = await pdfLib.PDFDocument.load(arrayBuffer, {
        ignoreEncryption: false,
      });

      setProgress(20);
      setProgressMessage('Creating new document...');

      const newPdf = await pdfLib.PDFDocument.create();
      const progressPerPage = 70 / activePages.length;

      for (let i = 0; i < activePages.length; i++) {
        if (cancelledRef.current) {
          setStatus('idle');
          return;
        }

        const pageInfo = activePages[i];
        setProgress(20 + (i * progressPerPage));
        setProgressMessage(`Processing page ${i + 1} of ${activePages.length}...`);

        if (pageInfo.pageNumber === -1) {
          // Add blank page
          const firstPage = sourcePdf.getPage(0);
          const { width, height } = firstPage.getSize();
          newPdf.addPage([width, height]);
        } else {
          // Copy existing page
          const pageIndex = pageInfo.pageNumber - 1;
          const [copiedPage] = await newPdf.copyPages(sourcePdf, [pageIndex]);
          
          // Apply rotation if needed
          if (pageInfo.rotation !== 0) {
            copiedPage.setRotation(pdfLib.degrees(pageInfo.rotation));
          }
          
          newPdf.addPage(copiedPage);
        }
      }

      setProgress(90);
      setProgressMessage('Saving PDF...');

      const pdfBytes = await newPdf.save();
      const blob = new Blob([new Uint8Array(pdfBytes)], { type: 'application/pdf' });

      setProgress(100);
      setProgressMessage('Complete!');
      setResult(blob);
      setStatus('complete');

    } catch (err) {
      if (!cancelledRef.current) {
        setError(err instanceof Error ? err.message : 'An unexpected error occurred.');
        setStatus('error');
      }
    }
  }, [file, pagePreviews]);

  const handleCancel = useCallback(() => {
    cancelledRef.current = true;
    setStatus('idle');
    setProgress(0);
  }, []);

  const formatSize = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  const isProcessing = status === 'processing' || status === 'uploading';
  const canProcess = file && pagePreviews.length > 0 && !isProcessing;
  const selectedCount = pagePreviews.filter(p => p.selected).length;
  const deletedCount = pagePreviews.filter(p => p.markedForDelete).length;

  // Operation tabs
  const operationTabs = [
    { id: 'organize', label: tTools('pdfMultiTool.organize') || 'Organize', content: null },
    { id: 'delete', label: tTools('pdfMultiTool.delete') || 'Delete', content: null },
    { id: 'rotate', label: tTools('pdfMultiTool.rotate') || 'Rotate', content: null },
    { id: 'duplicate', label: tTools('pdfMultiTool.duplicate') || 'Duplicate', content: null },
    { id: 'add-blank', label: tTools('pdfMultiTool.addBlank') || 'Add Blank', content: null },
  ];

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
          label={tTools('pdfMultiTool.uploadLabel') || 'Upload PDF File'}
          description={tTools('pdfMultiTool.uploadDescription') || 'Drag and drop a PDF file here, or click to browse.'}
        />
      )}

      {/* Error Message */}
      {error && (
        <div className="p-4 rounded-[var(--radius-md)] bg-red-50 border border-red-200 text-red-700" role="alert">
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
                  {deletedCount > 0 && ` • ${deletedCount} marked for deletion`}
                </p>
              </div>
            </div>
            <Button variant="ghost" size="sm" onClick={handleClearFile} disabled={isProcessing}>
              {t('buttons.remove') || 'Remove'}
            </Button>
          </div>
        </Card>
      )}

      {/* Operation Tabs and Page Grid */}
      {file && pagePreviews.length > 0 && (
        <Card variant="outlined" size="lg">
          {/* Operation Selector */}
          <div className="mb-4 border-b border-[hsl(var(--color-border))] pb-4">
            <div className="flex flex-wrap gap-2">
              {operationTabs.map(tab => (
                <button
                  key={tab.id}
                  onClick={() => setCurrentOperation(tab.id as Operation)}
                  className={`px-4 py-2 rounded-[var(--radius-md)] text-sm font-medium transition-colors
                    ${currentOperation === tab.id 
                      ? 'bg-[hsl(var(--color-primary))] text-white' 
                      : 'bg-[hsl(var(--color-muted))] text-[hsl(var(--color-muted-foreground))] hover:bg-[hsl(var(--color-muted)/0.8)]'
                    }`}
                  disabled={isProcessing}
                >
                  {tab.label}
                </button>
              ))}
            </div>
          </div>

          {/* Operation-specific controls */}
          <div className="mb-4">
            {currentOperation === 'organize' && (
              <p className="text-sm text-[hsl(var(--color-muted-foreground))]">
                {tTools('pdfMultiTool.organizeHint') || 'Drag and drop pages to reorder them.'}
              </p>
            )}
            
            {currentOperation === 'delete' && (
              <div className="flex flex-wrap items-center gap-2">
                <p className="text-sm text-[hsl(var(--color-muted-foreground))]">
                  {tTools('pdfMultiTool.deleteHint') || 'Click pages to mark them for deletion.'}
                </p>
                {selectedCount > 0 && (
                  <Button variant="outline" size="sm" onClick={handleDeleteSelected}>
                    Delete Selected ({selectedCount})
                  </Button>
                )}
              </div>
            )}
            
            {currentOperation === 'rotate' && (
              <div className="flex flex-wrap items-center gap-2">
                <p className="text-sm text-[hsl(var(--color-muted-foreground))] mr-2">
                  {tTools('pdfMultiTool.rotateHint') || 'Click rotation buttons on pages, or rotate all:'}
                </p>
                <Button variant="outline" size="sm" onClick={() => handleRotateAll(90)}>
                  Rotate All 90°
                </Button>
                <Button variant="outline" size="sm" onClick={() => handleRotateAll(180)}>
                  Rotate All 180°
                </Button>
                {selectedCount > 0 && (
                  <Button variant="outline" size="sm" onClick={() => handleRotateSelected(90)}>
                    Rotate Selected 90°
                  </Button>
                )}
              </div>
            )}

            {currentOperation === 'duplicate' && (
              <div className="flex flex-wrap items-center gap-2">
                <p className="text-sm text-[hsl(var(--color-muted-foreground))]">
                  {tTools('pdfMultiTool.duplicateHint') || 'Click the duplicate button on pages to copy them.'}
                </p>
                {selectedCount > 0 && (
                  <Button variant="outline" size="sm" onClick={handleDuplicateSelected}>
                    Duplicate Selected ({selectedCount})
                  </Button>
                )}
              </div>
            )}
            
            {currentOperation === 'add-blank' && (
              <div className="flex flex-wrap items-center gap-3">
                <label className="text-sm text-[hsl(var(--color-muted-foreground))]">
                  Insert
                  <input
                    type="number"
                    min="1"
                    max="10"
                    value={blankPageCount}
                    onChange={(e) => setBlankPageCount(Math.max(1, Math.min(10, parseInt(e.target.value) || 1)))}
                    className="mx-2 w-16 px-2 py-1 border rounded text-center"
                    disabled={isProcessing}
                  />
                  blank page(s) at position
                  <input
                    type="number"
                    min="1"
                    max={pagePreviews.length + 1}
                    value={blankPagePosition}
                    onChange={(e) => setBlankPagePosition(Math.max(1, parseInt(e.target.value) || 1))}
                    className="mx-2 w-16 px-2 py-1 border rounded text-center"
                    disabled={isProcessing}
                  />
                </label>
                <Button variant="outline" size="sm" onClick={handleAddBlankPage} disabled={isProcessing}>
                  Add Blank Page(s)
                </Button>
              </div>
            )}
          </div>

          {/* Selection controls */}
          {(currentOperation === 'delete' || currentOperation === 'rotate' || currentOperation === 'duplicate') && (
            <div className="flex gap-2 mb-4">
              <Button variant="ghost" size="sm" onClick={handleSelectAll} disabled={isProcessing}>
                Select All
              </Button>
              <Button variant="ghost" size="sm" onClick={handleDeselectAll} disabled={isProcessing || selectedCount === 0}>
                Deselect All
              </Button>
            </div>
          )}

          {/* Page Grid */}
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
            <div className="grid grid-cols-3 sm:grid-cols-4 md:grid-cols-5 lg:grid-cols-6 xl:grid-cols-8 gap-3 max-h-[500px] overflow-y-auto p-1">
              {pagePreviews.map((page, index) => (
                <div
                  key={`${page.pageNumber}-${index}`}
                  draggable={currentOperation === 'organize' && !isProcessing}
                  onDragStart={() => currentOperation === 'organize' && handleDragStart(index)}
                  onDragOver={(e) => currentOperation === 'organize' && handleDragOver(e, index)}
                  onDragEnd={handleDragEnd}
                  onClick={() => {
                    if (currentOperation === 'delete') handleToggleDelete(index);
                    if (['rotate', 'duplicate', 'delete'].includes(currentOperation)) handleToggleSelect(index);
                  }}
                  className={`
                    relative aspect-[3/4] rounded-[var(--radius-md)] border-2 overflow-hidden transition-all
                    ${draggedIndex === index ? 'opacity-50 border-dashed scale-95' : ''}
                    ${dragOverIndex === index ? 'border-[hsl(var(--color-primary))] ring-2 ring-[hsl(var(--color-primary)/0.3)]' : 'border-[hsl(var(--color-border))]'}
                    ${page.selected ? 'ring-2 ring-blue-500 border-blue-500' : ''}
                    ${page.markedForDelete ? 'opacity-40 border-red-500' : ''}
                    ${currentOperation === 'organize' && !isProcessing ? 'cursor-grab hover:border-[hsl(var(--color-primary)/0.5)]' : ''}
                    ${['delete', 'rotate', 'duplicate'].includes(currentOperation) ? 'cursor-pointer' : ''}
                  `}
                >
                  {page.thumbnail ? (
                    <img
                      src={page.thumbnail}
                      alt={`Page ${page.pageNumber}`}
                      className="w-full h-full object-cover"
                      style={{ transform: `rotate(${page.rotation}deg)` }}
                      draggable={false}
                    />
                  ) : (
                    <div className="w-full h-full bg-[hsl(var(--color-muted))] flex items-center justify-center">
                      <span className="text-lg font-medium text-[hsl(var(--color-muted-foreground))]">
                        {page.pageNumber === -1 ? 'Blank' : page.pageNumber}
                      </span>
                    </div>
                  )}

                  {/* Page number badge */}
                  <div className="absolute bottom-0 left-0 right-0 bg-black/70 text-white text-xs py-1 text-center font-medium">
                    {page.pageNumber === -1 ? 'Blank' : page.pageNumber}
                    {page.rotation !== 0 && ` (${page.rotation}°)`}
                  </div>
                  
                  {/* Position indicator */}
                  <div className="absolute top-1 left-1 w-5 h-5 bg-[hsl(var(--color-primary))] rounded-full flex items-center justify-center text-[10px] font-bold text-white">
                    {index + 1}
                  </div>

                  {/* Delete marker */}
                  {page.markedForDelete && (
                    <div className="absolute inset-0 flex items-center justify-center bg-red-500/30">
                      <svg className="w-8 h-8 text-red-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3">
                        <path d="M18 6L6 18M6 6l12 12" />
                      </svg>
                    </div>
                  )}

                  {/* Action buttons based on operation */}
                  {currentOperation === 'organize' && (
                    <div className="absolute top-1 right-1 flex flex-col gap-0.5">
                      <button
                        type="button"
                        onClick={(e) => { e.stopPropagation(); handleMovePage(index, index - 1); }}
                        disabled={index === 0 || isProcessing}
                        className="w-5 h-5 bg-white/90 rounded flex items-center justify-center hover:bg-white disabled:opacity-30"
                        aria-label="Move up"
                      >
                        <svg className="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3">
                          <path d="M18 15l-6-6-6 6" />
                        </svg>
                      </button>
                      <button
                        type="button"
                        onClick={(e) => { e.stopPropagation(); handleMovePage(index, index + 1); }}
                        disabled={index === pagePreviews.length - 1 || isProcessing}
                        className="w-5 h-5 bg-white/90 rounded flex items-center justify-center hover:bg-white disabled:opacity-30"
                        aria-label="Move down"
                      >
                        <svg className="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3">
                          <path d="M6 9l6 6 6-6" />
                        </svg>
                      </button>
                    </div>
                  )}

                  {currentOperation === 'rotate' && (
                    <div className="absolute top-1 right-1 flex flex-col gap-0.5">
                      <button
                        type="button"
                        onClick={(e) => { e.stopPropagation(); handleRotatePage(index, -90); }}
                        disabled={isProcessing}
                        className="w-5 h-5 bg-white/90 rounded flex items-center justify-center hover:bg-white"
                        aria-label="Rotate left"
                      >
                        <svg className="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
                          <path d="M3 3v5h5" />
                        </svg>
                      </button>
                      <button
                        type="button"
                        onClick={(e) => { e.stopPropagation(); handleRotatePage(index, 90); }}
                        disabled={isProcessing}
                        className="w-5 h-5 bg-white/90 rounded flex items-center justify-center hover:bg-white"
                        aria-label="Rotate right"
                      >
                        <svg className="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M21 12a9 9 0 1 1-9-9 9.75 9.75 0 0 1 6.74 2.74L21 8" />
                          <path d="M21 3v5h-5" />
                        </svg>
                      </button>
                    </div>
                  )}

                  {currentOperation === 'duplicate' && (
                    <div className="absolute top-1 right-1">
                      <button
                        type="button"
                        onClick={(e) => { e.stopPropagation(); handleDuplicatePage(index); }}
                        disabled={isProcessing}
                        className="w-5 h-5 bg-white/90 rounded flex items-center justify-center hover:bg-white"
                        aria-label="Duplicate page"
                      >
                        <svg className="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <rect x="9" y="9" width="13" height="13" rx="2" ry="2" />
                          <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
                        </svg>
                      </button>
                    </div>
                  )}
                </div>
              ))}
            </div>
          )}

          {/* Reset button */}
          <div className="mt-4 pt-4 border-t border-[hsl(var(--color-border))]">
            <Button variant="ghost" size="sm" onClick={handleReset} disabled={isProcessing}>
              {tTools('pdfMultiTool.reset') || 'Reset to Original'}
            </Button>
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
            onClick={handleProcess}
            disabled={!canProcess}
            loading={isProcessing}
          >
            {isProcessing 
              ? (t('status.processing') || 'Processing...') 
              : (tTools('pdfMultiTool.processButton') || 'Apply Changes')
            }
          </Button>

          {result && (
            <DownloadButton
              file={result}
              filename={file.name.replace('.pdf', '_edited.pdf')}
              variant="secondary"
              size="lg"
              showFileSize
            />
          )}
        </div>
      )}

      {/* Success Message */}
      {status === 'complete' && result && (
        <div className="p-4 rounded-[var(--radius-md)] bg-green-50 border border-green-200 text-green-700" role="status">
          <p className="text-sm font-medium">
            {tTools('pdfMultiTool.successMessage') || 'PDF processed successfully! Click the download button to save your file.'}
          </p>
        </div>
      )}
    </div>
  );
}

export default PDFMultiTool;
