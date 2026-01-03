'use client';

import React, { useState, useCallback } from 'react';
import { useTranslations } from 'next-intl';
import { FileUploader } from '../FileUploader';
import { ProcessingProgress } from '../ProcessingProgress';
import { DownloadButton } from '../DownloadButton';
import { Button } from '@/components/ui/Button';
import { Card } from '@/components/ui/Card';
import { compressPDF, type CompressPDFOptions, type CompressionQuality } from '@/lib/pdf/processors/compress';
import { useBatchProcessing, type BatchFile } from '@/lib/hooks/useBatchProcessing';
import { Trash2, FileArchive, Check, AlertCircle, Loader2, X } from 'lucide-react';

export interface CompressPDFToolProps {
  /** Custom class name */
  className?: string;
}

/**
 * CompressPDFTool Component
 * Requirements: 5.1, 5.2, 10.1
 * 
 * Provides the UI for compressing PDF files with quality options.
 * Supports batch processing of multiple files with ZIP download.
 */
export function CompressPDFTool({ className = '' }: CompressPDFToolProps) {
  const t = useTranslations('common');
  const tTools = useTranslations('tools');

  // Options
  const [quality, setQuality] = useState<CompressionQuality>('medium');
  const [removeMetadata, setRemoveMetadata] = useState(false);
  const [optimizeImages, setOptimizeImages] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // Batch processing hook
  const {
    files,
    isProcessing,
    overallProgress,
    completedCount,
    errorCount,
    addFiles,
    removeFile,
    clearFiles,
    startProcessing,
    cancelProcessing,
    downloadAsZip,
  } = useBatchProcessing({
    maxConcurrent: 2,
  });

  /**
   * Handle files selected from uploader
   */
  const handleFilesSelected = useCallback((newFiles: File[]) => {
    if (newFiles.length > 0) {
      addFiles(newFiles);
      setError(null);
    }
  }, [addFiles]);

  /**
   * Handle file upload error
   */
  const handleUploadError = useCallback((errorMessage: string) => {
    setError(errorMessage);
  }, []);

  /**
   * Compress processor for batch processing
   */
  const compressProcessor = useCallback(async (
    file: File,
    onProgress: (progress: number) => void
  ): Promise<Blob> => {
    const options: CompressPDFOptions = {
      quality,
      removeMetadata,
      optimizeImages,
      removeUnusedObjects: true,
    };

    const output = await compressPDF(
      file,
      options,
      (prog) => onProgress(prog)
    );

    if (output.success && output.result) {
      return output.result as Blob;
    }

    throw new Error(output.error?.message || 'Failed to compress PDF file.');
  }, [quality, removeMetadata, optimizeImages]);

  /**
   * Handle compress operation
   */
  const handleCompress = useCallback(async () => {
    if (files.length === 0) {
      setError('Please select PDF files to compress.');
      return;
    }
    setError(null);
    await startProcessing(compressProcessor);
  }, [files.length, startProcessing, compressProcessor]);

  /**
   * Handle download as ZIP
   */
  const handleDownloadZip = useCallback(async () => {
    await downloadAsZip('compressed-pdfs.zip');
  }, [downloadAsZip]);

  /**
   * Format file size
   */
  const formatSize = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
  };

  /**
   * Get status icon for a file
   */
  const getStatusIcon = (status: BatchFile['status']) => {
    switch (status) {
      case 'pending':
        return <div className="w-4 h-4 rounded-full bg-gray-300" />;
      case 'processing':
        return <Loader2 className="w-4 h-4 animate-spin text-blue-500" />;
      case 'completed':
        return <Check className="w-4 h-4 text-green-500" />;
      case 'error':
        return <AlertCircle className="w-4 h-4 text-red-500" />;
    }
  };

  const hasFiles = files.length > 0;
  const canCompress = hasFiles && !isProcessing;
  const hasCompletedFiles = completedCount > 0;
  const allCompleted = hasFiles && completedCount === files.length;

  return (
    <div className={`space-y-6 ${className}`.trim()}>
      {/* File Upload Area */}
      <FileUploader
        accept={['application/pdf', '.pdf']}
        multiple={true}
        maxFiles={10}
        maxSize={500 * 1024 * 1024} // 500MB
        onFilesSelected={handleFilesSelected}
        onError={handleUploadError}
        disabled={isProcessing}
        label={tTools('compressPdf.uploadLabel') || 'Upload PDF Files'}
        description={tTools('compressPdf.batchUploadDescription') || 'Drag and drop PDF files here. You can compress up to 10 files at once.'}
      />

      {/* Error Message */}
      {error && (
        <div
          className="p-4 rounded-[var(--radius-md)] bg-red-50 border border-red-200 text-red-700"
          role="alert"
        >
          <p className="text-sm">{error}</p>
        </div>
      )}

      {/* File List */}
      {hasFiles && (
        <Card variant="outlined" size="lg">
          <div className="flex items-center justify-between mb-4">
            <h3 className="text-lg font-medium text-[hsl(var(--color-foreground))]">
              {tTools('compressPdf.filesTitle') || 'Files to Compress'} ({files.length})
            </h3>
            <Button
              variant="ghost"
              size="sm"
              onClick={clearFiles}
              disabled={isProcessing}
            >
              <Trash2 className="w-4 h-4 mr-1" />
              {t('buttons.clearAll') || 'Clear All'}
            </Button>
          </div>

          <div className="space-y-2 max-h-64 overflow-y-auto">
            {files.map((batchFile) => (
              <div
                key={batchFile.id}
                className="flex items-center justify-between p-3 bg-[hsl(var(--color-muted)/0.3)] rounded-[var(--radius-md)]"
              >
                <div className="flex items-center gap-3 flex-1 min-w-0">
                  {getStatusIcon(batchFile.status)}
                  <div className="flex-1 min-w-0">
                    <p className="text-sm font-medium text-[hsl(var(--color-foreground))] truncate">
                      {batchFile.file.name}
                    </p>
                    <div className="flex items-center gap-2">
                      <span className="text-xs text-[hsl(var(--color-muted-foreground))]">
                        {formatSize(batchFile.file.size)}
                      </span>
                      {batchFile.status === 'processing' && (
                        <span className="text-xs text-blue-500">
                          {batchFile.progress}%
                        </span>
                      )}
                      {batchFile.status === 'completed' && batchFile.result && (
                        <span className="text-xs text-green-500">
                          → {formatSize(batchFile.result.size)}
                        </span>
                      )}
                      {batchFile.status === 'error' && (
                        <span className="text-xs text-red-500">
                          {batchFile.error}
                        </span>
                      )}
                    </div>
                  </div>
                </div>

                {/* Individual download for completed files */}
                {batchFile.status === 'completed' && batchFile.result && (
                  <DownloadButton
                    file={batchFile.result}
                    filename={`${batchFile.file.name.replace('.pdf', '')}_compressed.pdf`}
                    variant="ghost"
                    size="sm"
                  />
                )}

                {/* Remove button for pending files */}
                {batchFile.status === 'pending' && !isProcessing && (
                  <button
                    onClick={() => removeFile(batchFile.id)}
                    className="p-1 text-[hsl(var(--color-muted-foreground))] hover:text-red-500 transition-colors"
                    aria-label="Remove file"
                  >
                    <X className="w-4 h-4" />
                  </button>
                )}
              </div>
            ))}
          </div>
        </Card>
      )}

      {/* Compression Options */}
      {hasFiles && (
        <Card variant="outlined">
          <h3 className="text-lg font-medium text-[hsl(var(--color-foreground))] mb-4">
            {tTools('compressPdf.optionsTitle') || 'Compression Options'}
          </h3>

          <div className="space-y-4">
            {/* Quality Selection */}
            <div>
              <label className="block text-sm font-medium text-[hsl(var(--color-foreground))] mb-2">
                {tTools('compressPdf.qualityLabel') || 'Compression Quality'}
              </label>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
                {(['low', 'medium', 'high', 'maximum'] as CompressionQuality[]).map((q) => (
                  <button
                    key={q}
                    type="button"
                    onClick={() => setQuality(q)}
                    disabled={isProcessing}
                    className={`
                      px-4 py-2 rounded-[var(--radius-md)] border text-sm font-medium
                      transition-colors duration-200
                      ${quality === q
                        ? 'border-[hsl(var(--color-primary))] bg-[hsl(var(--color-primary))] text-[hsl(var(--color-primary-foreground))]'
                        : 'border-[hsl(var(--color-border))] hover:bg-[hsl(var(--color-muted)/0.5)]'
                      }
                      disabled:opacity-50 disabled:cursor-not-allowed
                    `}
                  >
                    {q === 'low' && (tTools('compressPdf.qualityLow') || 'Low')}
                    {q === 'medium' && (tTools('compressPdf.qualityMedium') || 'Medium')}
                    {q === 'high' && (tTools('compressPdf.qualityHigh') || 'High')}
                    {q === 'maximum' && (tTools('compressPdf.qualityMaximum') || 'Maximum')}
                  </button>
                ))}
              </div>
              <p className="mt-1 text-xs text-[hsl(var(--color-muted-foreground))]">
                {quality === 'low' && (tTools('compressPdf.qualityLowDesc') || 'Smallest file size, lower quality')}
                {quality === 'medium' && (tTools('compressPdf.qualityMediumDesc') || 'Balanced compression and quality')}
                {quality === 'high' && (tTools('compressPdf.qualityHighDesc') || 'Better quality, moderate compression')}
                {quality === 'maximum' && (tTools('compressPdf.qualityMaximumDesc') || 'Best quality, minimal compression')}
              </p>
            </div>

            {/* Additional Options */}
            <div className="space-y-3">
              <label className="flex items-center gap-3 cursor-pointer">
                <input
                  type="checkbox"
                  checked={optimizeImages}
                  onChange={(e) => setOptimizeImages(e.target.checked)}
                  disabled={isProcessing}
                  className="w-4 h-4 rounded border-[hsl(var(--color-border))] text-[hsl(var(--color-primary))] focus:ring-[hsl(var(--color-primary))]"
                />
                <span className="text-sm text-[hsl(var(--color-foreground))]">
                  {tTools('compressPdf.optimizeImages') || 'Optimize embedded images'}
                </span>
              </label>

              <label className="flex items-center gap-3 cursor-pointer">
                <input
                  type="checkbox"
                  checked={removeMetadata}
                  onChange={(e) => setRemoveMetadata(e.target.checked)}
                  disabled={isProcessing}
                  className="w-4 h-4 rounded border-[hsl(var(--color-border))] text-[hsl(var(--color-primary))] focus:ring-[hsl(var(--color-primary))]"
                />
                <span className="text-sm text-[hsl(var(--color-foreground))]">
                  {tTools('compressPdf.removeMetadata') || 'Remove metadata (title, author, etc.)'}
                </span>
              </label>
            </div>
          </div>
        </Card>
      )}

      {/* Processing Progress */}
      {isProcessing && (
        <ProcessingProgress
          progress={overallProgress}
          status="processing"
          message={`Compressing ${completedCount + 1}/${files.length}...`}
          onCancel={cancelProcessing}
          showPercentage
        />
      )}

      {/* Action Buttons */}
      <div className="flex flex-wrap items-center gap-4">
        <Button
          variant="primary"
          size="lg"
          onClick={handleCompress}
          disabled={!canCompress}
          loading={isProcessing}
        >
          {isProcessing
            ? (t('status.processing') || 'Processing...')
            : (tTools('compressPdf.compressButton') || 'Compress PDFs')
          }
        </Button>

        {hasCompletedFiles && (
          <Button
            variant="secondary"
            size="lg"
            onClick={handleDownloadZip}
            disabled={isProcessing}
          >
            <FileArchive className="w-4 h-4 mr-2" />
            {tTools('compressPdf.downloadAllZip') || 'Download All as ZIP'}
          </Button>
        )}
      </div>

      {/* Batch Completion Status */}
      {allCompleted && (
        <div
          className="p-4 rounded-[var(--radius-md)] bg-green-50 border border-green-200 text-green-700"
          role="status"
        >
          <p className="text-sm font-medium">
            {tTools('compressPdf.batchSuccessMessage') || `Successfully compressed ${completedCount} PDF file(s)!`}
          </p>
          {errorCount > 0 && (
            <p className="text-sm text-red-600 mt-1">
              {errorCount} file(s) failed to compress.
            </p>
          )}
        </div>
      )}
    </div>
  );
}

export default CompressPDFTool;
