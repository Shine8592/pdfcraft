'use client';

import React, { useState, useCallback, useRef } from 'react';
import { FileUploader } from '../FileUploader';
import { ProcessingProgress, ProcessingStatus } from '../ProcessingProgress';
import { DownloadButton } from '../DownloadButton';
import { Button } from '@/components/ui/Button';
import { Card } from '@/components/ui/Card';
import { removeBlankPages } from '@/lib/pdf/processors/remove-blank-pages';
import type { ProcessOutput } from '@/types/pdf';

export interface RemoveBlankPagesToolProps { className?: string; }

export function RemoveBlankPagesTool({ className = '' }: RemoveBlankPagesToolProps) {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<ProcessingStatus>('idle');
  const [progress, setProgress] = useState(0);
  const [result, setResult] = useState<Blob | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [threshold, setThreshold] = useState(99);
  const [removedCount, setRemovedCount] = useState(0);
  const cancelledRef = useRef(false);

  const handleProcess = useCallback(async () => {
    if (!file) return;
    cancelledRef.current = false;
    setStatus('processing'); setProgress(0); setError(null); setResult(null); setRemovedCount(0);
    try {
      const output: ProcessOutput = await removeBlankPages(file, { threshold }, (prog) => { if (!cancelledRef.current) setProgress(prog); });
      if (output.success && output.result) {
        setResult(output.result as Blob);
        setRemovedCount(output.metadata?.blankPagesRemoved as number || 0);
        setStatus('complete');
      } else { setError(output.error?.message || 'Failed.'); setStatus('error'); }
    } catch (err) { setError(err instanceof Error ? err.message : 'Error'); setStatus('error'); }
  }, [file, threshold]);

  const isProcessing = status === 'processing';

  return (
    <div className={`space-y-6 ${className}`.trim()}>
      {!file && <FileUploader accept={['application/pdf', '.pdf']} multiple={false} maxFiles={1} maxSize={100 * 1024 * 1024} onFilesSelected={(files) => { if (files.length > 0) { setFile(files[0]); setError(null); setResult(null); } }} onError={setError} disabled={isProcessing} label="Upload PDF File" description="Drag and drop a PDF file here." />}
      {error && <div className="p-4 rounded bg-red-50 border border-red-200 text-red-700"><p className="text-sm">{error}</p></div>}
      {file && (
        <>
          <Card variant="outlined"><div className="flex items-center justify-between"><p className="font-medium">{file.name}</p><Button variant="ghost" size="sm" onClick={() => { setFile(null); setResult(null); }} disabled={isProcessing}>Remove</Button></div></Card>
          <Card variant="outlined" size="lg">
            <h3 className="text-lg font-medium mb-4">Options</h3>
            <div>
              <label className="block text-sm font-medium mb-1">Blank Detection Threshold ({threshold}%)</label>
              <input type="range" value={threshold} onChange={(e) => setThreshold(parseInt(e.target.value))} min={90} max={100} className="w-full" disabled={isProcessing} />
              <p className="text-sm text-gray-500 mt-1">Higher values detect only nearly-white pages. Lower values may detect pages with minimal content.</p>
            </div>
          </Card>
        </>
      )}
      {isProcessing && <ProcessingProgress progress={progress} status={status} onCancel={() => { cancelledRef.current = true; setStatus('idle'); }} showPercentage />}
      {file && <div className="flex flex-wrap items-center gap-4"><Button variant="primary" size="lg" onClick={handleProcess} disabled={!file || isProcessing} loading={isProcessing}>{isProcessing ? 'Processing...' : 'Remove Blank Pages'}</Button>{result && <DownloadButton file={result} filename={file.name.replace('.pdf', '_no_blanks.pdf')} variant="secondary" size="lg" showFileSize />}</div>}
      {status === 'complete' && result && <div className="p-4 rounded bg-green-50 border border-green-200 text-green-700"><p className="text-sm font-medium">{removedCount > 0 ? `${removedCount} blank page(s) removed!` : 'No blank pages found.'}</p></div>}
    </div>
  );
}

export default RemoveBlankPagesTool;
