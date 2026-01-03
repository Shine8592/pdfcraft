'use client';

import React, { useState, useCallback, useRef } from 'react';
import { useTranslations } from 'next-intl';
import { FileUploader } from '../FileUploader';
import { ProcessingProgress, ProcessingStatus } from '../ProcessingProgress';
import { DownloadButton } from '../DownloadButton';
import { Button } from '@/components/ui/Button';
import { Card } from '@/components/ui/Card';
import { addWatermark, WatermarkOptions } from '@/lib/pdf/processors/watermark';
import type { ProcessOutput } from '@/types/pdf';

export interface WatermarkToolProps {
  className?: string;
}

type WatermarkType = 'text' | 'image';

export function WatermarkTool({ className = '' }: WatermarkToolProps) {
  const t = useTranslations('common');
  const tTools = useTranslations('tools.watermark');

  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<ProcessingStatus>('idle');
  const [progress, setProgress] = useState(0);
  const [progressMessage, setProgressMessage] = useState('');
  const [result, setResult] = useState<Blob | null>(null);
  const [error, setError] = useState<string | null>(null);
  
  // Watermark type
  const [watermarkType, setWatermarkType] = useState<WatermarkType>('text');
  
  // Text watermark options
  const [watermarkText, setWatermarkText] = useState('CONFIDENTIAL');
  const [fontSize, setFontSize] = useState(72);
  const [textColor, setTextColor] = useState('#888888');
  const [textOpacity, setTextOpacity] = useState(0.3);
  const [textAngle, setTextAngle] = useState(-45);
  
  // Image watermark options
  const [imageFile, setImageFile] = useState<File | null>(null);
  const [imageOpacity, setImageOpacity] = useState(0.3);
  const [imageAngle, setImageAngle] = useState(0);

  const cancelledRef = useRef(false);

  const handleFilesSelected = useCallback((files: File[]) => {
    if (files.length > 0) {
      setFile(files[0]);
      setError(null);
      setResult(null);
    }
  }, []);

  const handleImageSelected = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile) {
      if (selectedFile.type === 'image/png' || selectedFile.type === 'image/jpeg') {
        setImageFile(selectedFile);
        setError(null);
      } else {
        setError(tTools('unsupportedImage'));
      }
    }
  }, [tTools]);

  const handleClearFile = useCallback(() => {
    setFile(null);
    setResult(null);
    setError(null);
    setStatus('idle');
  }, []);

  const handleProcess = useCallback(async () => {
    if (!file) return;
    if (watermarkType === 'text' && !watermarkText.trim()) {
      setError(tTools('enterText'));
      return;
    }
    if (watermarkType === 'image' && !imageFile) {
      setError(tTools('selectImage'));
      return;
    }

    cancelledRef.current = false;
    setStatus('processing');
    setProgress(0);
    setError(null);
    setResult(null);

    try {
      // Parse hex color to RGB
      const hexToRgb = (hex: string) => {
        const r = parseInt(hex.slice(1, 3), 16) / 255;
        const g = parseInt(hex.slice(3, 5), 16) / 255;
        const b = parseInt(hex.slice(5, 7), 16) / 255;
        return { r, g, b };
      };

      let options: WatermarkOptions;
      
      if (watermarkType === 'text') {
        options = {
          type: 'text',
          text: watermarkText,
          fontSize,
          color: hexToRgb(textColor),
          opacity: textOpacity,
          rotation: textAngle,
          pages: 'all',
        };
      } else {
        const imageData = await imageFile!.arrayBuffer();
        options = {
          type: 'image',
          imageData,
          imageType: imageFile!.type === 'image/png' ? 'png' : 'jpg',
          opacity: imageOpacity,
          rotation: imageAngle,
          pages: 'all',
        };
      }

      const output: ProcessOutput = await addWatermark(file, options, (prog, message) => {
        if (!cancelledRef.current) {
          setProgress(prog);
          setProgressMessage(message || '');
        }
      });

      if (cancelledRef.current) {
        setStatus('idle');
        return;
      }

      if (output.success && output.result) {
        setResult(output.result as Blob);
        setStatus('complete');
      } else {
        setError(output.error?.message || tTools('failed'));
        setStatus('error');
      }
    } catch (err) {
      if (!cancelledRef.current) {
        setError(err instanceof Error ? err.message : tTools('failed'));
        setStatus('error');
      }
    }
  }, [file, watermarkType, watermarkText, fontSize, textColor, textOpacity, textAngle, imageFile, imageOpacity, imageAngle, tTools]);

  const formatSize = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  const isProcessing = status === 'processing';

  return (
    <div className={`space-y-6 ${className}`.trim()}>
      {!file && (
        <FileUploader
          accept={['application/pdf', '.pdf']}
          multiple={false}
          maxFiles={1}
          maxSize={100 * 1024 * 1024}
          onFilesSelected={handleFilesSelected}
          onError={setError}
          disabled={isProcessing}
          label={tTools('uploadLabel')}
          description={tTools('uploadDescription')}
        />
      )}

      {error && (
        <div className="p-4 rounded-lg bg-red-50 border border-red-200 text-red-700 dark:bg-red-900/20 dark:border-red-800 dark:text-red-400">
          <p className="text-sm">{error}</p>
        </div>
      )}

      {file && (
        <>
          <Card variant="outlined">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-3">
                <svg className="w-10 h-10 text-red-500" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z" />
                </svg>
                <div>
                  <p className="font-medium text-gray-900 dark:text-gray-100">{file.name}</p>
                  <p className="text-sm text-gray-500 dark:text-gray-400">{formatSize(file.size)}</p>
                </div>
              </div>
              <Button variant="ghost" size="sm" onClick={handleClearFile} disabled={isProcessing}>
                {t('buttons.remove')}
              </Button>
            </div>
          </Card>

          <Card variant="outlined" size="lg">
            <h3 className="text-lg font-medium mb-4 text-gray-900 dark:text-gray-100">
              {tTools('optionsTitle')}
            </h3>
            
            {/* Watermark Type Selection */}
            <div className="flex gap-6 mb-6">
              <label className="flex items-center gap-2 cursor-pointer">
                <input
                  type="radio"
                  name="watermark-type"
                  value="text"
                  checked={watermarkType === 'text'}
                  onChange={() => setWatermarkType('text')}
                  className="w-4 h-4 text-blue-600"
                  disabled={isProcessing}
                />
                <span className="text-sm font-medium text-gray-700 dark:text-gray-300">
                  {tTools('textWatermark')}
                </span>
              </label>
              <label className="flex items-center gap-2 cursor-pointer">
                <input
                  type="radio"
                  name="watermark-type"
                  value="image"
                  checked={watermarkType === 'image'}
                  onChange={() => setWatermarkType('image')}
                  className="w-4 h-4 text-blue-600"
                  disabled={isProcessing}
                />
                <span className="text-sm font-medium text-gray-700 dark:text-gray-300">
                  {tTools('imageWatermark')}
                </span>
              </label>
            </div>

            {/* Text Watermark Options */}
            {watermarkType === 'text' && (
              <div className="space-y-4">
                <div>
                  <label className="block text-sm font-medium mb-1 text-gray-700 dark:text-gray-300">
                    {tTools('watermarkText')}
                  </label>
                  <input
                    type="text"
                    value={watermarkText}
                    onChange={(e) => setWatermarkText(e.target.value)}
                    placeholder="CONFIDENTIAL"
                    className="w-full px-3 py-2 border rounded-lg bg-white dark:bg-gray-800 border-gray-300 dark:border-gray-600 text-gray-900 dark:text-gray-100"
                    disabled={isProcessing}
                  />
                </div>
                
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium mb-1 text-gray-700 dark:text-gray-300">
                      {tTools('fontSize')}
                    </label>
                    <input
                      type="number"
                      value={fontSize}
                      onChange={(e) => setFontSize(parseInt(e.target.value) || 72)}
                      min={10}
                      max={200}
                      className="w-full px-3 py-2 border rounded-lg bg-white dark:bg-gray-800 border-gray-300 dark:border-gray-600 text-gray-900 dark:text-gray-100"
                      disabled={isProcessing}
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium mb-1 text-gray-700 dark:text-gray-300">
                      {tTools('color')}
                    </label>
                    <div className="flex items-center gap-2">
                      <input
                        type="color"
                        value={textColor}
                        onChange={(e) => setTextColor(e.target.value)}
                        className="w-10 h-10 p-1 cursor-pointer rounded border border-gray-300 dark:border-gray-600"
                        disabled={isProcessing}
                      />
                      <input
                        type="text"
                        value={textColor}
                        onChange={(e) => setTextColor(e.target.value)}
                        className="flex-1 px-3 py-2 border rounded-lg bg-white dark:bg-gray-800 border-gray-300 dark:border-gray-600 text-gray-900 dark:text-gray-100 text-sm"
                        disabled={isProcessing}
                      />
                    </div>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium mb-1 text-gray-700 dark:text-gray-300">
                      {tTools('opacity')}: {Math.round(textOpacity * 100)}%
                    </label>
                    <input
                      type="range"
                      value={textOpacity}
                      onChange={(e) => setTextOpacity(parseFloat(e.target.value))}
                      min={0.1}
                      max={1}
                      step={0.1}
                      className="w-full"
                      disabled={isProcessing}
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium mb-1 text-gray-700 dark:text-gray-300">
                      {tTools('angle')}: {textAngle}°
                    </label>
                    <input
                      type="range"
                      value={textAngle}
                      onChange={(e) => setTextAngle(parseInt(e.target.value))}
                      min={-90}
                      max={90}
                      step={5}
                      className="w-full"
                      disabled={isProcessing}
                    />
                  </div>
                </div>
              </div>
            )}

            {/* Image Watermark Options */}
            {watermarkType === 'image' && (
              <div className="space-y-4">
                <div>
                  <label className="block text-sm font-medium mb-1 text-gray-700 dark:text-gray-300">
                    {tTools('watermarkImage')}
                  </label>
                  <input
                    type="file"
                    accept="image/png, image/jpeg"
                    onChange={handleImageSelected}
                    className="w-full px-3 py-2 border rounded-lg bg-white dark:bg-gray-800 border-gray-300 dark:border-gray-600 text-gray-900 dark:text-gray-100 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-semibold file:bg-blue-600 file:text-white hover:file:bg-blue-700"
                    disabled={isProcessing}
                  />
                  {imageFile && (
                    <p className="text-sm text-gray-500 dark:text-gray-400 mt-1">
                      {imageFile.name} ({formatSize(imageFile.size)})
                    </p>
                  )}
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium mb-1 text-gray-700 dark:text-gray-300">
                      {tTools('opacity')}: {Math.round(imageOpacity * 100)}%
                    </label>
                    <input
                      type="range"
                      value={imageOpacity}
                      onChange={(e) => setImageOpacity(parseFloat(e.target.value))}
                      min={0.1}
                      max={1}
                      step={0.1}
                      className="w-full"
                      disabled={isProcessing}
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium mb-1 text-gray-700 dark:text-gray-300">
                      {tTools('angle')}: {imageAngle}°
                    </label>
                    <input
                      type="range"
                      value={imageAngle}
                      onChange={(e) => setImageAngle(parseInt(e.target.value))}
                      min={-90}
                      max={90}
                      step={5}
                      className="w-full"
                      disabled={isProcessing}
                    />
                  </div>
                </div>
              </div>
            )}
          </Card>
        </>
      )}

      {isProcessing && (
        <ProcessingProgress
          progress={progress}
          status={status}
          message={progressMessage}
          onCancel={() => { cancelledRef.current = true; setStatus('idle'); }}
          showPercentage
        />
      )}

      {file && (
        <div className="flex flex-wrap items-center gap-4">
          <Button
            variant="primary"
            size="lg"
            onClick={handleProcess}
            disabled={!file || isProcessing || (watermarkType === 'text' && !watermarkText.trim()) || (watermarkType === 'image' && !imageFile)}
            loading={isProcessing}
          >
            {isProcessing ? t('status.processing') : tTools('addButton')}
          </Button>
          {result && (
            <DownloadButton
              file={result}
              filename={file.name.replace('.pdf', '_watermarked.pdf')}
              variant="secondary"
              size="lg"
              showFileSize
            />
          )}
        </div>
      )}

      {status === 'complete' && result && (
        <div className="p-4 rounded-lg bg-green-50 border border-green-200 text-green-700 dark:bg-green-900/20 dark:border-green-800 dark:text-green-400">
          <p className="text-sm font-medium">{tTools('successMessage')}</p>
        </div>
      )}
    </div>
  );
}

export default WatermarkTool;
