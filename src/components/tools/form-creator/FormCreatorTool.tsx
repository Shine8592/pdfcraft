'use client';

import React, { useState, useCallback, useRef, useEffect } from 'react';
import { useTranslations } from 'next-intl';
import * as pdfjsLib from 'pdfjs-dist';
import { FileUploader } from '../FileUploader';
import { ProcessingProgress, ProcessingStatus } from '../ProcessingProgress';
import { DownloadButton } from '../DownloadButton';
import { Button } from '@/components/ui/Button';
import { Card } from '@/components/ui/Card';
import { createForm, FormField } from '@/lib/pdf/processors/form-creator';
import { configurePdfjsWorker } from '@/lib/pdf/loader';
import type { ProcessOutput } from '@/types/pdf';

// Set worker source
if (typeof window !== 'undefined' && !pdfjsLib.GlobalWorkerOptions.workerSrc) {
  configurePdfjsWorker(pdfjsLib);
}

export interface FormCreatorToolProps {
  className?: string;
}

type FieldType = 'text' | 'checkbox' | 'dropdown' | 'radio';

interface VisualField extends FormField {
  id: string;
  selected?: boolean;
}

export function FormCreatorTool({ className = '' }: FormCreatorToolProps) {
  const t = useTranslations('common');
  const tTools = useTranslations('tools');

  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<ProcessingStatus>('idle');
  const [progress, setProgress] = useState(0);
  const [progressMessage, setProgressMessage] = useState('');
  const [result, setResult] = useState<Blob | null>(null);
  const [error, setError] = useState<string | null>(null);

  // Visual editor state
  const [fields, setFields] = useState<VisualField[]>([]);
  const [selectedFieldId, setSelectedFieldId] = useState<string | null>(null);
  const [currentTool, setCurrentTool] = useState<FieldType | 'select'>('select');
  const [totalPages, setTotalPages] = useState(0);
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState({ width: 0, height: 0 });
  const [scale, setScale] = useState(1);

  // Drag state
  const [isDragging, setIsDragging] = useState(false);
  const [isResizing, setIsResizing] = useState(false);
  const [dragStart, setDragStart] = useState({ x: 0, y: 0 });
  const [dragOffset, setDragOffset] = useState({ x: 0, y: 0 });

  const canvasRef = useRef<HTMLCanvasElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const cancelledRef = useRef(false);
  const pdfDocRef = useRef<pdfjsLib.PDFDocumentProxy | null>(null);

  // Generate unique ID
  const generateId = () => `field_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;

  // Load PDF
  const loadPdf = useCallback(async (pdfFile: File) => {
    try {
      const arrayBuffer = await pdfFile.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      pdfDocRef.current = pdf;
      setTotalPages(pdf.numPages);
      setCurrentPage(1);
      renderPage(pdf, 1);
    } catch (err) {
      console.error('Failed to load PDF:', err);
      setError('Failed to load PDF file.');
    }
  }, []);

  // Render page
  const renderPage = async (pdf: pdfjsLib.PDFDocumentProxy, pageNum: number) => {
    if (!canvasRef.current) return;

    try {
      const page = await pdf.getPage(pageNum);
      const viewport = page.getViewport({ scale: 1 });
      
      // Calculate scale to fit container
      const containerWidth = containerRef.current?.clientWidth || 600;
      const newScale = Math.min(containerWidth / viewport.width, 1);
      setScale(newScale);

      const scaledViewport = page.getViewport({ scale: newScale });
      setPageSize({ width: scaledViewport.width, height: scaledViewport.height });

      const canvas = canvasRef.current;
      const ctx = canvas.getContext('2d');
      if (!ctx) return;

      canvas.width = scaledViewport.width;
      canvas.height = scaledViewport.height;

      await page.render({ canvasContext: ctx, viewport: scaledViewport }).promise;
    } catch (err) {
      console.error('Failed to render page:', err);
    }
  };

  // Re-render when page changes
  useEffect(() => {
    if (pdfDocRef.current && currentPage > 0) {
      renderPage(pdfDocRef.current, currentPage);
    }
  }, [currentPage]);

  // Handle canvas click to add field
  const handleCanvasClick = useCallback((e: React.MouseEvent<HTMLDivElement>) => {
    if (currentTool === 'select' || !canvasRef.current) return;

    const rect = canvasRef.current.getBoundingClientRect();
    const x = (e.clientX - rect.left) / scale;
    const y = (e.clientY - rect.top) / scale;

    // Convert to PDF coordinates (origin at bottom-left)
    const pdfY = (pageSize.height / scale) - y;

    const defaultSizes: Record<FieldType, { width: number; height: number }> = {
      text: { width: 200, height: 24 },
      checkbox: { width: 20, height: 20 },
      dropdown: { width: 200, height: 24 },
      radio: { width: 20, height: 20 },
    };

    const newField: VisualField = {
      id: generateId(),
      type: currentTool,
      name: `${currentTool}_${fields.filter(f => f.type === currentTool).length + 1}`,
      pageNumber: currentPage,
      x: x,
      y: pdfY,
      width: defaultSizes[currentTool].width,
      height: defaultSizes[currentTool].height,
      options: currentTool === 'dropdown' || currentTool === 'radio' ? ['Option 1', 'Option 2', 'Option 3'] : undefined,
    };

    setFields(prev => [...prev, newField]);
    setSelectedFieldId(newField.id);
    setCurrentTool('select');
  }, [currentTool, currentPage, fields, pageSize, scale]);

  // Handle field selection
  const handleFieldClick = useCallback((e: React.MouseEvent, fieldId: string) => {
    e.stopPropagation();
    if (currentTool === 'select') {
      setSelectedFieldId(fieldId);
    }
  }, [currentTool]);

  // Handle field drag start
  const handleFieldMouseDown = useCallback((e: React.MouseEvent, fieldId: string, isResize: boolean = false) => {
    e.stopPropagation();
    e.preventDefault();
    
    if (currentTool !== 'select') return;

    const field = fields.find(f => f.id === fieldId);
    if (!field) return;

    setSelectedFieldId(fieldId);
    setDragStart({ x: e.clientX, y: e.clientY });
    setDragOffset({ x: field.x, y: field.y });

    if (isResize) {
      setIsResizing(true);
    } else {
      setIsDragging(true);
    }
  }, [currentTool, fields]);

  // Handle mouse move for drag/resize
  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (!isDragging && !isResizing) return;
      if (!selectedFieldId) return;

      const deltaX = (e.clientX - dragStart.x) / scale;
      const deltaY = (e.clientY - dragStart.y) / scale;

      setFields(prev => prev.map(field => {
        if (field.id !== selectedFieldId) return field;

        if (isResizing) {
          return {
            ...field,
            width: Math.max(20, field.width + deltaX),
            height: Math.max(20, field.height + deltaY),
          };
        } else {
          return {
            ...field,
            x: dragOffset.x + deltaX,
            y: dragOffset.y - deltaY, // Invert Y for PDF coordinates
          };
        }
      }));

      setDragStart({ x: e.clientX, y: e.clientY });
      if (!isResizing) {
        setDragOffset(prev => ({ x: prev.x + deltaX, y: prev.y - deltaY }));
      }
    };

    const handleMouseUp = () => {
      setIsDragging(false);
      setIsResizing(false);
    };

    if (isDragging || isResizing) {
      window.addEventListener('mousemove', handleMouseMove);
      window.addEventListener('mouseup', handleMouseUp);
    }

    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [isDragging, isResizing, selectedFieldId, dragStart, dragOffset, scale]);

  // Delete selected field
  const handleDeleteField = useCallback(() => {
    if (!selectedFieldId) return;
    setFields(prev => prev.filter(f => f.id !== selectedFieldId));
    setSelectedFieldId(null);
  }, [selectedFieldId]);

  // Update selected field
  const updateSelectedField = useCallback((updates: Partial<VisualField>) => {
    if (!selectedFieldId) return;
    setFields(prev => prev.map(f => f.id === selectedFieldId ? { ...f, ...updates } : f));
  }, [selectedFieldId]);

  // Handle file selection
  const handleFilesSelected = useCallback((files: File[]) => {
    if (files.length > 0) {
      setFile(files[0]);
      setError(null);
      setResult(null);
      setFields([]);
      setSelectedFieldId(null);
      loadPdf(files[0]);
    }
  }, [loadPdf]);

  // Clear file
  const handleClearFile = useCallback(() => {
    setFile(null);
    setResult(null);
    setError(null);
    setStatus('idle');
    setFields([]);
    setSelectedFieldId(null);
    setTotalPages(0);
    setCurrentPage(1);
    pdfDocRef.current = null;
  }, []);

  // Process PDF
  const handleProcess = useCallback(async () => {
    if (!file || fields.length === 0) return;

    cancelledRef.current = false;
    setStatus('processing');
    setProgress(0);
    setError(null);
    setResult(null);

    try {
      // Convert visual fields to form fields
      const formFields: FormField[] = fields.map(({ id, selected, ...field }) => field);

      const output: ProcessOutput = await createForm(file, { fields: formFields }, (prog, message) => {
        if (!cancelledRef.current) {
          setProgress(prog);
          setProgressMessage(message || '');
        }
      });

      if (output.success && output.result) {
        setResult(output.result as Blob);
        setStatus('complete');
      } else {
        setError(output.error?.message || 'Failed to create form.');
        setStatus('error');
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An unexpected error occurred.');
      setStatus('error');
    }
  }, [file, fields]);

  const formatSize = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  const isProcessing = status === 'processing';
  const selectedField = fields.find(f => f.id === selectedFieldId);
  const currentPageFields = fields.filter(f => f.pageNumber === currentPage);

  // Tool buttons
  const tools = [
    { type: 'select' as const, icon: '↖', labelKey: 'selectTool' },
    { type: 'text' as const, icon: 'T', labelKey: 'textFieldTool' },
    { type: 'checkbox' as const, icon: '☑', labelKey: 'checkboxTool' },
    { type: 'dropdown' as const, icon: '▼', labelKey: 'dropdownTool' },
    { type: 'radio' as const, icon: '◉', labelKey: 'radioTool' },
  ];

  // Get field style for overlay
  const getFieldStyle = (field: VisualField): React.CSSProperties => {
    const pdfHeight = pageSize.height / scale;
    const screenY = (pdfHeight - field.y) * scale;
    
    return {
      position: 'absolute',
      left: field.x * scale,
      top: screenY - (field.height * scale),
      width: field.width * scale,
      height: field.height * scale,
      border: field.id === selectedFieldId ? '2px solid #3b82f6' : '2px dashed #6b7280',
      backgroundColor: field.id === selectedFieldId ? 'rgba(59, 130, 246, 0.1)' : 'rgba(107, 114, 128, 0.1)',
      cursor: currentTool === 'select' ? 'move' : 'default',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      fontSize: '10px',
      color: '#374151',
      userSelect: 'none',
    };
  };

  // Get field icon
  const getFieldIcon = (type: FieldType) => {
    switch (type) {
      case 'text': return '📝';
      case 'checkbox': return '☑';
      case 'dropdown': return '▼';
      case 'radio': return '◉';
    }
  };

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
          label={tTools('formCreator.uploadLabel') || 'Upload PDF File'}
          description={tTools('formCreator.uploadDescription') || 'Drag and drop a PDF file here.'}
        />
      )}

      {error && (
        <div className="p-4 rounded-[var(--radius-md)] bg-red-50 border border-red-200 text-red-700">
          <p className="text-sm">{error}</p>
        </div>
      )}

      {file && (
        <div className="grid lg:grid-cols-3 gap-6">
          {/* Visual Editor */}
          <div className="lg:col-span-2 space-y-4">
            {/* Toolbar */}
            <Card variant="outlined" size="sm">
              <div className="flex flex-wrap items-center gap-2">
                {/* Tool buttons */}
                <div className="flex items-center gap-1 border-r border-gray-200 pr-3 mr-2">
                  {tools.map(tool => (
                    <button
                      key={tool.type}
                      onClick={() => setCurrentTool(tool.type)}
                      className={`
                        px-3 py-2 rounded-[var(--radius-sm)] text-sm font-medium transition-colors
                        ${currentTool === tool.type
                          ? 'bg-blue-500 text-white'
                          : 'bg-gray-100 hover:bg-gray-200 text-gray-700'
                        }
                      `}
                      title={tTools(`formCreator.${tool.labelKey}`)}
                    >
                      <span className="mr-1">{tool.icon}</span>
                      <span className="hidden sm:inline">{tTools(`formCreator.${tool.labelKey}`)}</span>
                    </button>
                  ))}
                </div>

                {/* Page navigation */}
                <div className="flex items-center gap-2">
                  <Button
                    variant="ghost"
                    size="sm"
                    onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                    disabled={currentPage <= 1}
                  >
                    ←
                  </Button>
                  <span className="text-sm">
                    Page {currentPage} / {totalPages}
                  </span>
                  <Button
                    variant="ghost"
                    size="sm"
                    onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                    disabled={currentPage >= totalPages}
                  >
                    →
                  </Button>
                </div>

                {/* Delete button */}
                {selectedFieldId && (
                  <Button
                    variant="ghost"
                    size="sm"
                    onClick={handleDeleteField}
                    className="ml-auto text-red-500 hover:text-red-700"
                  >
                    🗑 {tTools('formCreator.deleteTool')}
                  </Button>
                )}
              </div>
            </Card>

            {/* Canvas with field overlays */}
            <Card variant="outlined" size="sm">
              <div 
                ref={containerRef}
                className="relative bg-gray-100 rounded-[var(--radius-md)] overflow-auto"
                style={{ maxHeight: '600px' }}
              >
                <div
                  className="relative inline-block"
                  onClick={handleCanvasClick}
                  style={{ cursor: currentTool !== 'select' ? 'crosshair' : 'default' }}
                >
                  <canvas
                    ref={canvasRef}
                    className="shadow-lg bg-white"
                  />
                  
                  {/* Field overlays */}
                  {currentPageFields.map(field => (
                    <div
                      key={field.id}
                      style={getFieldStyle(field)}
                      onClick={(e) => handleFieldClick(e, field.id)}
                      onMouseDown={(e) => handleFieldMouseDown(e, field.id)}
                    >
                      <span>{getFieldIcon(field.type)} {field.name}</span>
                      
                      {/* Resize handle */}
                      {field.id === selectedFieldId && (
                        <div
                          className="absolute bottom-0 right-0 w-3 h-3 bg-blue-500 cursor-se-resize"
                          onMouseDown={(e) => handleFieldMouseDown(e, field.id, true)}
                        />
                      )}
                    </div>
                  ))}
                </div>
              </div>

              <p className="text-xs text-gray-500 mt-2 text-center">
                {currentTool === 'select' 
                  ? tTools('formCreator.selectHint')
                  : tTools('formCreator.addFieldHint', { type: tTools(`formCreator.${tools.find(t => t.type === currentTool)?.labelKey || 'textFieldTool'}`) })
                }
              </p>
            </Card>
          </div>

          {/* Properties Panel */}
          <div className="space-y-4">
            {/* File info */}
            <Card variant="outlined">
              <div className="flex items-center justify-between">
                <div>
                  <p className="font-medium text-sm truncate">{file.name}</p>
                  <p className="text-xs text-gray-500">
                    {formatSize(file.size)} • {totalPages} pages
                  </p>
                </div>
                <Button variant="ghost" size="sm" onClick={handleClearFile} disabled={isProcessing}>
                  ✕
                </Button>
              </div>
            </Card>

            {/* Field properties */}
            {selectedField ? (
              <Card variant="outlined" size="lg">
                <h3 className="text-lg font-medium mb-4">{tTools('formCreator.fieldProperties')}</h3>
                
                <div className="space-y-4">
                  <div>
                    <label className="block text-sm font-medium mb-1">{tTools('formCreator.fieldName')}</label>
                    <input
                      type="text"
                      value={selectedField.name}
                      onChange={(e) => updateSelectedField({ name: e.target.value })}
                      className="w-full px-3 py-2 border rounded-[var(--radius-md)] text-sm"
                    />
                  </div>

                  <div>
                    <label className="block text-sm font-medium mb-1">{tTools('formCreator.fieldType')}</label>
                    <select
                      value={selectedField.type}
                      onChange={(e) => updateSelectedField({ type: e.target.value as FieldType })}
                      className="w-full px-3 py-2 border rounded-[var(--radius-md)] text-sm"
                    >
                      <option value="text">{tTools('formCreator.textFieldTool')}</option>
                      <option value="checkbox">{tTools('formCreator.checkboxTool')}</option>
                      <option value="dropdown">{tTools('formCreator.dropdownTool')}</option>
                      <option value="radio">{tTools('formCreator.radioTool')}</option>
                    </select>
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-sm font-medium mb-1">{tTools('formCreator.fieldX')}</label>
                      <input
                        type="number"
                        value={Math.round(selectedField.x)}
                        onChange={(e) => updateSelectedField({ x: parseInt(e.target.value) || 0 })}
                        className="w-full px-3 py-2 border rounded-[var(--radius-md)] text-sm"
                      />
                    </div>
                    <div>
                      <label className="block text-sm font-medium mb-1">{tTools('formCreator.fieldY')}</label>
                      <input
                        type="number"
                        value={Math.round(selectedField.y)}
                        onChange={(e) => updateSelectedField({ y: parseInt(e.target.value) || 0 })}
                        className="w-full px-3 py-2 border rounded-[var(--radius-md)] text-sm"
                      />
                    </div>
                    <div>
                      <label className="block text-sm font-medium mb-1">{tTools('formCreator.fieldWidth')}</label>
                      <input
                        type="number"
                        value={Math.round(selectedField.width)}
                        onChange={(e) => updateSelectedField({ width: parseInt(e.target.value) || 20 })}
                        className="w-full px-3 py-2 border rounded-[var(--radius-md)] text-sm"
                      />
                    </div>
                    <div>
                      <label className="block text-sm font-medium mb-1">{tTools('formCreator.fieldHeight')}</label>
                      <input
                        type="number"
                        value={Math.round(selectedField.height)}
                        onChange={(e) => updateSelectedField({ height: parseInt(e.target.value) || 20 })}
                        className="w-full px-3 py-2 border rounded-[var(--radius-md)] text-sm"
                      />
                    </div>
                  </div>

                  {selectedField.type === 'text' && (
                    <label className="flex items-center gap-2">
                      <input
                        type="checkbox"
                        checked={selectedField.multiline || false}
                        onChange={(e) => updateSelectedField({ multiline: e.target.checked })}
                        className="w-4 h-4"
                      />
                      <span className="text-sm">{tTools('formCreator.multiline')}</span>
                    </label>
                  )}

                  {(selectedField.type === 'dropdown' || selectedField.type === 'radio') && (
                    <div>
                      <label className="block text-sm font-medium mb-1">{tTools('formCreator.options')}</label>
                      <textarea
                        value={(selectedField.options || []).join('\n')}
                        onChange={(e) => updateSelectedField({ options: e.target.value.split('\n').filter(o => o.trim()) })}
                        className="w-full px-3 py-2 border rounded-[var(--radius-md)] text-sm"
                        rows={4}
                      />
                    </div>
                  )}

                  <Button
                    variant="ghost"
                    size="sm"
                    onClick={handleDeleteField}
                    className="w-full text-red-500 hover:bg-red-50"
                  >
                    {tTools('formCreator.deleteField')}
                  </Button>
                </div>
              </Card>
            ) : (
              <Card variant="outlined" size="lg">
                <h3 className="text-lg font-medium mb-4">{tTools('formCreator.fieldProperties')}</h3>
                <p className="text-sm text-gray-500">
                  {tTools('formCreator.noFieldSelected')}
                </p>
              </Card>
            )}

            {/* Fields list */}
            <Card variant="outlined" size="lg">
              <h3 className="text-lg font-medium mb-4">
                {tTools('formCreator.allFields', { count: fields.length })}
              </h3>
              
              {fields.length === 0 ? (
                <p className="text-sm text-gray-500">
                  {tTools('formCreator.noFieldsYet')}
                </p>
              ) : (
                <div className="space-y-2 max-h-48 overflow-y-auto">
                  {fields.map(field => (
                    <div
                      key={field.id}
                      onClick={() => {
                        setSelectedFieldId(field.id);
                        setCurrentPage(field.pageNumber);
                      }}
                      className={`
                        p-2 rounded-[var(--radius-sm)] cursor-pointer text-sm
                        flex items-center justify-between
                        ${field.id === selectedFieldId
                          ? 'bg-blue-50 border border-blue-200'
                          : 'bg-gray-50 hover:bg-gray-100'
                        }
                      `}
                    >
                      <span>
                        {getFieldIcon(field.type)} {field.name}
                      </span>
                      <span className="text-xs text-gray-500">
                        {tTools('formCreator.page', { number: field.pageNumber })}
                      </span>
                    </div>
                  ))}
                </div>
              )}
            </Card>
          </div>
        </div>
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
            disabled={!file || fields.length === 0 || isProcessing}
            loading={isProcessing}
          >
            {isProcessing ? 'Processing...' : (tTools('formCreator.createButton') || 'Create Form')}
          </Button>
          {result && (
            <DownloadButton
              file={result}
              filename={file.name.replace('.pdf', '_form.pdf')}
              variant="secondary"
              size="lg"
              showFileSize
            />
          )}
        </div>
      )}

      {status === 'complete' && result && (
        <div className="p-4 rounded-[var(--radius-md)] bg-green-50 border border-green-200 text-green-700">
          <p className="text-sm font-medium">
            {tTools('formCreator.successMessage') || 'Form created successfully!'}
          </p>
        </div>
      )}
    </div>
  );
}

export default FormCreatorTool;
