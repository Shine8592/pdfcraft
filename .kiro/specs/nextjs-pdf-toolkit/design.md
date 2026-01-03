# Design Document

## Overview

PDFCraft是一个基于Next.js 14的现代化PDF工具平台，采用App Router架构，支持9种语言的国际化，提供67+个PDF处理工具。系统采用客户端处理模式，所有PDF操作在用户浏览器中完成，确保隐私安全。

### 核心技术栈
- **框架**: Next.js 14 (App Router)
- **语言**: TypeScript 5.x
- **样式**: Tailwind CSS 4.x
- **状态管理**: Zustand
- **国际化**: next-intl
- **PDF处理**: pdf-lib, pdfjs-dist, coherentpdf, qpdf-wasm
- **测试**: Vitest + fast-check (Property-Based Testing)

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                        PDFCraft Architecture                      │
├─────────────────────────────────────────────────────────────────┤
│                                                                   │
│  ┌─────────────┐    ┌─────────────┐    ┌─────────────┐          │
│  │   Browser   │    │  Next.js    │    │   Static    │          │
│  │   Client    │◄──►│  App Router │◄──►│   Assets    │          │
│  └─────────────┘    └─────────────┘    └─────────────┘          │
│         │                  │                                      │
│         ▼                  ▼                                      │
│  ┌─────────────┐    ┌─────────────┐                              │
│  │ Web Workers │    │    i18n     │                              │
│  │ (PDF Proc)  │    │  Messages   │                              │
│  └─────────────┘    └─────────────┘                              │
│                                                                   │
└─────────────────────────────────────────────────────────────────┘
```

### 目录结构

```
pdfcraft/
├── app/
│   ├── [locale]/
│   │   ├── layout.tsx              # 带语言的根布局
│   │   ├── page.tsx                # 首页
│   │   ├── tools/
│   │   │   ├── page.tsx            # 工具列表页
│   │   │   └── [tool]/
│   │   │       └── page.tsx        # 工具详情页
│   │   ├── about/
│   │   ├── faq/
│   │   └── privacy/
│   ├── api/                        # API路由（如需要）
│   └── globals.css
├── components/
│   ├── ui/                         # 基础UI组件
│   ├── layout/                     # 布局组件
│   ├── tools/                      # 工具相关组件
│   └── common/                     # 通用组件
├── lib/
│   ├── pdf/                        # PDF处理逻辑
│   ├── i18n/                       # 国际化配置
│   ├── hooks/                      # 自定义Hooks
│   └── utils/                      # 工具函数
├── messages/                       # 翻译文件
│   ├── en.json
│   ├── ja.json
│   ├── ko.json
│   ├── es.json
│   ├── fr.json
│   ├── de.json
│   ├── zh.json
│   ├── pt.json
│   └── ar.json
├── public/
│   ├── images/
│   ├── workers/                    # Web Worker文件
│   └── fonts/
├── types/                          # TypeScript类型定义
└── config/                         # 配置文件
```

## Components and Interfaces

### 1. 核心组件

#### Layout Components

```typescript
// components/layout/Header.tsx
interface HeaderProps {
  locale: Locale;
  showSearch?: boolean;
}

// components/layout/Footer.tsx
interface FooterProps {
  locale: Locale;
}

// components/layout/Navigation.tsx
interface NavigationProps {
  locale: Locale;
  currentPath: string;
}

// components/layout/MobileMenu.tsx
interface MobileMenuProps {
  isOpen: boolean;
  onClose: () => void;
  locale: Locale;
}
```

#### Tool Components

```typescript
// components/tools/ToolCard.tsx
interface ToolCardProps {
  tool: Tool;
  locale: Locale;
}

// components/tools/ToolGrid.tsx
interface ToolGridProps {
  tools: Tool[];
  category?: string;
  locale: Locale;
}

// components/tools/ToolPage.tsx
interface ToolPageProps {
  tool: Tool;
  locale: Locale;
  content: ToolContent;
}

// components/tools/FileUploader.tsx
interface FileUploaderProps {
  accept: string[];
  multiple?: boolean;
  maxSize?: number;
  onFilesSelected: (files: File[]) => void;
  onError: (error: string) => void;
}

// components/tools/ProcessingProgress.tsx
interface ProcessingProgressProps {
  progress: number;
  status: 'idle' | 'processing' | 'complete' | 'error';
  message?: string;
}

// components/tools/DownloadButton.tsx
interface DownloadButtonProps {
  file: Blob | null;
  filename: string;
  disabled?: boolean;
}
```

#### UI Components

```typescript
// components/ui/Button.tsx
interface ButtonProps {
  variant: 'primary' | 'secondary' | 'outline' | 'ghost';
  size: 'sm' | 'md' | 'lg';
  loading?: boolean;
  disabled?: boolean;
  children: React.ReactNode;
  onClick?: () => void;
}

// components/ui/Card.tsx
interface CardProps {
  children: React.ReactNode;
  className?: string;
  hover?: boolean;
}

// components/ui/Modal.tsx
interface ModalProps {
  isOpen: boolean;
  onClose: () => void;
  title: string;
  children: React.ReactNode;
}

// components/ui/Tabs.tsx
interface TabsProps {
  tabs: { id: string; label: string; content: React.ReactNode }[];
  defaultTab?: string;
}
```

### 2. PDF处理接口

```typescript
// lib/pdf/types.ts
interface PDFProcessor {
  process(input: ProcessInput): Promise<ProcessOutput>;
  validate(file: File): ValidationResult;
  getProgress(): number;
  cancel(): void;
}

interface ProcessInput {
  files: File[];
  options: Record<string, unknown>;
}

interface ProcessOutput {
  success: boolean;
  result?: Blob | Blob[];
  error?: string;
  metadata?: Record<string, unknown>;
}

interface ValidationResult {
  valid: boolean;
  errors: string[];
}

// lib/pdf/processors/merge.ts
interface MergeOptions {
  preserveBookmarks: boolean;
  pageOrder: 'sequential' | 'interleaved';
}

// lib/pdf/processors/split.ts
interface SplitOptions {
  ranges: PageRange[];
  outputFormat: 'single' | 'multiple';
}

interface PageRange {
  start: number;
  end: number;
}

// lib/pdf/processors/compress.ts
interface CompressOptions {
  quality: 'low' | 'medium' | 'high';
  removeMetadata: boolean;
  optimizeImages: boolean;
}
```

### 3. 国际化接口

```typescript
// lib/i18n/types.ts
type Locale = 'en' | 'ja' | 'ko' | 'es' | 'fr' | 'de' | 'zh' | 'pt' | 'ar';

interface LocaleConfig {
  code: Locale;
  name: string;
  nativeName: string;
  direction: 'ltr' | 'rtl';
  dateFormat: string;
}

// 翻译消息结构
interface Messages {
  common: CommonMessages;
  tools: ToolMessages;
  errors: ErrorMessages;
  seo: SEOMessages;
}

interface ToolMessages {
  [toolId: string]: {
    name: string;
    subtitle: string;
    description: string;
    howToUse: string[];
    useCases: UseCase[];
    faq: FAQ[];
  };
}
```

## Data Models

### Tool Model

```typescript
interface Tool {
  id: string;                    // 唯一标识符，如 'merge-pdf'
  slug: string;                  // URL路径，如 'merge-pdf'
  icon: string;                  // Lucide图标名称
  category: ToolCategory;        // 工具分类
  acceptedFormats: string[];     // 接受的文件格式
  outputFormat: string;          // 输出格式
  maxFileSize: number;           // 最大文件大小（字节）
  maxFiles: number;              // 最大文件数量
  features: string[];            // 功能特性列表
  relatedTools: string[];        // 相关工具ID
}

type ToolCategory = 
  | 'popular'
  | 'edit-annotate'
  | 'convert-to-pdf'
  | 'convert-from-pdf'
  | 'organize-manage'
  | 'optimize-repair'
  | 'secure-pdf';
```

### Tool Content Model (SEO)

```typescript
interface ToolContent {
  title: string;
  metaDescription: string;
  keywords: string[];
  description: string;           // 详细描述（支持Markdown）
  howToUse: HowToStep[];
  useCases: UseCase[];
  faq: FAQ[];
  screenshots?: Screenshot[];
}

interface HowToStep {
  step: number;
  title: string;
  description: string;
  image?: string;
}

interface UseCase {
  title: string;
  description: string;
  icon: string;
}

interface FAQ {
  question: string;
  answer: string;
}

interface Screenshot {
  src: string;
  alt: string;
  caption: string;
}
```

### User Preferences Model

```typescript
interface UserPreferences {
  locale: Locale;
  theme: 'light' | 'dark' | 'system';
  recentTools: string[];         // 最近使用的工具ID
  recentFiles: RecentFile[];     // 最近处理的文件
  defaultOptions: Record<string, unknown>;
}

interface RecentFile {
  name: string;
  size: number;
  processedAt: Date;
  toolUsed: string;
}
```

### Processing State Model

```typescript
interface ProcessingState {
  status: 'idle' | 'uploading' | 'processing' | 'complete' | 'error';
  progress: number;              // 0-100
  currentStep: string;
  files: UploadedFile[];
  result: ProcessResult | null;
  error: ProcessError | null;
}

interface UploadedFile {
  id: string;
  file: File;
  preview?: string;
  status: 'pending' | 'processing' | 'complete' | 'error';
}

interface ProcessResult {
  blob: Blob;
  filename: string;
  size: number;
  metadata?: Record<string, unknown>;
}

interface ProcessError {
  code: string;
  message: string;
  details?: string;
}
```



## Correctness Properties

*A property is a characteristic or behavior that should hold true across all valid executions of a system-essentially, a formal statement about what the system should do. Properties serve as the bridge between human-readable specifications and machine-verifiable correctness guarantees.*

Based on the acceptance criteria analysis, the following correctness properties have been identified:

### Property 1: Meta Tags Completeness
*For any* page in the application and any supported locale, rendering that page SHALL produce HTML containing all required meta tags (title, description, og:title, og:description, twitter:card).
**Validates: Requirements 1.3, 4.1**

### Property 2: Brand Consistency
*For any* rendered page in the application, the page content SHALL contain the brand name "PDFCraft" in the header or title area.
**Validates: Requirements 2.1**

### Property 3: Tool Card Rendering
*For any* tool in the tools configuration, rendering its card component SHALL produce output containing the tool's icon, name, and description.
**Validates: Requirements 2.5**

### Property 4: Language Preference Persistence (Round-Trip)
*For any* supported locale, setting the language preference and then retrieving it SHALL return the same locale value.
**Validates: Requirements 3.2**

### Property 5: URL Locale Prefix
*For any* page path and any supported locale, the generated URL SHALL contain the locale code as a prefix segment (e.g., /en/, /ja/, /ar/).
**Validates: Requirements 3.4**

### Property 6: Translation Fallback
*For any* translation key that does not exist in a non-English locale, requesting that key SHALL return the English translation value.
**Validates: Requirements 3.6**

### Property 7: Tool Page Content Completeness
*For any* tool page, the rendered content SHALL include: a description section, a how-to-use section with at least 3 steps, a use-cases section with at least 3 scenarios, and an FAQ section with at least 3 questions.
**Validates: Requirements 4.2, 4.3, 4.4, 4.5, 12.1-12.5**

### Property 8: Structured Data Presence
*For any* tool page, the rendered HTML SHALL contain valid JSON-LD script tags with @type "SoftwareApplication" and "FAQPage".
**Validates: Requirements 4.7**

### Property 9: Error Message Mapping
*For any* defined error code in the system, there SHALL exist a corresponding user-friendly error message in all supported locales.
**Validates: Requirements 5.5**

### Property 10: Tool Category Assignment
*For any* tool in the system, the tool SHALL be assigned to exactly one of the 7 defined categories.
**Validates: Requirements 6.1**

### Property 11: Search Result Relevance
*For any* search query string, the returned tool results SHALL only include tools whose name or description contains a fuzzy match to the query.
**Validates: Requirements 6.2**

### Property 12: Related Tools Definition
*For any* tool in the system, the tool SHALL have at least 2 related tools defined, and all related tool IDs SHALL reference existing tools.
**Validates: Requirements 6.5**

### Property 13: Project Save Round-Trip
*For any* valid project state saved to IndexedDB, retrieving that project SHALL return data equivalent to the original saved state.
**Validates: Requirements 10.2**

## Error Handling

### Error Categories

```typescript
enum ErrorCategory {
  FILE_ERROR = 'FILE_ERROR',
  PROCESSING_ERROR = 'PROCESSING_ERROR',
  VALIDATION_ERROR = 'VALIDATION_ERROR',
  NETWORK_ERROR = 'NETWORK_ERROR',
  BROWSER_ERROR = 'BROWSER_ERROR',
}

interface AppError {
  code: string;
  category: ErrorCategory;
  message: string;
  details?: string;
  recoverable: boolean;
  suggestedAction?: string;
}
```

### Error Codes

| Code | Category | Description | Suggested Action |
|------|----------|-------------|------------------|
| FILE_TOO_LARGE | FILE_ERROR | File exceeds maximum size limit | Compress or split the file |
| FILE_TYPE_INVALID | FILE_ERROR | File type not supported | Convert to supported format |
| FILE_CORRUPTED | FILE_ERROR | File is corrupted or unreadable | Try a different file |
| PDF_ENCRYPTED | PROCESSING_ERROR | PDF is password protected | Provide password or decrypt first |
| PDF_MALFORMED | PROCESSING_ERROR | PDF structure is invalid | Try repair tool first |
| PROCESSING_TIMEOUT | PROCESSING_ERROR | Operation took too long | Try with smaller file |
| MEMORY_EXCEEDED | BROWSER_ERROR | Browser ran out of memory | Close other tabs, try smaller file |
| WORKER_FAILED | BROWSER_ERROR | Web Worker crashed | Refresh page and retry |

### Error Handling Strategy

1. **Validation First**: Validate files before processing to catch errors early
2. **Graceful Degradation**: If Web Worker fails, fall back to main thread (with warning)
3. **User-Friendly Messages**: All errors display localized, actionable messages
4. **Recovery Options**: Provide clear next steps for each error type
5. **Logging**: Log errors for debugging (client-side only, no server transmission)

## Testing Strategy

### Dual Testing Approach

This project uses both unit testing and property-based testing to ensure comprehensive coverage:

#### Unit Tests
- Test specific examples and edge cases
- Verify component rendering
- Test utility functions with known inputs/outputs
- Integration tests for PDF processing workflows

#### Property-Based Tests
- **Framework**: fast-check (JavaScript property-based testing library)
- **Minimum Iterations**: 100 per property
- **Tag Format**: `**Feature: nextjs-pdf-toolkit, Property {number}: {property_text}**`

### Test Categories

1. **Component Tests**
   - Render tests for all UI components
   - Accessibility tests (ARIA, keyboard navigation)
   - Responsive layout tests

2. **PDF Processing Tests**
   - Unit tests for each processor function
   - Property tests for round-trip operations (serialize/deserialize)
   - Edge case tests (empty files, large files, encrypted files)

3. **i18n Tests**
   - Property tests for translation fallback
   - RTL layout tests for Arabic
   - Locale URL generation tests

4. **Integration Tests**
   - Full workflow tests for each tool
   - File upload/download tests
   - State persistence tests

### Property-Based Test Implementation

```typescript
// Example: Property 6 - Translation Fallback
import { fc } from 'fast-check';
import { getTranslation } from '@/lib/i18n';

/**
 * **Feature: nextjs-pdf-toolkit, Property 6: Translation Fallback**
 * **Validates: Requirements 3.6**
 */
test('missing translations fall back to English', () => {
  fc.assert(
    fc.property(
      fc.constantFrom('ja', 'ko', 'es', 'fr', 'de', 'zh', 'pt', 'ar'),
      fc.string({ minLength: 1 }),
      (locale, key) => {
        const result = getTranslation(locale, `nonexistent.${key}`);
        const englishResult = getTranslation('en', `nonexistent.${key}`);
        return result === englishResult;
      }
    ),
    { numRuns: 100 }
  );
});
```

### Test File Structure

```
src/
├── __tests__/
│   ├── components/
│   │   ├── ToolCard.test.tsx
│   │   ├── FileUploader.test.tsx
│   │   └── ...
│   ├── lib/
│   │   ├── pdf/
│   │   │   ├── merge.test.ts
│   │   │   ├── split.test.ts
│   │   │   └── ...
│   │   └── i18n/
│   │       └── translations.test.ts
│   ├── properties/
│   │   ├── meta-tags.property.test.ts
│   │   ├── tool-content.property.test.ts
│   │   ├── i18n.property.test.ts
│   │   └── ...
│   └── integration/
│       ├── merge-workflow.test.ts
│       └── ...
```
