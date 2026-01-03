/**
 * PDF Processors
 * Requirements: 5.1
 * 
 * Export all PDF processing implementations.
 */

// Merge processor
export {
  MergePDFProcessor,
  createMergeProcessor,
  mergePDFs,
} from './merge';

// Split processor
export {
  SplitPDFProcessor,
  createSplitProcessor,
  splitPDF,
  parsePageRanges,
  createSplitEveryNPages,
  createSplitEveryPage,
} from './split';

// Organize processor
export {
  OrganizePDFProcessor,
  createOrganizeProcessor,
  organizePDF,
  type OrganizeOptions,
} from './organize';

// Extract pages processor
export {
  ExtractPagesPDFProcessor,
  createExtractProcessor,
  extractPages,
  parsePageSelection,
  type ExtractOptions,
} from './extract';

// Delete pages processor
export {
  DeletePagesPDFProcessor,
  createDeleteProcessor,
  deletePages,
  parsePageSelection as parseDeletePageSelection,
  type DeleteOptions,
} from './delete';

// Alternate merge processor
export {
  AlternateMergePDFProcessor,
  createAlternateMergeProcessor,
  alternateMergePDFs,
  type AlternateMergeOptions,
} from './alternate-merge';

// Divide pages processor
export {
  DividePagesPDFProcessor,
  createDivideProcessor,
  dividePages,
  type DivideOptions,
  type DivisionType,
} from './divide';

// Attachments processors
export {
  AddAttachmentsPDFProcessor,
  ExtractAttachmentsPDFProcessor,
  EditAttachmentsPDFProcessor,
  createAddAttachmentsProcessor,
  createExtractAttachmentsProcessor,
  createEditAttachmentsProcessor,
  addAttachments,
  extractAttachments,
  editAttachments,
  type AttachmentInfo,
  type AddAttachmentsOptions,
  type EditAttachmentsOptions,
} from './attachments';

// PDF to ZIP processor
export {
  PDFToZipProcessor,
  createPDFToZipProcessor,
  packagePDFsToZip,
  type PDFToZipOptions,
} from './pdf-to-zip';

// Redaction processor
export {
  applyRedactions,
  validateRedactionAreas,
  type RedactionArea,
  type RedactionOptions,
  type RedactionResult,
} from './redact';

// Rotate processor
export {
  RotatePDFProcessor,
  createRotateProcessor,
  rotatePDF,
  rotateAllPages,
  type RotateOptions,
} from './rotate';

// Add blank page processor
export {
  AddBlankPagePDFProcessor,
  createAddBlankPageProcessor,
  addBlankPages,
  type AddBlankPageOptions,
} from './add-blank-page';

// Reverse pages processor
export {
  ReversePagesPDFProcessor,
  createReverseProcessor,
  reversePages,
} from './reverse';

// N-Up processor
export {
  NUpPDFProcessor,
  createNUpProcessor,
  createNUpPDF,
  type NUpOptions,
} from './n-up';

// Combine single page processor
export {
  CombineSinglePagePDFProcessor,
  createCombineSinglePageProcessor,
  combineSinglePage,
  type CombineSinglePageOptions,
} from './combine-single-page';

// Posterize processor
export {
  PosterizePDFProcessor,
  createPosterizeProcessor,
  posterizePDF,
  type PosterizeOptions,
} from './posterize';

// Edit metadata processor
export {
  EditMetadataPDFProcessor,
  createEditMetadataProcessor,
  editPDFMetadata,
  type EditMetadataOptions,
} from './edit-metadata';

// Bookmark processor
export {
  BookmarkProcessor,
  createBookmarkProcessor,
  processBookmarks,
  type BookmarkItem,
  type BookmarkOptions,
} from './bookmark';

// Table of contents processor
export {
  TableOfContentsProcessor,
  createTableOfContentsProcessor,
  generateTableOfContents,
  type TOCOptions,
} from './table-of-contents';

// Page numbers processor
export {
  PageNumbersProcessor,
  createPageNumbersProcessor,
  addPageNumbers,
  type PageNumberOptions,
} from './page-numbers';

// Watermark processor
export {
  WatermarkProcessor,
  createWatermarkProcessor,
  addWatermark,
  type WatermarkOptions,
} from './watermark';

// Header footer processor
export {
  HeaderFooterProcessor,
  createHeaderFooterProcessor,
  addHeaderFooter,
  type HeaderFooterOptions,
} from './header-footer';

// Invert colors processor
export {
  InvertColorsProcessor,
  createInvertColorsProcessor,
  invertColors,
  type InvertColorsOptions,
} from './invert-colors';

// Background color processor
export {
  BackgroundColorProcessor,
  createBackgroundColorProcessor,
  addBackgroundColor,
  type BackgroundColorOptions,
} from './background-color';

// Text color processor
export {
  TextColorProcessor,
  createTextColorProcessor,
  changeTextColor,
  type TextColorOptions,
} from './text-color';

// Sign processor
export {
  SignProcessor,
  createSignProcessor,
  signPDF,
  type SignatureItem,
  type SignOptions,
} from './sign';

// Stamps processor
export {
  StampsProcessor,
  createStampsProcessor,
  addStamps,
  type StampItem,
  type StampsOptions,
} from './stamps';

// Remove annotations processor
export {
  RemoveAnnotationsProcessor,
  createRemoveAnnotationsProcessor,
  removeAnnotations,
  type RemoveAnnotationsOptions,
} from './remove-annotations';

// Crop processor
export {
  CropProcessor,
  createCropProcessor,
  cropPDF,
  type CropData,
  type CropOptions,
} from './crop';

// Form filler processor
export {
  FormFillerProcessor,
  createFormFillerProcessor,
  fillForm,
  getFormFields,
  type FormFieldValue,
  type FormFillerOptions,
} from './form-filler';

// Form creator processor
export {
  FormCreatorProcessor,
  createFormCreatorProcessor,
  createForm,
  type FormField,
  type FormCreatorOptions,
} from './form-creator';

// Remove blank pages processor
export {
  RemoveBlankPagesProcessor,
  createRemoveBlankPagesProcessor,
  removeBlankPages,
  type RemoveBlankPagesOptions,
} from './remove-blank-pages';

// Image to PDF processor
export {
  ImageToPDFProcessor,
  createImageToPDFProcessor,
  imagesToPDF,
  PAGE_SIZES,
  type ImageToPDFOptions,
  type PageSizeType,
} from './image-to-pdf';

// Text to PDF processor
export {
  TextToPDFProcessor,
  createTextToPDFProcessor,
  textToPDF,
  TEXT_PAGE_SIZES,
  type TextToPDFOptions,
  type TextPageSizeType,
  type FontId,
} from './text-to-pdf';

// JSON to PDF processor
export {
  JSONToPDFProcessor,
  createJSONToPDFProcessor,
  jsonToPDF,
  JSON_PAGE_SIZES,
  type JSONToPDFOptions,
  type JSONPageSizeType,
} from './json-to-pdf';

// PDF to Image processor
export {
  PDFToImageProcessor,
  createPDFToImageProcessor,
  pdfToImages,
  type PDFToImageOptions,
  type ImageFormat,
} from './pdf-to-image';

// PDF to Greyscale processor
export {
  PDFToGreyscaleProcessor,
  createPDFToGreyscaleProcessor,
  pdfToGreyscale,
  type PDFToGreyscaleOptions,
} from './pdf-to-greyscale';

// PDF to JSON processor
export {
  PDFToJSONProcessor,
  createPDFToJSONProcessor,
  pdfToJSON,
  type PDFToJSONOptions,
  type PDFJSONOutput,
  type OutlineItem,
  type PageContent,
} from './pdf-to-json';

// OCR processor
export {
  OCRProcessor,
  createOCRProcessor,
  ocrPDF,
  OCR_LANGUAGE_NAMES,
  type OCROptions,
  type OCRLanguage,
} from './ocr';

// Fix page size processor
export {
  FixPageSizeProcessor,
  createFixPageSizeProcessor,
  fixPageSize,
  PAGE_SIZES as FIX_PAGE_SIZES,
  type FixPageSizeOptions,
  type PageSizePreset,
  type ScaleMode,
} from './fix-page-size';

// Linearize processor
export {
  LinearizePDFProcessor,
  createLinearizeProcessor,
  linearizePDF,
  type LinearizePDFOptions,
} from './linearize';

// Page dimensions processor
export {
  PageDimensionsProcessor,
  createPageDimensionsProcessor,
  analyzePageDimensions,
  type PageDimensionsOptions,
  type PageDimension,
} from './page-dimensions';

// Remove restrictions processor
export {
  RemoveRestrictionsProcessor,
  createRemoveRestrictionsProcessor,
  removeRestrictions,
  type RemoveRestrictionsOptions,
} from './remove-restrictions';

// Compress processor
export {
  CompressPDFProcessor,
  createCompressProcessor,
  compressPDF,
  type CompressPDFOptions,
  type CompressionQuality,
} from './compress';

// Repair processor
export {
  RepairPDFProcessor,
  createRepairProcessor,
  repairPDF,
  type RepairPDFOptions,
} from './repair';

// Encrypt processor
export {
  EncryptPDFProcessor,
  createEncryptProcessor,
  encryptPDF,
  type EncryptPDFOptions,
  type PDFPermissions,
} from './encrypt';

// Decrypt processor
export {
  DecryptPDFProcessor,
  createDecryptProcessor,
  decryptPDF,
  type DecryptPDFOptions,
} from './decrypt';

// Sanitize processor
export {
  SanitizePDFProcessor,
  createSanitizeProcessor,
  sanitizePDF,
  type SanitizePDFOptions,
} from './sanitize';

// Flatten processor
export {
  FlattenPDFProcessor,
  createFlattenProcessor,
  flattenPDF,
  type FlattenPDFOptions,
} from './flatten';

// Remove metadata processor
export {
  RemoveMetadataProcessor,
  createRemoveMetadataProcessor,
  removeMetadata,
  type RemoveMetadataOptions,
} from './remove-metadata';

// Change permissions processor
export {
  ChangePermissionsProcessor,
  createChangePermissionsProcessor,
  changePermissions,
  type ChangePermissionsOptions,
  type PDFPermissionSettings,
} from './change-permissions';
