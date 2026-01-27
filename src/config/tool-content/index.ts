/**
 * Tool content exports for all languages
 * Requirements: 3.1 - Multi-language support
 */

export { toolContentEn } from './en';
export { toolContentJa } from './ja';
export { toolContentKo } from './ko';
export { toolContentEs } from './es';
export { toolContentFr } from './fr';
export { toolContentDe } from './de';
export { toolContentZh } from './zh';
export { toolContentPt } from './pt';
export { toolContentAr } from './ar';

import { toolContentEn } from './en';
import { toolContentJa } from './ja';
import { toolContentKo } from './ko';
import { toolContentEs } from './es';
import { toolContentFr } from './fr';
import { toolContentDe } from './de';
import { toolContentZh } from './zh';
import { toolContentPt } from './pt';
import { toolContentAr } from './ar';
import { ToolContent } from '@/types/tool';

export type Locale = 'en' | 'ja' | 'ko' | 'es' | 'fr' | 'de' | 'zh' | 'zh-TW' | 'pt' | 'ar';

/**
 * Get tool content for a specific locale
 * Falls back to English if translation not found
 * zh-TW falls back to zh (Simplified Chinese) content
 * ar falls back to en content for now
 */
export function getToolContent(locale: Locale, toolId: string): ToolContent | undefined {
  const contentMap: Record<string, Record<string, ToolContent>> = {
    en: toolContentEn,
    ja: toolContentJa,
    ko: toolContentKo,
    es: toolContentEs,
    fr: toolContentFr,
    de: toolContentDe,
    zh: toolContentZh,
    pt: toolContentPt,
    ar: toolContentAr,
  };

  // Map zh-TW to zh (use Simplified Chinese content for Traditional Chinese)
  const effectiveLocale = (locale === 'zh-TW' ? 'zh' : locale) as string;

  const localeContent = contentMap[effectiveLocale];
  if (localeContent && localeContent[toolId]) {
    return localeContent[toolId];
  }

  // Fallback to English
  return toolContentEn[toolId];
}

