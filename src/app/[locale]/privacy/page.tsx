import { setRequestLocale } from 'next-intl/server';
import { locales, type Locale } from '@/lib/i18n/config';
import PrivacyPageClient from './PrivacyPageClient';

export function generateStaticParams() {
  return locales.map((locale) => ({ locale }));
}

interface PrivacyPageProps {
  params: Promise<{ locale: string }>;
}

export default async function PrivacyPage({ params }: PrivacyPageProps) {
  const { locale } = await params;

  // Enable static rendering
  setRequestLocale(locale);

  return <PrivacyPageClient locale={locale as Locale} />;
}
