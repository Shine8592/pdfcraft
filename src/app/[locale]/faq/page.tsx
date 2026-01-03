import { setRequestLocale } from 'next-intl/server';
import { locales, type Locale } from '@/lib/i18n/config';
import FAQPageClient from './FAQPageClient';

export function generateStaticParams() {
  return locales.map((locale) => ({ locale }));
}

interface FAQPageProps {
  params: Promise<{ locale: string }>;
}

export default async function FAQPage({ params }: FAQPageProps) {
  const { locale } = await params;

  // Enable static rendering
  setRequestLocale(locale);

  return <FAQPageClient locale={locale as Locale} />;
}
