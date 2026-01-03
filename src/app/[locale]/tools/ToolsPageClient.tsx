'use client';

import { useState, useMemo, useCallback, useEffect } from 'react';
import { useTranslations } from 'next-intl';
import { useSearchParams } from 'next/navigation';
import { Search, X, Filter } from 'lucide-react';
import { Header } from '@/components/layout/Header';
import { Footer } from '@/components/layout/Footer';
import { ToolGrid } from '@/components/tools/ToolGrid';
import { Button } from '@/components/ui/Button';
import { Card } from '@/components/ui/Card';
import { getAllTools, getToolsByCategory } from '@/config/tools';
import { toolMatchesQuery } from '@/lib/utils/search';
import { type Locale } from '@/lib/i18n/config';
import { CATEGORY_INFO, type ToolCategory } from '@/types/tool';

type CategoryFilter = ToolCategory | 'all';

interface ToolsPageClientProps {
  locale: Locale;
  localizedToolContent?: Record<string, { title: string; description: string }>;
}

export default function ToolsPageClient({ locale, localizedToolContent }: ToolsPageClientProps) {
  const t = useTranslations();
  const searchParams = useSearchParams();
  const allTools = getAllTools();

  const categoryTranslationKeys: Record<ToolCategory, string> = {
    'edit-annotate': 'editAnnotate',
    'convert-to-pdf': 'convertToPdf',
    'convert-from-pdf': 'convertFromPdf',
    'organize-manage': 'organizeManage',
    'optimize-repair': 'optimizeRepair',
    'secure-pdf': 'securePdf',
  };

  // Read initial values from URL search params (client-side)
  const initialCategory = searchParams.get('category') || 'all';
  const initialQuery = searchParams.get('q') || '';

  const [searchQuery, setSearchQuery] = useState(initialQuery);
  const [selectedCategory, setSelectedCategory] = useState<CategoryFilter>(
    (initialCategory as ToolCategory) || 'all'
  );

  // Sync state with URL params when they change
  useEffect(() => {
    const category = searchParams.get('category') || 'all';
    const query = searchParams.get('q') || '';
    setSelectedCategory(category as CategoryFilter);
    setSearchQuery(query);
  }, [searchParams]);
  const [showFilters, setShowFilters] = useState(false);

  // Filter tools based on search and category
  const filteredTools = useMemo(() => {
    let tools = allTools;

    // Filter by category
    if (selectedCategory !== 'all') {
      tools = getToolsByCategory(selectedCategory);
    }

    // Filter by search query
    if (searchQuery.trim()) {
      tools = tools.filter(tool => toolMatchesQuery(tool, searchQuery));
    }

    return tools;
  }, [allTools, selectedCategory, searchQuery]);

  // Category options
  const categories: { value: CategoryFilter; label: string }[] = [
    { value: 'all', label: 'All Tools' },
    { value: 'edit-annotate', label: t('home.categories.editAnnotate') },
    { value: 'convert-to-pdf', label: t('home.categories.convertToPdf') },
    { value: 'convert-from-pdf', label: t('home.categories.convertFromPdf') },
    { value: 'organize-manage', label: t('home.categories.organizeManage') },
    { value: 'optimize-repair', label: t('home.categories.optimizeRepair') },
    { value: 'secure-pdf', label: t('home.categories.securePdf') },
  ];

  const handleClearSearch = useCallback(() => {
    setSearchQuery('');
  }, []);

  const handleClearFilters = useCallback(() => {
    setSearchQuery('');
    setSelectedCategory('all');
  }, []);

  return (
    <div className="min-h-screen flex flex-col bg-[hsl(var(--color-background))]">
      <Header locale={locale} />

      <main className="flex-1">
        {/* Page Header */}
        <section className="relative pt-36 pb-20 overflow-hidden">
          {/* Animated Background Blobs (Subtle) */}
          <div className="absolute top-0 left-0 w-full h-full overflow-hidden pointer-events-none -z-10">
            <div className="absolute top-0 right-0 w-[500px] h-[500px] bg-[hsl(var(--color-primary)/0.05)] rounded-full mix-blend-multiply filter blur-3xl opacity-50" />
            <div className="absolute bottom-0 left-0 w-[400px] h-[400px] bg-[hsl(var(--color-accent)/0.05)] rounded-full mix-blend-multiply filter blur-3xl opacity-50" />
          </div>

          <div className="container mx-auto px-4 relative z-10">
            <div className="max-w-3xl mx-auto text-center">
              <h1 className="text-4xl md:text-5xl font-bold text-[hsl(var(--color-foreground))] mb-6">
                <span className="text-gradient">Professional PDF Tools</span>
              </h1>
              <p className="text-lg text-[hsl(var(--color-muted-foreground))] mb-10 leading-relaxed">
                {allTools.length}+ free, secure, and easy-to-use tools for all your PDF needs.
              </p>

              {/* Search Bar */}
              <div className="relative max-w-2xl mx-auto">
                <div className="relative group">
                  <Search className="absolute left-5 top-1/2 -translate-y-1/2 h-5 w-5 text-[hsl(var(--color-muted-foreground))] group-focus-within:text-[hsl(var(--color-primary))] transition-colors" aria-hidden="true" />
                  <input
                    type="search"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder={t('tools.search.placeholder')}
                    className="w-full pl-14 pr-12 py-4 text-lg rounded-2xl border border-[hsl(var(--color-border))] bg-white/80 dark:bg-slate-900/80 backdrop-blur-sm shadow-sm focus:outline-none focus:ring-4 focus:ring-[hsl(var(--color-primary)/0.1)] focus:border-[hsl(var(--color-primary))] transition-all"
                    aria-label="Search tools"
                  />
                  {searchQuery && (
                    <button
                      onClick={handleClearSearch}
                      className="absolute right-4 top-1/2 -translate-y-1/2 p-2 hover:bg-[hsl(var(--color-muted))] rounded-full transition-colors"
                      aria-label="Clear search"
                    >
                      <X className="h-5 w-5 text-[hsl(var(--color-muted-foreground))]" aria-hidden="true" />
                    </button>
                  )}
                </div>
              </div>
            </div>
          </div>
        </section>

        {/* Filters and Tools */}
        <section className="py-8 bg-[hsl(var(--color-muted)/0.3)] min-h-[500px]">
          <div className="container mx-auto px-4">
            {/* Filter Bar */}
            <div className="flex flex-col md:flex-row items-center gap-6 mb-10 sticky top-20 z-40 py-4 px-6 rounded-2xl glass-card transition-all">
              {/* Mobile Filter Toggle */}
              <Button
                variant="outline"
                className="md:hidden w-full"
                onClick={() => setShowFilters(!showFilters)}
                aria-expanded={showFilters}
                aria-controls="category-filters"
              >
                <Filter className="h-4 w-4 mr-2" aria-hidden="true" />
                Filters
              </Button>

              {/* Category Filters */}
              <div
                className={`flex flex-wrap gap-2 ${showFilters ? 'block w-full' : 'hidden md:flex flex-1'}`}
                role="group"
                aria-label="Filter by category"
              >
                {categories.map((cat) => (
                  <button
                    key={cat.value}
                    onClick={() => setSelectedCategory(cat.value)}
                    aria-pressed={selectedCategory === cat.value}
                    className={`
                      px-4 py-2 rounded-full text-sm font-medium transition-all duration-200
                      ${selectedCategory === cat.value
                        ? 'bg-[hsl(var(--color-primary))] text-white shadow-md shadow-primary/25'
                        : 'bg-transparent text-[hsl(var(--color-muted-foreground))] hover:bg-[hsl(var(--color-muted))] hover:text-[hsl(var(--color-foreground))]'
                      }
                    `}
                  >
                    {cat.label}
                    {cat.value !== 'all' && (
                      <span className={`ml-1 text-xs ${selectedCategory === cat.value ? 'opacity-100' : 'opacity-60'}`}>
                        ({getToolsByCategory(cat.value).length})
                      </span>
                    )}
                  </button>
                ))}
              </div>

              {/* Clear Filters */}
              {(searchQuery || selectedCategory !== 'all') && (
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={handleClearFilters}
                  className="ml-auto text-sm text-[hsl(var(--color-muted-foreground))]"
                >
                  Clear all
                </Button>
              )}
            </div>

            {/* Results Count */}
            <div className="mb-6 px-2">
              <p className="text-sm text-[hsl(var(--color-muted-foreground))]">
                {filteredTools.length === allTools.length
                  ? `Showing all ${allTools.length} tools`
                  : `Showing ${filteredTools.length} of ${allTools.length} tools`}
                {searchQuery && ` for "${searchQuery}"`}
                {selectedCategory !== 'all' && ` in ${t(`home.categories.${categoryTranslationKeys[selectedCategory]}`)}`}
              </p>
            </div>

            {/* Tools Grid */}
            {filteredTools.length > 0 ? (
              selectedCategory === 'all' && !searchQuery ? (
                // Show grouped by category when no filters
                <ToolGrid
                  tools={filteredTools}
                  locale={locale}
                  localizedToolContent={localizedToolContent}
                  showCategoryHeaders
                />
              ) : (
                // Show flat grid when filtered
                <ToolGrid
                  tools={filteredTools}
                  locale={locale}
                  localizedToolContent={localizedToolContent}
                />
              )
            ) : (
              // No results
              <Card className="p-16 text-center glass-card border-dashed border-2">
                <div className="max-w-md mx-auto flex flex-col items-center">
                  <div className="w-20 h-20 bg-[hsl(var(--color-muted))] rounded-full flex items-center justify-center mb-6">
                    <Search className="h-10 w-10 text-[hsl(var(--color-muted-foreground))]" aria-hidden="true" />
                  </div>
                  <h3 className="text-xl font-bold text-[hsl(var(--color-foreground))] mb-2">
                    No tools found
                  </h3>
                  <p className="text-[hsl(var(--color-muted-foreground))] mb-8">
                    {t('tools.search.noResults', { query: searchQuery })}
                  </p>
                  <Button variant="outline" onClick={handleClearFilters} className="px-8">
                    Clear filters
                  </Button>
                </div>
              </Card>
            )}
          </div>
        </section>
      </main>

      <Footer locale={locale} />
    </div>
  );
}
