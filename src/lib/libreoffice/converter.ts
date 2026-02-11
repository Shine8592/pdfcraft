/**
 * LibreOffice WASM Converter
 * 
 * Uses @matbee/libreoffice-converter for document conversion.
 * 
 * Uses BrowserConverter (main-thread mode) instead of WorkerBrowserConverter
 * because the worker mode has a hardcoded 10-second timeout for the worker
 * to load and send a "loaded" message. On slower connections or CDN-proxied
 * deployments, the large browser.worker.global.js file often fails to load
 * within this window, causing "Worker load timeout" errors.
 * 
 * BrowserConverter loads directly on the main thread with a 60-second timeout,
 * which is more reliable for production deployments.
 */

import { BrowserConverter } from '@matbee/libreoffice-converter/browser';

const LIBREOFFICE_PATH = '/libreoffice-wasm/';

/**
 * CJK font files to inject into LibreOffice WASM virtual filesystem.
 * These are fetched from /fonts/ and written to /instdir/share/fonts/truetype/
 * so LibreOffice can render CJK characters correctly.
 */
const CJK_FONTS = [
    { url: '/fonts/NotoSansSC-Regular.ttf', filename: 'NotoSansSC-Regular.ttf' },
];

export interface LoadProgress {
    phase: 'loading' | 'initializing' | 'converting' | 'complete' | 'ready';
    percent: number;
    message: string;
}

export type ProgressCallback = (progress: LoadProgress) => void;

let converterInstance: LibreOfficeConverter | null = null;

export class LibreOfficeConverter {
    private converter: BrowserConverter | null = null;
    private initialized = false;
    private initializing = false;
    private basePath: string;
    private fontsInstalled = false;

    constructor(basePath?: string) {
        this.basePath = basePath || LIBREOFFICE_PATH;
    }

    async initialize(onProgress?: ProgressCallback): Promise<void> {
        if (this.initialized) return;

        if (this.initializing) {
            while (this.initializing) {
                await new Promise(r => setTimeout(r, 100));
            }
            return;
        }

        this.initializing = true;
        let progressCallback = onProgress;

        try {
            progressCallback?.({ phase: 'loading', percent: 0, message: 'Loading conversion engine...' });

            this.converter = new BrowserConverter({
                sofficeJs: `${this.basePath}soffice.js`,
                sofficeWasm: `${this.basePath}soffice.wasm`,
                sofficeData: `${this.basePath}soffice.data`,
                sofficeWorkerJs: `${this.basePath}soffice.worker.js`,
                verbose: false,
                onProgress: (info: { phase: string; percent: number; message: string }) => {
                    if (progressCallback && !this.initialized) {
                        progressCallback({
                            phase: info.phase as LoadProgress['phase'],
                            percent: Math.min(info.percent, 90),
                            message: `Loading conversion engine (${Math.round(info.percent)}%)...`
                        });
                    }
                },
                onReady: () => {
                    console.log('[LibreOffice] WASM ready');
                },
                onError: (error: Error) => {
                    console.error('[LibreOffice] Error:', error);
                },
            });

            await this.converter.initialize();

            // Install CJK fonts into the WASM virtual filesystem
            progressCallback?.({ phase: 'initializing', percent: 92, message: 'Installing CJK fonts...' });
            await this.installCJKFonts();

            this.initialized = true;
            progressCallback?.({ phase: 'ready', percent: 100, message: 'Conversion engine ready!' });
            progressCallback = undefined;
        } finally {
            this.initializing = false;
        }
    }

    /**
     * Install CJK fonts into LibreOffice WASM virtual filesystem.
     * This is necessary because the default soffice.data doesn't include
     * CJK fonts, causing Chinese/Japanese/Korean characters to render
     * as garbled text or empty boxes in converted documents.
     */
    private async installCJKFonts(): Promise<void> {
        if (this.fontsInstalled) return;

        // Access the Emscripten module's virtual filesystem
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const module = (this.converter as any)?.module;
        if (!module?.FS) {
            console.warn('[LibreOffice] Cannot access WASM FS, CJK fonts not installed');
            return;
        }

        const FS = module.FS;

        // Ensure the font directories exist
        const fontDirs = [
            '/instdir/share/fonts',
            '/instdir/share/fonts/truetype',
        ];
        for (const dir of fontDirs) {
            try { FS.mkdir(dir); } catch { /* directory may already exist */ }
        }

        // Fetch and install each CJK font
        for (const font of CJK_FONTS) {
            try {
                console.log(`[LibreOffice] Downloading CJK font: ${font.filename}...`);
                const response = await fetch(font.url);
                if (!response.ok) {
                    console.warn(`[LibreOffice] Failed to fetch font ${font.url}: ${response.status}`);
                    continue;
                }
                const fontBuffer = await response.arrayBuffer();
                const fontData = new Uint8Array(fontBuffer);

                const fontPath = `/instdir/share/fonts/truetype/${font.filename}`;
                FS.writeFile(fontPath, fontData);
                console.log(`[LibreOffice] Installed CJK font: ${fontPath} (${(fontData.length / 1024 / 1024).toFixed(1)}MB)`);
            } catch (err) {
                console.warn(`[LibreOffice] Failed to install font ${font.filename}:`, err);
            }
        }

        this.fontsInstalled = true;
    }

    isReady(): boolean {
        return this.initialized && this.converter !== null;
    }

    async convert(file: File, outputFormat: string): Promise<Blob> {
        if (!this.converter) {
            throw new Error('Converter not initialized');
        }

        const arrayBuffer = await file.arrayBuffer();
        const uint8Array = new Uint8Array(arrayBuffer);
        const ext = file.name.split('.').pop()?.toLowerCase() || '';

        const result = await this.converter.convert(uint8Array, {
            outputFormat: outputFormat as any,
            inputFormat: ext as any,
        }, file.name);

        const data = new Uint8Array(result.data);
        return new Blob([data], { type: result.mimeType });
    }

    async convertToPdf(file: File): Promise<Blob> {
        return this.convert(file, 'pdf');
    }

    async destroy(): Promise<void> {
        if (this.converter) {
            await this.converter.destroy();
        }
        this.converter = null;
        this.initialized = false;
    }
}

export function getLibreOfficeConverter(basePath?: string): LibreOfficeConverter {
    if (!converterInstance) {
        converterInstance = new LibreOfficeConverter(basePath);
    }
    return converterInstance;
}
