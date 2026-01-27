import createNextIntlPlugin from 'next-intl/plugin';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const withNextIntl = createNextIntlPlugin('./src/i18n/request.ts');

/** @type {import('next').NextConfig} */
const nextConfig = {
  // Enable static export for deployment flexibility
  output: 'export',

  // Webpack configuration for WASM modules
  webpack: (config, { isServer, webpack }) => {
    // Handle qpdf-wasm and other modules that use Node.js built-ins
    if (!isServer) {
      config.resolve.fallback = {
        ...config.resolve.fallback,
        fs: false,
        path: false,
        crypto: false,
        module: false,
        url: false,
        worker_threads: false,
        canvas: false,
      };
    }

    // Also add module and canvas to alias for some packages that use it
    config.resolve.alias = {
      ...config.resolve.alias,
      'module': false,
      'canvas': path.resolve(__dirname, './src/lib/mocks/canvas.js'),
    };

    // Ignore the dynamic import of 'module' in gs-wasm
    config.plugins.push(
      new webpack.IgnorePlugin({
        resourceRegExp: /^module$/
      })
    );

    // Enable WebAssembly
    config.experiments = {
      ...config.experiments,
      asyncWebAssembly: true,
    };

    return config;
  },

  // Image optimization configuration
  // Note: unoptimized is required for static export
  images: {
    unoptimized: true,
    // Define allowed image formats
    formats: ['image/avif', 'image/webp'],
    // Define device sizes for responsive images
    deviceSizes: [640, 750, 828, 1080, 1200, 1920, 2048, 3840],
    // Define image sizes for srcset
    imageSizes: [16, 32, 48, 64, 96, 128, 256, 384],
    // Minimum cache TTL for optimized images (in seconds)
    minimumCacheTTL: 60 * 60 * 24 * 30, // 30 days
  },

  turbopack: {
    resolveAlias: {
      canvas: './src/lib/mocks/canvas.js',
    },
  },

  // Trailing slash for static hosting compatibility
  trailingSlash: true,

  // Strict mode for better development experience
  reactStrictMode: true,

  // TypeScript configuration
  typescript: {
    // Allow production builds even with type errors during development
    ignoreBuildErrors: false,
  },

  // ESLint configuration
  eslint: {
    // Run ESLint during builds
    ignoreDuringBuilds: false,
  },

  // Compiler options for performance
  compiler: {
    // Remove console.log in production
    removeConsole: process.env.NODE_ENV === 'production' ? {
      exclude: ['error', 'warn'],
    } : false,
  },

  // Headers configuration for caching
  // Note: These headers are applied when running with `next start`
  // For static export, configure headers in your hosting platform
  async headers() {
    return [
      {
        // Static assets - long cache
        source: '/:path*.(ico|jpg|jpeg|png|gif|svg|webp|avif|woff|woff2|ttf|eot)',
        headers: [
          {
            key: 'Cache-Control',
            value: 'public, max-age=31536000, immutable',
          },
        ],
      },
      {
        // JavaScript and CSS - cache with revalidation
        source: '/:path*.(js|css)',
        headers: [
          {
            key: 'Cache-Control',
            value: 'public, max-age=31536000, immutable',
          },
        ],
      },
      {
        // HTML pages - short cache with revalidation
        source: '/:path*',
        headers: [
          {
            key: 'Cache-Control',
            value: 'public, max-age=0, must-revalidate',
          },
        ],
      },
      {
        // Security headers for all routes
        source: '/:path*',
        headers: [
          {
            key: 'X-Content-Type-Options',
            value: 'nosniff',
          },
          {
            key: 'X-Frame-Options',
            value: 'SAMEORIGIN',
          },
          {
            key: 'X-XSS-Protection',
            value: '1; mode=block',
          },
          {
            key: 'Referrer-Policy',
            value: 'strict-origin-when-cross-origin',
          },
        ],
      },
    ];
  },
};

export default withNextIntl(nextConfig);
