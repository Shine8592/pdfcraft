/**
 * PyMuPDF Loader
 * Dynamically loads PyMuPDF WASM module using script tag
 */

// Singleton instance
let pymupdfInstance: any = null;
let loadingPromise: Promise<any> | null = null;

// Declare global type
declare global {
  interface Window {
    loadPyodide?: (config: any) => Promise<any>;
  }
}

/**
 * Load PyMuPDF using Pyodide directly
 * This avoids the ES module import issues with @bentopdf/pymupdf-wasm
 */
export async function loadPyMuPDF(): Promise<any> {
  if (pymupdfInstance) {
    return pymupdfInstance;
  }

  if (loadingPromise) {
    return loadingPromise;
  }

  loadingPromise = (async () => {
    try {
      const basePath = `${window.location.origin}/pymupdf-wasm/`;

      // Load Pyodide script
      await loadScript(`${basePath}pyodide.js`);

      // Initialize Pyodide
      const pyodide = await window.loadPyodide!({
        indexURL: basePath,
      });

      // Load required packages
      await pyodide.loadPackage('micropip');
      const micropip = pyodide.pyimport('micropip');

      // Install pymupdf and pdf2docx from local wheels
      await micropip.install(`${basePath}pymupdf-1.26.3-cp313-none-pyodide_2025_0_wasm32.whl`);
      await micropip.install(`${basePath}fonttools-4.56.0-py3-none-any.whl`);
      await micropip.install(`${basePath}lxml-5.4.0-cp313-cp313-pyodide_2025_0_wasm32.whl`);
      await micropip.install(`${basePath}numpy-2.2.5-cp313-cp313-pyodide_2025_0_wasm32.whl`);
      await micropip.install(`${basePath}python_docx-1.2.0-py3-none-any.whl`);
      await micropip.install(`${basePath}typing_extensions-4.12.2-py3-none-any.whl`);
      await micropip.install(`${basePath}pdf2docx-0.5.8-py3-none-any.whl`);

      // Import pymupdf
      await pyodide.runPythonAsync('import pymupdf');

      // Create a wrapper object with pdfToDocx method
      pymupdfInstance = {
        pyodide,
        async pdfToDocx(file: File): Promise<Blob> {
          const arrayBuffer = await file.arrayBuffer();
          const pdfData = new Uint8Array(arrayBuffer);

          // Write PDF to virtual filesystem
          pyodide.FS.writeFile('/input.pdf', pdfData);

          // Convert using pdf2docx
          const result = await pyodide.runPythonAsync(`
import base64
from pdf2docx import Converter

cv = Converter('/input.pdf')
cv.convert('/output.docx')
cv.close()

with open('/output.docx', 'rb') as f:
    docx_data = f.read()

base64.b64encode(docx_data).decode('ascii')
`);

          // Clean up
          try {
            pyodide.FS.unlink('/input.pdf');
            pyodide.FS.unlink('/output.docx');
          } catch {
            // Ignore cleanup errors
          }

          // Convert base64 to Blob
          const binary = atob(result);
          const bytes = new Uint8Array(binary.length);
          for (let i = 0; i < binary.length; i++) {
            bytes[i] = binary.charCodeAt(i);
          }

          return new Blob([bytes], {
            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          });
        },
      };

      return pymupdfInstance;
    } catch (error) {
      loadingPromise = null;
      throw error;
    }
  })();

  return loadingPromise;
}

/**
 * Load a script dynamically
 */
function loadScript(src: string): Promise<void> {
  return new Promise((resolve, reject) => {
    // Check if already loaded
    if (document.querySelector(`script[src="${src}"]`)) {
      resolve();
      return;
    }

    const script = document.createElement('script');
    script.src = src;
    script.async = true;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error(`Failed to load script: ${src}`));
    document.head.appendChild(script);
  });
}

/**
 * Reset the loader (for testing)
 */
export function resetPyMuPDF(): void {
  pymupdfInstance = null;
  loadingPromise = null;
}
