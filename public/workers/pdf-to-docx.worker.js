/**
 * PDF to DOCX Worker (via Pyodide + pdf2docx)
 */

import { loadPyodide } from '/pymupdf-wasm/pyodide.js';

let pyodide = null;
let initPromise = null;

async function init() {
  if (pyodide) return pyodide;

  self.postMessage({ type: 'status', message: 'Loading Python environment...' });

  // Initialize Pyodide
  pyodide = await loadPyodide({
    indexURL: '/pymupdf-wasm/',
    fullStdLib: false // We use our own wheels mostly
  });

  self.postMessage({ type: 'status', message: 'Installing dependencies...' });

  // We will use pyodide.loadPackage directly instead of micropip
  // because micropip seems to be missing in the local distribution.

  // Helper to install checks
  const install = async (url) => {
    // self.postMessage({ type: 'status', message: `Installing ${url.split('/').pop()}...` });
    await pyodide.loadPackage(url);
  };

  // Install wheels in order
  const basePath = '/pymupdf-wasm/';

  // We need to install dependencies.
  // Note: some packages might be able to be loaded from pyodide standard lib if available (like numpy), 
  // but we have wheels provided in the folder, best to use them to match versions if possible.
  // However, `numpy` wheel is large.

  // Install Core Wheels
  // Order matters for some

  // Mock missing non-critical dependencies
  pyodide.runPython(`
    import sys
    from types import ModuleType
    
    # Mock tqdm (used for progress bars)
    tqdm_mod = ModuleType("tqdm")
    def tqdm(iterable=None, *args, **kwargs):
        return iterable if iterable else []
    tqdm_mod.tqdm = tqdm
    sys.modules["tqdm"] = tqdm_mod
    
    # Mock fire (CLI tool, not needed for library usage)
    fire_mod = ModuleType("fire")
    sys.modules["fire"] = fire_mod
  `);

  await install(basePath + 'numpy-2.2.5-cp313-cp313-pyodide_2025_0_wasm32.whl');
  await install(basePath + 'typing_extensions-4.12.2-py3-none-any.whl');
  // Packaging is missing locally but required by fonttools. Try fetching from CDN.
  try {
    await install('https://cdn.jsdelivr.net/pyodide/v0.26.1/full/packaging-24.1-py3-none-any.whl');
  } catch (e) {
    console.warn("Failed to load packaging from CDN, fonttools might fail:", e);
  }
  await install(basePath + 'fonttools-4.56.0-py3-none-any.whl');
  await install(basePath + 'lxml-5.4.0-cp313-cp313-pyodide_2025_0_wasm32.whl');
  await install(basePath + 'pymupdf-1.26.3-cp313-none-pyodide_2025_0_wasm32.whl');
  await install(basePath + 'python_docx-1.2.0-py3-none-any.whl');
  // opencv is huge, only install if pdf2docx strictly requires it (it usually does for image extraction)
  await install(basePath + 'opencv_python-4.11.0.86-cp313-cp313-pyodide_2025_0_wasm32.whl');

  // Finally pdf2docx
  self.postMessage({ type: 'status', message: 'Installing pdf2docx...' });
  await install(basePath + 'pdf2docx-0.5.8-py3-none-any.whl');

  // Define the python processing script
  self.postMessage({ type: 'status', message: 'Initializing converter script...' });

  pyodide.runPython(`
import os
import sys
from pdf2docx import Converter

def convert_pdf_to_docx(input_obj):
    # Convert JsProxy (Uint8Array) to bytes-like object
    # to_py() converts JS TypedArray to Python memoryview
    if hasattr(input_obj, "to_py"):
        input_bytes = input_obj.to_py()
    else:
        input_bytes = input_obj

    # Write input PDF
    with open("input.pdf", "wb") as f:
        f.write(input_bytes)
        
    # Convert
    cv = Converter("input.pdf")
    # Convert to docx
    # start=0, end=None means all pages
    cv.convert("output.docx", start=0, end=None)
    cv.close()
    
    # Read output
    with open("output.docx", "rb") as f:
        docx_bytes = f.read()
        
    # Cleanup
    if os.path.exists("input.pdf"):
        os.remove("input.pdf")
    if os.path.exists("output.docx"):
        os.remove("output.docx")
        
    return docx_bytes
  `);

  return pyodide;
}

self.onmessage = async (event) => {
  const { type, id, data } = event.data;

  try {
    if (type === 'init') {
      if (!initPromise) initPromise = init();
      await initPromise;
      self.postMessage({ id, type: 'init-complete' });
      return;
    }

    if (type === 'convert') {
      if (!pyodide) {
        if (!initPromise) initPromise = init();
        await initPromise;
      }

      const { file } = data; // File object
      const arrayBuffer = await file.arrayBuffer();
      const inputBytes = new Uint8Array(arrayBuffer);

      self.postMessage({ type: 'status', message: 'Processing PDF...' });

      // Call Python function
      const convertFunc = pyodide.globals.get('convert_pdf_to_docx');

      // Convert takes bytes, returns bytes (as PyProxy or Uint8Array)
      // We pass the TypedArray directly which Pyodide handles as bytes
      const resultProxy = convertFunc(inputBytes);
      const resultBytes = resultProxy.toJs();
      resultProxy.destroy();

      const resultBlob = new Blob([resultBytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

      self.postMessage({
        id,
        type: 'convert-complete',
        result: resultBlob
      });
    }

  } catch (error) {
    console.error('Worker error:', error);
    self.postMessage({
      id,
      type: 'error',
      error: error.message || String(error)
    });
  }
};
