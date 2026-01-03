const fs = require('fs');
const path = require('path');

// Read the original file
const inputPath = path.join(__dirname, '../node_modules/@bentopdf/pymupdf-wasm/dist/index.js');
const outputPath = path.join(__dirname, '../public/pymupdf-wasm/pymupdf-lite.js');

let content = fs.readFileSync(inputPath, 'utf8');

// Remove the gs-wasm import
content = content.replace(
  'import loadGhostscriptWASM from "@bentopdf/gs-wasm";',
  '// Ghostscript import removed for lite version'
);

// Replace the convertPdfToRgb function with a stub
const oldFunctionRegex = /async function convertPdfToRgb\(pdfData\) \{[\s\S]*?return copy;\s*\}/;
const newFunction = `async function convertPdfToRgb(pdfData) {
  // Skip Ghostscript conversion in lite version, return original data
  console.log("[convertPdfToRgb] Lite version - skipping Ghostscript RGB conversion");
  return pdfData;
}`;

content = content.replace(oldFunctionRegex, newFunction);

// Write the modified file
fs.writeFileSync(outputPath, content, 'utf8');

console.log('Created pymupdf-lite.js successfully');
