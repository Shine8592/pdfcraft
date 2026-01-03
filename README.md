# PDFCraft

<div align="center">
  <img src="public/images/logo.png" alt="PDFCraft Logo" width="120" height="120" />
  <h1>Professional PDF Tools</h1>
  <p>
    <strong>Free, Private & Browser-Based</strong>
  </p>
  <p>
    Merge, split, compress, convert, and edit PDF files online without uploading to servers.
  </p>
</div>

<div align="center">

![Next.js](https://img.shields.io/badge/Next.js-15-black?style=flat-square&logo=next.js)
![React](https://img.shields.io/badge/React-19-blue?style=flat-square&logo=react)
![TypeScript](https://img.shields.io/badge/TypeScript-5-blue?style=flat-square&logo=typescript)
![Tailwind CSS](https://img.shields.io/badge/Tailwind-4-38bdf8?style=flat-square&logo=tailwindcss)

</div>

## 📖 About

**PDFCraft** is a comprehensive suite of PDF tools designed for privacy and performance. Unlike many online converters, PDFCraft processes your files entirely within your browser using WebAssembly technology. Your documents **never** leave your device, ensuring maximum security for your sensitive data.

This project is built with modern web technologies to provide a slick, app-like experience directly in the browser.

## ✨ Key Features

- **🔒 100% Private**: All processing happens client-side. No file uploads to external servers.
- **🚀 Fast & Responsive**: Powered by Next.js and WebAssembly for near-native performance.
- **🛠️ Comprehensive Toolset**: Over 60+ tools to handle any PDF task.
- **🎨 Modern UI**: Clean, accessible, and responsive design built with Tailwind CSS.

## 🧰 Available Tools

### Organize & Manage
- **Merge PDF**: Combine multiple PDFs into one.
- **Split PDF**: Separate specific pages or ranges.
- **Organize PDF**: Reorder, delete, or rotate pages.
- **Extract Pages**: Extract specific pages to a new file.

### Edit & Annotate
- **Edit PDF**: Add text, shapes, and annotations.
- **Sign PDF**: Add digital signatures.
- **Watermark**: Apply text or image watermarks.
- **Page Numbers**: Add customizable page numbering.

### Convert to PDF
- **Images to PDF**: Convert JPG, PNG, WebP, HEIC, and more to PDF.
- **Office to PDF**: Convert Word, Excel, PowerPoint to PDF (Client-side).
- **Code/Text to PDF**: Convert TXT, JSON, HTML to PDF.

### Convert from PDF
- **PDF to Images**: Extract pages as JPG, PNG, WebP, or BMP.
- **PDF to Office**: Convert PDF to Word, Excel, or PowerPoint.
- **PDF to JSON/Text**: Extract text and metadata.

### Optimize & Repair
- **Compress PDF**: Reduce file size while maintaining quality.
- **Repair PDF**: Fix corrupted PDF files.
- **Flatten PDF**: Merge annotations and layers.

## 💻 Tech Stack

- **Framework**: [Next.js 15](https://nextjs.org/) (App Router)
- **Language**: [TypeScript](https://www.typescriptlang.org/)
- **Styling**: [Tailwind CSS 4](https://tailwindcss.com/)
- **PDF Processing**:
  - [PDF.js](https://github.com/mozilla/pdf.js)
  - [pdf-lib](https://github.com/Hopding/pdf-lib)
  - [PyMuPDF (WASM)](https://pymupdf.readthedocs.io/)
- **State Management**: [Zustand](https://github.com/pmndrs/zustand)

## 🚀 Getting Started

To run this project locally, follow these steps:

### Prerequisites

- Node.js 18.17 or later
- npm, yarn, or pnpm

### Installation

1.  **Clone the repository**
    ```bash
    git clone https://github.com/yourusername/pdfcraft.git
    cd pdfcraft
    ```

2.  **Install dependencies**
    ```bash
    npm install
    # or
    yarn install
    # or
    pnpm install
    ```

3.  **Start the development server**
    ```bash
    npm run dev
    # or
    yarn dev
    # or
    pnpm dev
    ```

4.  **Open your browser**
    Navigate to [http://localhost:3000](http://localhost:3000) to see the application running.

## 📜 Scripts

- `npm run dev`: Starts the development server with Turbopack.
- `npm run build`: Builds the application for production.
- `npm run start`: Starts the production server.
- `npm run lint`: Lints the code using ESLint.
- `npm run test`: Runs tests using Vitest.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1.  Fork the project
2.  Create your feature branch (`git checkout -b feature/AmazingFeature`)
3.  Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4.  Push to the branch (`git push origin feature/AmazingFeature`)
5.  Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

<div align="center">
  Built with ❤️ by the PDFCraft Team
</div>
