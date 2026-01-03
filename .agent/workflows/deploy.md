---
description: 如何在生产环境中构建和部署 PDFCraft
---

# 生产环境部署指南

PDFCraft 项目配置为静态导出 (`output: 'export'`)，这意味着它可以部署到任何支持静态网站托管的服务上，而无需 Node.js 服务器。

## 1. 构建项目

在部署之前，通过以下命令构建项目生成静态文件：

```bash
npm run build
```

构建完成后，所有静态文件将位于 `out` 目录中。

*注意：`npm run build` 会自动忽略 `bentopdf-main` 目录（已在 `tsconfig.json` 中配置）。*

## 2. 部署选项

你可以将 `out` 目录中的内容部署到以下任意平台：

### A. Vercel (推荐)
1. 安装 Vercel CLI: `npm i -g vercel`
2. 运行 `vercel` 命令。
3. 按照提示操作，设置构建命令为 `npm run build`，输出目录为 `out`。
4. 或者直接连接你的 GitHub 仓库，Vercel 会自动检测 Next.js 并进行配置。

### B. Nginx / Apache
将 `out` 目录的内容复制到你的 web 服务器的根目录。

**Nginx 配置示例:**
```nginx
server {
    listen 80;
    server_name example.com;
    root /path/to/your/pdfcraft/out;
    index index.html;

    # 处理静态文件
    location / {
        try_files $uri $uri.html $uri/ =404;
    }

    # 开启 gzip 压缩
    gzip on;
    gzip_types text/plain text/css application/json application/javascript text/xml application/xml application/xml+rss text/javascript;
}
```

### C. GitHub Pages
1. 将 `out` 目录推送到你的仓库的 `gh-pages` 分支。
2. 在仓库设置中开启 GitHub Pages。
3. 注意：如果不用自定义域名，可能需要修改 `next.config.js` 中的 `basePath`。

### D. Netlify
1. 将仓库连接到 Netlify。
2. 设置构建命令: `npm run build`
3. 设置发布目录: `out`

## 3. 注意事项
- **Headers 配置**: `next.config.js` 中的 `headers` 配置在静态导出模式下不会自动生效。你需要根据所选的托管平台（如 Vercel 的 `vercel.json` 或 Nginx 配置）单独配置 HTTP headers。
- **图片优化**: 由于静态导出不支持 Next.js 的默认图片优化服务器，项目已配置 `images: { unoptimized: true }`。

## 4. 验证部署
部署后，请检查以下功能以确保一切正常：
- 多语言路由 (如 `/en`, `/zh`)
- 工具页面加载
- WebAssembly (PDF 处理) 功能是否正常工作
