# Word to PDF

A responsive, browser-based DOCX-to-PDF converter with a clean lavender and neutral interface.

## Features
- Upload or drag-and-drop a `.docx` file
- Convert in the browser and download a PDF
- 15 MB upload limit
- No conversion server or account required
- Responsive layout and subtle developer credit

## Deploy to Cloudflare Pages
This is a static HTML/CSS/JavaScript site.

- Framework preset: `None`
- Build command: `exit 0`
- Build output directory: `.`
- Root directory: leave blank

Connect the GitHub repository to Cloudflare Pages and deploy.

## Important limitations
- Supports `.docx` only; legacy `.doc` files are not supported.
- Conversion uses browser-side DOCX rendering and HTML-to-PDF libraries loaded from CDNs.
- Complex page layouts, headers/footers, columns, fonts, and page breaks may render differently from Microsoft Word. Review the PDF after conversion.
- The document is processed locally in the browser by this app. Third-party CDN requests load the conversion libraries; no document upload endpoint is used.

## Local testing
Open `index.html` in a modern browser while connected to the internet, or serve the folder with a local static web server.
