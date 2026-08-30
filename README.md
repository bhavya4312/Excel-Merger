# Excel Merge Tool

A client-side web application that merges **Rahul Medical & Surgicals** and **Rahul Surgical** outstanding reports into a single combined Excel and PDF output — grouped by party (customer).

🔗 **Live:** [https://bhavya4312.github.io/Excel-Merger](https://bhavya4312.github.io/Excel-Merger)

## What It Does

Takes two Excel files — one from each business entity — and:

1. **Parses** each file to extract party-wise outstanding transactions.
2. **Matches** parties across both files by name + city.
3. **Merges** transactions under each party, split into `--- Medical ---` and `--- Surgical ---` sections.
4. **Generates** a styled Excel file (`.xlsx`), an auto-formatted single-page PDF (`.pdf`), and a **ZIP archive of individual high-DPI party image cards (`.png`)** ready to share on WhatsApp or mobile.

## Tech Stack

| Layer | Technology |
|---|---|
| **UI** | React 19, Tailwind CSS, Lucide Icons |
| **Build** | Vite 7 |
| **Excel Processing** | [Pyodide](https://pyodide.org/) (Python WASM) with `pandas`, `openpyxl`, `xlrd` in a **Dedicated Web Worker** |
| **PDF Generation** | [jsPDF](https://github.com/parallax/jsPDF) + [jspdf-autotable](https://github.com/simonbengtsson/jsPDF-AutoTable) (Dynamically Imported) |
| **ZIP Packaging** | [JSZip](https://stuk.github.io/jszip/) (Dynamically Imported) |
| **Hosting** | GitHub Pages via `gh-pages` |

> **100% Client-Side & Non-Blocking** — Python WebAssembly executes off the main thread in a dedicated Web Worker with 0ms Total Blocking Time.

## ⚡ Performance & Quality Benchmarks (Lighthouse)

| Category | Score | Metric |
|---|:---:|---|
| **Accessibility** | **100 / 100** | WCAG AA color contrast, ARIA landmarks, keyboard navigation |
| **Best Practices** | **100 / 100** | Modern Web Standards, PWA manifest, HTTPS readiness |
| **SEO** | **100 / 100** | Optimized metadata, OpenGraph tags, crawlable structure |
| **Performance** | **97 / 100** | Web Worker offloading, code splitting, 0ms Total Blocking Time |

## Architecture

The app is built around a non-blocking Web Worker pipeline ([`src/pyodideWorker.js`](src/pyodideWorker.js)):

```
┌─────────────────────────────────┐           ┌─────────────────────────────────┐
│        Main UI Thread           │           │     Background Web Worker       │
│  (React 19 + Tailwind CSS)      │           │   (Pyodide WASM + Pandas/Excel) │
├─────────────────────────────────┤           ├─────────────────────────────────┤
│  • Instant 0ms TBT rendering    │           │  • Download & compile Pyodide   │
│  • Drag-and-drop file upload    │  postMsg  │  • Fast row-1 label detector    │
│  • Single-page PDF & image cards│ ────────▶ │  • Python merge engine          │
│  • Multi-channel ZIP packaging  │ ◀──────── │  • In-memory workbook styling   │
└─────────────────────────────────┘           └─────────────────────────────────┘
```

1. **Worker Init** — Loads Python WASM runtime off-main-thread and installs `pandas`, `openpyxl`, `xlrd`.
2. **Upload & Label** — User uploads two `.xls`/`.xlsx` files. Web Worker inspects header rows concurrently without blocking the UI.
3. **Python Processing** — Worker runs the merge pipeline, groups transactions by party, and streams progress logs.
4. **PDF & Image Generation** — On-demand dynamic imports render formatted single-page PDFs and 2x high-DPI statement cards.
5. **Preview & Export** — Download styled Excel, PDF, or complete **Party Images (ZIP)**.

For detailed information on the input Excel file format, see [`DATA_FORMAT.md`](DATA_FORMAT.md).

## Features

- **Individual Party Image Cards (ZIP)** — Generates high-DPI (2x retina) PNG statement cards for each customer, with company branding, transaction tables, and summary footer, packaged in a single ZIP (`Party_Wise_Reports_YYYY-MM-DD.zip`) for easy sharing via WhatsApp or mobile.
- **Auto-detect file labels** — Automatically identifies whether an uploaded file is a Medical or Surgical report by reading Row 1.
- **Drag-and-drop upload** — Drop Excel files directly onto the upload zones with visual feedback.
- **File validation** — Validates file type (.xls/.xlsx) and size on upload, before processing.
- **Merge preview** — After processing, shows a summary with Medical/Surgical/Grand totals and a scrollable party list.
- **Indian number formatting** — All monetary amounts displayed with Indian comma grouping (e.g., ₹12,34,567.00).
- **User-friendly errors** — Technical Python errors are translated into plain English messages.
- **Progress stepper** — Visual 3-step indicator (Upload → Process → Results) guides the user.
- **Loading progress bar** — Multi-step progress indicator during Pyodide initialization with retry support.
- **Date-stamped filenames** — Output files are named `Merged_Report_YYYY-MM-DD.xlsx` / `.pdf` and `Party_Wise_Reports_YYYY-MM-DD.zip`.
- **Styled Excel output** — Color-coded sections, Indian number format on cells, sheet named "Outstanding Report".
- **PDF with title** — Single-page PDF with "Outstanding Report — DD Mon YYYY" title.

## Getting Started

### Prerequisites

- Node.js (v18+)
- npm

### Install & Run

```bash
# Install dependencies
npm install

# Start dev server
npm run dev
```

The app will be available at `http://localhost:5173/Excel-Merger/`.

### Build & Deploy

```bash
# Build for production
npm run build

# Deploy to GitHub Pages
npm run deploy
```

### Testing

After making changes, run the audit to verify no transactions are lost:

```bash
npm run audit
```

This cross-references every invoice, receipt, and party total from the input files (`test-files/RMS KORBA.xlsx` and `RS KORBA.xlsx`) against the merged output (`test-files/Merged_Report_*.xlsx`). Requires `python3` and `openpyxl` (`pip3 install openpyxl`).

**Steps:** Run the app → merge the test files → save output to `test-files/` → run `npm run audit`.

## Project Structure

```
Excel-Merger/
├── index.html              # Entry HTML (loads Pyodide CDN script)
├── src/
│   ├── main.jsx            # React entry point
│   ├── App.jsx             # Main application component (all logic)
│   ├── App.css             # App-level styles
│   └── index.css           # Tailwind imports
├── scripts/
│   └── audit.py            # Transaction audit script (npm run audit)
├── test-files/             # Sample input files for testing
│   ├── RMS KORBA.xlsx      # Sample Rahul Medical & Surgicals report
│   └── RS KORBA.xlsx       # Sample Rahul Surgical report
├── DATA_FORMAT.md          # Detailed input/output format reference
├── vite.config.js          # Vite config (base path for GH Pages)
├── tailwind.config.js      # Tailwind configuration
└── package.json
```
