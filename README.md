# Excel Merge Tool

A client-side web application that merges **Rahul Medical & Surgicals** and **Rahul Surgical** outstanding reports into a single combined Excel and PDF output — grouped by party (customer).

🔗 **Live:** [https://bhavya4312.github.io/Excel-Merger](https://bhavya4312.github.io/Excel-Merger)

## What It Does

Takes two Excel files — one from each business entity — and:

1. **Parses** each file to extract party-wise outstanding transactions.
2. **Matches** parties across both files by name + city.
3. **Merges** transactions under each party, split into `--- Medical ---` and `--- Surgical ---` sections.
4. **Generates** a styled Excel file (`.xlsx`) with color-coded sections and an auto-formatted single-page PDF.

## Tech Stack

| Layer | Technology |
|---|---|
| **UI** | React 19, Tailwind CSS, Lucide Icons |
| **Build** | Vite 7 |
| **Excel Processing** | [Pyodide](https://pyodide.org/) (Python in WebAssembly) with `pandas`, `openpyxl`, `xlrd` |
| **PDF Generation** | [jsPDF](https://github.com/parallax/jsPDF) + [jspdf-autotable](https://github.com/simonbengtsson/jsPDF-AutoTable) |
| **Hosting** | GitHub Pages via `gh-pages` |

> **No backend required** — all processing happens entirely in the browser.

## Architecture

The app is a single React component ([`src/App.jsx`](src/App.jsx)) with 5 stages:

```
┌──────────────┐     ┌──────────────┐     ┌─────────────────┐     ┌──────────────┐     ┌──────────────┐
│  1. Init     │────▶│  2. Upload   │────▶│  3. Python      │────▶│  4. PDF Gen  │────▶│  5. Download │
│  Pyodide     │     │  & Label     │     │  Processing     │     │  (jsPDF)     │     │  & Reset     │
└──────────────┘     └──────────────┘     └─────────────────┘     └──────────────┘     └──────────────┘
```

1. **Pyodide Init** — Loads Python WASM runtime and installs `pandas`, `openpyxl`, `xlrd`. Shows progress bar with retry on failure.
2. **Upload & Label** — User uploads two `.xls`/`.xlsx` files via click or drag-and-drop. Labels are auto-detected from the file's header row.
3. **Python Processing** — An embedded Python script parses the Excel files, groups transactions by party, and produces a styled merged `.xlsx`.
4. **PDF Generation** — The structured data (returned as JSON from Python) is rendered into a formatted single-page PDF using jsPDF.
5. **Preview & Download** — Shows a merge summary (Medical/Surgical/Grand totals, party list) and offers Excel and PDF downloads.

For detailed information on the input Excel file format, see [`DATA_FORMAT.md`](DATA_FORMAT.md).

## Features

- **Auto-detect file labels** — Automatically identifies whether an uploaded file is a Medical or Surgical report by reading Row 1.
- **Drag-and-drop upload** — Drop Excel files directly onto the upload zones with visual feedback.
- **File validation** — Validates file type (.xls/.xlsx) and size on upload, before processing.
- **Merge preview** — After processing, shows a summary with Medical/Surgical/Grand totals and a scrollable party list.
- **Indian number formatting** — All monetary amounts displayed with Indian comma grouping (e.g., ₹12,34,567.00).
- **User-friendly errors** — Technical Python errors are translated into plain English messages.
- **Progress stepper** — Visual 3-step indicator (Upload → Process → Results) guides the user.
- **Loading progress bar** — Multi-step progress indicator during Pyodide initialization with retry support.
- **Date-stamped filenames** — Output files are named `Merged_Report_YYYY-MM-DD.xlsx` / `.pdf`.
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
