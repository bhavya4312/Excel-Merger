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

1. **Pyodide Init** — Loads Python WASM runtime and installs `pandas`, `openpyxl`, `xlrd`.
2. **Upload & Label** — User uploads two `.xls`/`.xlsx` files and labels each as "RAHUL SURGICAL" or "RAHUL MEDICAL & SURGICAL".
3. **Python Processing** — An embedded Python script parses the Excel files, groups transactions by party, and produces a styled merged `.xlsx`.
4. **PDF Generation** — The structured data (returned as JSON from Python) is rendered into a formatted single-page PDF using jsPDF.
5. **Download & Reset** — User downloads the merged Excel and/or PDF, or resets to start over.

For detailed information on the input Excel file format, see [`DATA_FORMAT.md`](DATA_FORMAT.md).

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

## Project Structure

```
Excel-Merger/
├── index.html              # Entry HTML (loads Pyodide CDN script)
├── src/
│   ├── main.jsx            # React entry point
│   ├── App.jsx             # Main application component (all logic)
│   ├── App.css             # Styles (mostly unused Vite defaults)
│   └── index.css           # Tailwind imports
├── public/                 # Static assets
├── test-files/             # Sample input files for testing
│   ├── RMS KORBA.xlsx      # Sample Rahul Medical & Surgicals report
│   └── RS KORBA.xlsx       # Sample Rahul Surgical report
├── vite.config.js          # Vite config (base path for GH Pages)
├── tailwind.config.js      # Tailwind configuration
└── package.json
```
