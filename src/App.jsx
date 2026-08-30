import React, { useState, useEffect, useRef, useMemo } from 'react';
import { Upload, FileSpreadsheet, RefreshCw, AlertCircle, CheckCircle, ChevronDown, ChevronUp, Terminal, RotateCcw, FileText, AlertTriangle } from 'lucide-react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';

// ============================================================================
// Constants (outside component — no re-creation on render)
// ============================================================================
const OPTIONS = ["RAHUL SURGICAL", "RAHUL MEDICAL & SURGICAL"];
const MAX_FILE_SIZE_MB = 50;
const VALID_EXTENSIONS = ['xls', 'xlsx'];

const LOADING_STEPS = [
  "Loading Python Runtime...",
  "Setting up package manager...",
  "Installing Python libraries (pandas, openpyxl, xlrd)...",
  "Ready!",
];

// ============================================================================
// Utility Functions
// ============================================================================

/** Format a number with Indian comma grouping (1,23,456.00) */
const formatIndianNumber = (num) => {
  if (num === '' || num === null || num === undefined || isNaN(num)) return '';
  return new Intl.NumberFormat('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(Number(num));
};

/** Get today's date as YYYY-MM-DD */
const getDateString = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
};

/** Get a human-readable date for display */
const getDisplayDate = () => {
  return new Date().toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
};

/** Translate raw Python/JS errors into user-friendly messages */
const classifyError = (errorMessage) => {
  const msg = (errorMessage || '').toLowerCase();
  if (msg.includes('no module') || msg.includes('import'))
    return "A required Python library failed to load. Please refresh the page and try again.";
  if (msg.includes('keyerror') || msg.includes('indexerror') || msg.includes('column'))
    return "The uploaded file doesn't match the expected MARG ERP Outstanding Report format. Please check your files.";
  if (msg.includes('file') && (msg.includes('not found') || msg.includes('no such')))
    return "One of the uploaded files could not be read. Please try re-uploading.";
  if (msg.includes('memory') || msg.includes('heap'))
    return "The files are too large to process in the browser. Try with smaller files.";
  if (msg.includes('no matching parties') || msg.includes('empty'))
    return "No data was found in the uploaded files. Please check that you're uploading MARG ERP Outstanding Reports.";
  return "Something went wrong during processing. Check the System Logs below for details.";
};

/** Validate a file before processing */
const validateFile = (file) => {
  if (!file) return "No file selected.";
  const ext = file.name.split('.').pop().toLowerCase();
  if (!VALID_EXTENSIONS.includes(ext))
    return `Invalid file type (.${ext}). Only .xls and .xlsx files are supported.`;
  if (file.size === 0) return "The file is empty (0 bytes).";
  if (file.size > MAX_FILE_SIZE_MB * 1024 * 1024)
    return `File too large (${(file.size / (1024 * 1024)).toFixed(1)}MB). Maximum is ${MAX_FILE_SIZE_MB}MB.`;
  return null;
};

/** Auto-detect the file label by reading Row 1 via Pyodide */
const autoDetectLabel = async (py, file) => {
  const ext = file.name.split('.').pop().toLowerCase();
  const tempName = `_detect_${Date.now()}.${ext}`;
  const buffer = await file.arrayBuffer();
  py.FS.writeFile(tempName, new Uint8Array(buffer));
  try {
    const engine = ext === 'xls' ? 'xlrd' : 'openpyxl';
    const result = await py.runPythonAsync(`
import pandas as pd
_df = pd.read_excel("${tempName}", header=None, nrows=1, engine='${engine}')
str(_df.iloc[0, 0]).strip()
`);
    if (result.includes('RAHUL MEDICAL')) return 'RAHUL MEDICAL & SURGICAL';
    if (result.includes('RAHUL SURGICAL')) return 'RAHUL SURGICAL';
    return null;
  } finally {
    try { py.FS.unlink(tempName); } catch { /* ignore cleanup errors */ }
  }
};

/** Extract summary info from processed data for the preview panel */
const extractPreviewInfo = (data) => {
  const parties = [];
  let medicalTotal = 0;
  let surgicalTotal = 0;

  data.forEach(row => {
    const valA = String(row[0] || '').trim();
    const valB = String(row[1] || '').trim();

    // Party header detection (same logic as PDF generator)
    const isPartyHeader = (valA && !valB &&
      !valA.toLowerCase().startsWith("rcpt") &&
      !valA.startsWith("Outstanding Summary") &&
      !valA.startsWith("---") &&
      valA !== "Type");

    if (isPartyHeader) {
      parties.push({ name: valA, total: Number(row[6]) || 0 });
    }

    // Parse the Outstanding Summary line for per-entity totals
    if (valA.startsWith("Outstanding Summary")) {
      const medMatch = valA.match(/Medical:\s*([\d,.]+)/);
      const surgMatch = valA.match(/Surgical:\s*([\d,.]+)/);
      if (medMatch) medicalTotal += parseFloat(medMatch[1].replace(/,/g, ''));
      if (surgMatch) surgicalTotal += parseFloat(surgMatch[1].replace(/,/g, ''));
    }
  });

  const grandTotal = parties.reduce((sum, p) => sum + p.total, 0);
  return { parties, medicalTotal, surgicalTotal, grandTotal };
};


// ============================================================================
// App Component
// ============================================================================
const App = () => {
  // --- Application State ---
  const [pyodide, setPyodide] = useState(null);
  const [loading, setLoading] = useState(true);
  const [loadingStep, setLoadingStep] = useState(0);
  const [initError, setInitError] = useState(null);
  const [logs, setLogs] = useState([]);
  const [showLogs, setShowLogs] = useState(false);

  const pyodideInitializing = useRef(false);

  // --- File State ---
  const [files, setFiles] = useState({
    file1: { file: null, label: null, name: null },
    file2: { file: null, label: null, name: null }
  });
  const [fileErrors, setFileErrors] = useState({ file1: null, file2: null });
  const [dragActive, setDragActive] = useState({ file1: false, file2: false });

  // --- Processing State ---
  const [processedFileUrl, setProcessedFileUrl] = useState(null);
  const [processedPdfUrl, setProcessedPdfUrl] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState(null);
  const [warningCount, setWarningCount] = useState(0);

  // --- Preview State ---
  const [previewData, setPreviewData] = useState(null);
  const [showPreview, setShowPreview] = useState(true);

  // --- Computed Values ---
  const currentStep = useMemo(() => {
    if (processedFileUrl) return 3;
    if (files.file1.file && files.file2.file && files.file1.label && files.file2.label && files.file1.label !== files.file2.label) return 2;
    return 1;
  }, [files, processedFileUrl]);

  // ----------------------------------------------------------------------
  // 1. Initialize Pyodide (with progress steps and retry)
  // ----------------------------------------------------------------------
  const doInitPyodide = async () => {
    try {
      setInitError(null);
      setLoading(true);
      setLoadingStep(0);

      if (!window.loadPyodide) {
        throw new Error("Pyodide script not found. Please check your internet connection and refresh.");
      }

      const py = await window.loadPyodide();
      setLoadingStep(1);

      await py.loadPackage("micropip");
      const micropip = py.pyimport("micropip");
      setLoadingStep(2);

      await py.loadPackage("pandas");
      await micropip.install("xlrd");
      await micropip.install("openpyxl");
      setLoadingStep(3);

      setPyodide(py);
      setLoading(false);
      addLog("Python environment ready. Ready to process files.");
    } catch (err) {
      setInitError(`Failed to load Python environment: ${err.message}`);
      setLoading(false);
      pyodideInitializing.current = false;
    }
  };

  useEffect(() => {
    if (pyodideInitializing.current) return;
    pyodideInitializing.current = true;
    doInitPyodide();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const retryInit = () => {
    pyodideInitializing.current = true;
    doInitPyodide();
  };

  const addLog = (msg) => setLogs(prev => [...prev, `[${new Date().toLocaleTimeString()}] ${msg}`]);

  // ----------------------------------------------------------------------
  // 2. File Handling — validation, auto-detect, drag-and-drop
  // ----------------------------------------------------------------------
  const processUploadedFile = async (uploadedFile, fileKey) => {
    // Validate
    const validationError = validateFile(uploadedFile);
    if (validationError) {
      setFileErrors(prev => ({ ...prev, [fileKey]: validationError }));
      return;
    }
    setFileErrors(prev => ({ ...prev, [fileKey]: null }));

    // Clear previous results
    if (processedFileUrl) URL.revokeObjectURL(processedFileUrl);
    if (processedPdfUrl) URL.revokeObjectURL(processedPdfUrl);
    setProcessedFileUrl(null);
    setProcessedPdfUrl(null);
    setPreviewData(null);
    setError(null);
    setWarningCount(0);

    // Set file initially (without label auto-detect)
    setFiles(prev => ({
      ...prev,
      [fileKey]: { ...prev[fileKey], file: uploadedFile, name: uploadedFile.name }
    }));

    // Auto-detect label via Pyodide
    if (pyodide) {
      try {
        const detectedLabel = await autoDetectLabel(pyodide, uploadedFile);
        if (detectedLabel) {
          setFiles(prev => ({
            ...prev,
            [fileKey]: { ...prev[fileKey], file: uploadedFile, name: uploadedFile.name, label: detectedLabel }
          }));
          addLog(`Auto-detected "${uploadedFile.name}" as ${detectedLabel}`);
        }
      } catch {
        // Auto-detect failed silently — user can still pick manually
        addLog(`Could not auto-detect label for "${uploadedFile.name}". Please select manually.`);
      }
    }
  };

  const handleFileUpload = (e, fileKey) => {
    const uploadedFile = e.target.files[0];
    if (!uploadedFile) return;
    processUploadedFile(uploadedFile, fileKey);
  };

  const handleLabelChange = (val, fileKey) => {
    setFiles(prev => ({ ...prev, [fileKey]: { ...prev[fileKey], label: val } }));
  };

  // --- Drag and Drop ---
  const handleDragOver = (e, fileKey) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(prev => ({ ...prev, [fileKey]: true }));
  };

  const handleDragLeave = (e, fileKey) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(prev => ({ ...prev, [fileKey]: false }));
  };

  const handleDrop = (e, fileKey) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(prev => ({ ...prev, [fileKey]: false }));
    const droppedFile = e.dataTransfer?.files?.[0];
    if (droppedFile) {
      processUploadedFile(droppedFile, fileKey);
    }
  };

  // --- Reset ---
  const resetApp = () => {
    setFiles({
      file1: { file: null, label: null, name: null },
      file2: { file: null, label: null, name: null }
    });
    setFileErrors({ file1: null, file2: null });
    if (processedFileUrl) URL.revokeObjectURL(processedFileUrl);
    if (processedPdfUrl) URL.revokeObjectURL(processedPdfUrl);
    setProcessedFileUrl(null);
    setProcessedPdfUrl(null);
    setPreviewData(null);
    setLogs([]);
    setError(null);
    setWarningCount(0);
    setShowLogs(false);
    setShowPreview(true);
    addLog("App reset. Ready for new files.");
  };

  // ----------------------------------------------------------------------
  // 3. PDF Generation (Single Page — preserved)
  // ----------------------------------------------------------------------
  const generatePDF = (data) => {
    addLog("Generating formatted single-page PDF document...");
    const body = [];

    data.forEach(row => {
      const valA = String(row[0] || '').trim();
      const valB = String(row[1] || '').trim();

      // Handle empty spacing rows
      if (row.every(c => c === "" || c === null)) {
        body.push([{ content: '', colSpan: 7, styles: { minCellHeight: 15, fillColor: [255, 255, 255], lineWidth: 0 } }]);
        return;
      }

      const isPartyHeader = (valA && !valB && !valA.toLowerCase().startsWith("rcpt") && !valA.startsWith("Outstanding Summary") && !valA.startsWith("---") && valA !== "Type");

      if (isPartyHeader) {
        body.push([
          { content: valA, colSpan: 6, styles: { fillColor: [217, 225, 242], fontStyle: 'bold', halign: 'center', lineWidth: 0.5 } },
          { content: formatIndianNumber(row[6]), styles: { fillColor: [217, 225, 242], fontStyle: 'bold', halign: 'right', lineWidth: 0.5 } }
        ]);
      } else if (valA === "Type") {
        body.push(row.map(cell => ({ content: String(cell), styles: { fillColor: [242, 242, 242], fontStyle: 'bold', halign: 'center', lineWidth: 0.5 } })));
      } else if (valA.startsWith("---")) {
        const fill = valA.includes("Medical") ? [198, 224, 180] : [248, 203, 173];
        body.push([{ content: valA, colSpan: 7, styles: { fillColor: fill, fontStyle: 'bold', halign: 'center', lineWidth: 0.5 } }]);
      } else if (valA.startsWith("Outstanding Summary")) {
        body.push([{ content: valA, colSpan: 7, styles: { fillColor: [68, 114, 196], textColor: [255, 255, 255], fontStyle: 'bold', halign: 'center', lineWidth: 0.5 } }]);
      } else {
        // Normal data rows — format numeric columns with Indian grouping
        body.push(row.map((cell, i) => {
          let displayValue = (cell !== "" && cell !== null) ? String(cell) : "";
          // Format Debit (3), Credit (4), Balance (5) with Indian numbers
          if (i >= 3 && i <= 5 && cell !== "" && cell !== null && !isNaN(cell)) {
            displayValue = formatIndianNumber(cell);
          }
          return {
            content: displayValue,
            styles: { halign: (i > 2 && cell !== "") ? 'right' : 'left', lineWidth: 0.1, lineColor: [200, 200, 200] }
          };
        }));
      }
    });

    // ---------------------------------------------------------
    // DYNAMIC HEIGHT CALCULATION:
    // 1. Draw table on a fake, infinitely tall document
    // ---------------------------------------------------------
    const A4_WIDTH_PT = 595.28;
    const TOP_MARGIN = 45; // Extra space for title
    const dummyDoc = new jsPDF('p', 'pt', [A4_WIDTH_PT, 99999]);

    autoTable(dummyDoc, {
      body: body,
      theme: 'grid',
      styles: { fontSize: 8, cellPadding: 3, font: 'helvetica' },
      margin: { top: TOP_MARGIN, left: 30, right: 30 },
      tableWidth: 'auto',
    });

    // 2. Extract exactly how tall the table ended up being
    const totalContentHeight = dummyDoc.lastAutoTable.finalY + 30;
    const finalPageHeight = Math.max(841.89, totalContentHeight);

    // ---------------------------------------------------------
    // FINAL PDF GENERATION:
    // Create the actual document using the exactly measured height
    // ---------------------------------------------------------
    const finalDoc = new jsPDF('p', 'pt', [A4_WIDTH_PT, finalPageHeight]);

    // Add title
    finalDoc.setFontSize(13);
    finalDoc.setFont('helvetica', 'bold');
    finalDoc.setTextColor(50, 50, 50);
    finalDoc.text(`Outstanding Report  —  ${getDisplayDate()}`, A4_WIDTH_PT / 2, 25, { align: 'center' });

    autoTable(finalDoc, {
      body: body,
      theme: 'grid',
      styles: { fontSize: 8, cellPadding: 3, textColor: [0, 0, 0], font: 'helvetica' },
      margin: { top: TOP_MARGIN, left: 30, right: 30 },
      tableWidth: 'auto',
    });

    return finalDoc;
  };

  // ----------------------------------------------------------------------
  // 4. The Python Script (Template String) — with warning tracking & formatting
  // ----------------------------------------------------------------------
  const getPythonScript = (file1Name, file2Name) => `
import pandas as pd
import os
import json
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

FILE_1 = "${file1Name}" 
FILE_2 = "${file2Name}" 
OUTPUT_FILE = "Final_Merged_Report.xlsx"

# Track warnings for rows that couldn't be parsed
_warnings = []

def format_indian(num):
    """Format a number with Indian comma grouping (1,23,456.00)"""
    try:
        num = float(num)
    except (ValueError, TypeError):
        return str(num)
    is_neg = num < 0
    abs_num = abs(num)
    int_part = int(abs_num)
    dec_str = f"{abs_num - int_part:.2f}"[2:]
    s = str(int_part)
    if len(s) <= 3:
        result = s
    else:
        result = s[-3:]
        s = s[:-3]
        while len(s) > 2:
            result = s[-2:] + ',' + result
            s = s[:-2]
        if s:
            result = s + ',' + result
    return ('-' if is_neg else '') + result + '.' + dec_str

def process_file(filepath, master_data, file_label):
    if not os.path.exists(filepath): return
    fname = os.path.basename(filepath)
    try:
        if filepath.lower().endswith('.xls'):
            df = pd.read_excel(filepath, header=None, engine='xlrd')
        else:
            df = pd.read_excel(filepath, header=None, engine='openpyxl')
    except Exception as e:
        _warnings.append(f"[{fname}] Could not read file: {e}")
        return

    current_party_key = None
    for index, row in df.iterrows():
        row_data = ["" if pd.isna(x) else x for x in row.tolist()]
        if len(row_data) < 2: continue

        # Separator rows (all empty) — reset party context so subsequent
        # non-party rows (city subtotals, headers) don't leak into party data
        if not any(str(x).strip() for x in row_data):
            current_party_key = None
            continue

        col_b_value = str(row_data[1]).strip()
        if col_b_value.startswith('-'):
            party_name = col_b_value.lstrip('-').strip()
            city = str(row_data[3]).strip() if len(row_data) > 3 else ""
            unique_key = (party_name + city).lower().replace(" ", "")
            try: amount = float(row_data[4]) if len(row_data) > 4 else 0.0
            except ValueError: amount = 0.0

            if unique_key not in master_data:
                master_data[unique_key] = {
                    'original_name': party_name, 'city': city,
                    'transactions_file1': [], 'transactions_file2': [],
                    'total_combined': 0.0, 'totals_breakdown': {'file1': 0.0, 'file2': 0.0}
                }
            master_data[unique_key]['totals_breakdown'][file_label] += amount
            master_data[unique_key]['total_combined'] += amount
            current_party_key = unique_key

        elif current_party_key:
            col_a_val = str(row_data[0]).strip()

            # Skip the MARG ERP footer/ad row
            if col_a_val.startswith("Digital Purchase"):
                continue

            is_rcpt_row = col_a_val.lower().startswith("rcpt")

            if len(row_data) > 3:
                val_d = row_data[3]
                if isinstance(val_d, str) and " " in val_d.strip():
                    parts = val_d.strip().split()
                    if len(parts) == 2:
                        try:
                            debit, credit = float(parts[0]), float(parts[1])
                            row_data[3] = debit
                            row_data.insert(4, credit)
                        except ValueError:
                            _warnings.append(f"[{fname}] Row {index + 1}: Could not parse debit/credit from '{val_d}'")
            try:
                if len(row_data) > 3:
                    current_debit = float(row_data[3])
                    if current_debit < 0:
                        row_data[3] = 0.0
                        if len(row_data) > 4: row_data[4] = abs(current_debit)
                        else: row_data.append(abs(current_debit))
            except (ValueError, TypeError):
                # Rcpt rows have None in col D — that's expected, not a warning
                if not is_rcpt_row:
                    _warnings.append(f"[{fname}] Row {index + 1}: Could not process debit value '{row_data[3] if len(row_data) > 3 else 'N/A'}'")

            if not is_rcpt_row: row_data[0] = ""
            master_data[current_party_key][f'transactions_{file_label}'].append(row_data)

def style_excel_file(filename):
    wb = load_workbook(filename)
    ws = wb.active
    ws.title = "Outstanding Report"
    HEADER_SIZE, REGULAR_SIZE = 14, 12
    party_header_font = Font(bold=True, size=HEADER_SIZE, color="000000")
    summary_font = Font(bold=True, size=HEADER_SIZE, color="FFFFFF")
    section_font = Font(bold=True, size=12)
    sub_header_font = Font(bold=True, size=REGULAR_SIZE)
    regular_font = Font(size=REGULAR_SIZE)

    party_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    medical_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    surgical_fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
    sub_header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    summary_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    thin = Side(style='thin')
    no_side = Side(border_style=None)
    thin_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    indian_num_format = '#,##,##0.00'

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        if len(row) < 2: continue
        val_a = str(row[0].value).strip() if row[0].value else ""
        val_b = str(row[1].value).strip() if row[1].value else ""

        is_party_header = (val_a and not val_b and not val_a.lower().startswith("rcpt") and not val_a.startswith("Outstanding Summary") and not val_a.startswith("---"))

        if is_party_header:
            if len(row) >= 7:
                for i in range(7):
                    row[i].fill = party_fill
                    row[i].border = Border(top=thin, left=thin if i==6 else no_side, right=thin if i>=5 else no_side, bottom=no_side)
                row[0].font, row[0].alignment = party_header_font, Alignment(horizontal='center', vertical='center')
                row[6].font, row[6].alignment = party_header_font, Alignment(horizontal='right', vertical='center')
                if isinstance(row[6].value, (int, float)):
                    row[6].number_format = indian_num_format
                ws.merge_cells(start_row=row[0].row, start_column=1, end_row=row[0].row, end_column=6)
        elif val_a.startswith("---"):
            if len(row) >= 7:
                ws.merge_cells(start_row=row[0].row, start_column=1, end_row=row[0].row, end_column=7)
                row[0].font, row[0].alignment = section_font, Alignment(horizontal='center', vertical='center')
                for cell in row: cell.fill, cell.border = (medical_fill if "Medical" in val_a else surgical_fill), thin_border
        elif val_b == "Bill":
            if len(row) >= 7:
                for cell in row: cell.font, cell.fill, cell.border, cell.alignment = sub_header_font, sub_header_fill, thin_border, Alignment(horizontal='center')
        elif val_a.startswith("Outstanding Summary"):
            if len(row) >= 7:
                for i in range(7): row[i].fill, row[i].border = summary_fill, Border(top=thin, left=no_side, right=no_side, bottom=no_side)
                row[0].font, row[0].alignment = summary_font, Alignment(horizontal='center', vertical='center')
                ws.merge_cells(start_row=row[0].row, start_column=1, end_row=row[0].row, end_column=7)
        else:
            if any(cell.value for cell in row):
                if len(row) >= 7:
                    for idx_c, cell in enumerate(row):
                        cell.font, cell.border = regular_font, thin_border
                        # Apply Indian number format to Debit(D=3), Credit(E=4), Balance(F=5) columns
                        if idx_c in (3, 4, 5) and isinstance(cell.value, (int, float)):
                            cell.number_format = indian_num_format

    if ws.max_column > 0:
        for i in range(1, ws.max_column + 1):
            col_letter = get_column_letter(i)
            max_length = 0
            for cell in ws[col_letter]:
                try:
                    if cell.value:
                        val = str(cell.value)
                        current_row = ws[cell.row]
                        if len(current_row) < 2: continue
                        val_b_check = str(current_row[1].value).strip() if len(current_row) > 1 and current_row[1].value else ""
                        if not (val.startswith("Outstanding Summary") or val.startswith("---") or val_b_check == "Bill"):
                            max_length = max(max_length, len(val))
                except: pass
            ws.column_dimensions[col_letter].width = max(max_length + 2, 8) if i == 1 else min((max_length + 2), 60)
    wb.save(filename)

# Execution
combined_data = {}
if os.path.exists(FILE_1): process_file(FILE_1, combined_data, 'file1')
if os.path.exists(FILE_2): process_file(FILE_2, combined_data, 'file2')

final_output_rows = []
transaction_header = ["Type", "Bill", "Date", "Debit", "Credit", "Balance", "Days"]

for key, data in combined_data.items():
    header_row = [f"{data['original_name']}  ({data['city']})", "", "", "", "", "", data['total_combined']]
    final_output_rows.append(header_row)
    final_output_rows.append(transaction_header)

    if data['transactions_file1']:
        final_output_rows.append(["--- Medical ---", "", "", "", "", "", ""])
        for trans in data['transactions_file1']:
            padded = list(trans)
            while len(padded) < 7: padded.append("")
            final_output_rows.append(padded[:7])

    if data['transactions_file2']:
        final_output_rows.append(["--- Surgical ---", "", "", "", "", "", ""])
        for trans in data['transactions_file2']:
            padded = list(trans)
            while len(padded) < 7: padded.append("")
            final_output_rows.append(padded[:7])

    t1, t2, total = data['totals_breakdown']['file1'], data['totals_breakdown']['file2'], data['total_combined']
    final_output_rows.append([f"Outstanding Summary   |   Medical: {format_indian(t1)}   |   Surgical: {format_indian(t2)}   |   Total: {format_indian(total)}", "", "", "", "", "", ""])
    final_output_rows.append([""] * 7)

# 1. Save Excel
df = pd.DataFrame(final_output_rows)
df.to_excel(OUTPUT_FILE, index=False, header=False)
style_excel_file(OUTPUT_FILE)

# 2. Return JSON data to Javascript for PDF Generation
clean_rows = []
for r in final_output_rows:
    clean_row = []
    for val in r:
        if pd.isna(val): clean_row.append("")
        elif isinstance(val, (int, float)): clean_row.append(float(val))
        else: clean_row.append(str(val))
    clean_rows.append(clean_row)

json.dumps({"data": clean_rows, "warnings": len(_warnings), "warning_details": _warnings[:20]})
`;

  // ----------------------------------------------------------------------
  // 5. Logic: Execute Pipeline
  // ----------------------------------------------------------------------
  const handleMerge = async () => {
    if (!pyodide) { setError("Python environment is not ready."); return; }

    setIsProcessing(true);
    setLogs([]);
    setError(null);
    setWarningCount(0);
    setPreviewData(null);
    if (processedFileUrl) URL.revokeObjectURL(processedFileUrl);
    if (processedPdfUrl) URL.revokeObjectURL(processedPdfUrl);
    setProcessedFileUrl(null);
    setProcessedPdfUrl(null);

    const { file1, file2 } = files;
    if (!file1.file || !file2.file) { setError("Please upload both files."); setIsProcessing(false); return; }
    if (!file1.label || !file2.label) { setError("Please select the source label for both files."); setIsProcessing(false); return; }
    if (file1.label === file2.label) { setError("Both files have the same label. Each file must belong to a different entity."); setIsProcessing(false); return; }

    try {
      addLog("Reading uploaded files...");
      let medicalVfsName = "";
      let surgicalVfsName = "";

      const processUpload = async (fData) => {
        const ext = fData.file.name.split('.').pop();
        const arrayBuffer = await fData.file.arrayBuffer();
        const targetName = fData.label === "RAHUL MEDICAL & SURGICAL" ? `medical_input.${ext}` : `surgical_input.${ext}`;
        if (fData.label === "RAHUL MEDICAL & SURGICAL") medicalVfsName = targetName;
        else surgicalVfsName = targetName;

        pyodide.FS.writeFile(targetName, new Uint8Array(arrayBuffer));
        addLog(`Saved as "${targetName}" in virtual memory.`);
      };

      await processUpload(file1);
      await processUpload(file2);

      addLog("Executing Python logic...");
      pyodide.setStdout({ batched: (msg) => addLog(`[PY] ${msg}`) });

      const finalScript = getPythonScript(medicalVfsName, surgicalVfsName);

      // Get JSON payload back from Python
      const jsonResult = await pyodide.runPythonAsync(finalScript);
      const result = JSON.parse(jsonResult);
      const parsedData = result.data;
      const warnCount = result.warnings || 0;
      const warnDetails = result.warning_details || [];

      // Show warnings if any
      if (warnCount > 0) {
        setWarningCount(warnCount);
        warnDetails.forEach(w => addLog(`[WARNING] ${w}`));
        addLog(`${warnCount} warning(s) during processing. Some rows may have been skipped.`);
      }

      // Build preview
      const preview = extractPreviewInfo(parsedData);
      setPreviewData(preview);

      // Create PDF
      const pdfDoc = generatePDF(parsedData);
      const pdfBlob = pdfDoc.output('blob');
      setProcessedPdfUrl(URL.createObjectURL(pdfBlob));

      // Fetch Excel
      if (pyodide.FS.analyzePath("Final_Merged_Report.xlsx").exists) {
        const fileContent = pyodide.FS.readFile("Final_Merged_Report.xlsx");
        const blob = new Blob([fileContent], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        setProcessedFileUrl(URL.createObjectURL(blob));
        addLog("Merge successful! Excel & PDF ready for download.");
      }

    } catch (err) {
      console.error(err);
      setError(classifyError(err.message));
      addLog(`Error: ${err.message}`);
      setShowLogs(true);
    } finally {
      setIsProcessing(false);
    }
  };

  // --- Output filename with date ---
  const outputBaseName = `Merged_Report_${getDateString()}`;

  // ====================================================================
  // RENDER
  // ====================================================================
  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans p-4 sm:p-6">
      <div className="max-w-3xl mx-auto bg-white shadow-xl rounded-xl overflow-hidden border border-slate-200">

        {/* ── Header ── */}
        <div className="bg-blue-600 p-6 text-white">
          <h1 className="text-2xl font-bold flex items-center gap-2">
            <FileSpreadsheet aria-hidden="true" /> Excel Merge Tool
          </h1>
          <p className="text-blue-100 mt-2 text-sm">
            Securely merge and export Surgical and Medical reports to Excel and PDF.
          </p>
        </div>

        {/* ── Progress Stepper ── */}
        {!loading && (
          <div className="px-6 pt-5 pb-1" aria-label="Progress steps">
            <div className="flex items-center justify-center">
              {[
                { num: 1, label: "Upload" },
                { num: 2, label: "Process" },
                { num: 3, label: "Results" }
              ].map((step, idx) => {
                const isCompleted = currentStep > step.num;
                const isCurrent = currentStep === step.num;
                return (
                  <React.Fragment key={step.num}>
                    {idx > 0 && (
                      <div className={`h-0.5 w-12 sm:w-20 transition-colors ${isCompleted ? 'bg-green-500' : 'bg-slate-200'}`} />
                    )}
                    <div className="flex flex-col items-center gap-1">
                      <div className={`w-8 h-8 rounded-full flex items-center justify-center text-sm font-bold transition-colors ${
                        isCompleted ? 'bg-green-500 text-white' :
                        isCurrent ? 'bg-blue-600 text-white' :
                        'bg-slate-200 text-slate-400'
                      }`}>
                        {isCompleted ? <CheckCircle size={16} aria-hidden="true" /> : step.num}
                      </div>
                      <span className={`text-xs font-medium ${isCurrent ? 'text-blue-600' : isCompleted ? 'text-green-600' : 'text-slate-400'}`}>
                        {step.label}
                      </span>
                    </div>
                  </React.Fragment>
                );
              })}
            </div>
          </div>
        )}

        {/* ── Loading State ── */}
        {loading ? (
          <div className="p-10 flex flex-col items-center justify-center text-center">
            {initError ? (
              <div className="space-y-4">
                <AlertCircle className="text-red-500 mx-auto" size={40} aria-hidden="true" />
                <p className="text-red-600 font-medium text-sm">{initError}</p>
                <button
                  onClick={retryInit}
                  className="px-5 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 text-sm font-medium transition-colors"
                >
                  Retry
                </button>
              </div>
            ) : (
              <>
                {/* Progress dots */}
                <div className="w-full max-w-xs mb-5">
                  <div className="flex justify-between mb-2">
                    {LOADING_STEPS.map((_, i) => (
                      <div key={i} className="flex flex-col items-center">
                        <div className={`w-3 h-3 rounded-full transition-colors ${
                          i < loadingStep ? 'bg-green-500' :
                          i === loadingStep ? 'bg-blue-600 animate-pulse' :
                          'bg-slate-200'
                        }`} />
                      </div>
                    ))}
                  </div>
                  <div className="h-2 bg-slate-200 rounded-full overflow-hidden">
                    <div
                      className="h-full bg-blue-600 rounded-full transition-all duration-700 ease-out"
                      style={{ width: `${((loadingStep + 0.5) / LOADING_STEPS.length) * 100}%` }}
                    />
                  </div>
                </div>
                <p className="text-slate-600 font-medium text-sm">{LOADING_STEPS[loadingStep]}</p>
                <p className="text-slate-400 text-xs mt-2">This usually takes 10–20 seconds on first visit</p>
              </>
            )}
          </div>
        ) : (
          <div className="p-6 space-y-6">
            {/* ── Error Banner ── */}
            {error && (
              <div className="bg-red-50 text-red-700 p-4 rounded-lg flex items-start gap-3 border border-red-200" role="alert">
                <AlertCircle className="shrink-0 mt-0.5" size={20} aria-hidden="true" />
                <div className="text-sm font-medium">{error}</div>
              </div>
            )}

            {/* ── Warning Banner ── */}
            {warningCount > 0 && !error && (
              <div className="bg-amber-50 text-amber-700 p-4 rounded-lg flex items-start gap-3 border border-amber-200" role="alert">
                <AlertTriangle className="shrink-0 mt-0.5" size={20} aria-hidden="true" />
                <div className="text-sm font-medium">
                  {warningCount} row(s) had unexpected data and may have been skipped. Check the System Logs for details.
                </div>
              </div>
            )}

            {/* ── File Upload Inputs ── */}
            <div className="grid md:grid-cols-2 gap-6">
              {[1, 2].map((num) => {
                const key = `file${num}`;
                const fData = files[key];
                const otherKey = num === 1 ? 'file2' : 'file1';
                const otherLabel = files[otherKey].label;

                return (
                  <div key={key} className="bg-slate-50 p-4 rounded-lg border border-slate-200">
                    <h2 className="font-semibold text-slate-700 mb-3 flex items-center gap-2">
                      <div className="bg-blue-100 text-blue-700 w-6 h-6 rounded-full flex items-center justify-center text-xs">{num}</div>
                      Attachment {num}
                    </h2>

                    {/* Drop Zone */}
                    <div
                      className="relative group mb-3"
                      onDragOver={(e) => handleDragOver(e, key)}
                      onDragEnter={(e) => handleDragOver(e, key)}
                      onDragLeave={(e) => handleDragLeave(e, key)}
                      onDrop={(e) => handleDrop(e, key)}
                    >
                      <input
                        type="file"
                        accept=".xls,.xlsx"
                        onChange={(e) => handleFileUpload(e, key)}
                        className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                        aria-label={`Upload Excel file for Attachment ${num}`}
                      />
                      <div className={`border-2 border-dashed rounded-lg p-6 text-center transition-all ${
                        dragActive[key] ? 'border-blue-500 bg-blue-50 scale-[1.02]' :
                        fData.file ? 'border-green-400 bg-green-50' :
                        'border-slate-300 hover:border-blue-400 bg-white'
                      }`}>
                        {dragActive[key] ? (
                          <div className="flex flex-col items-center text-blue-500">
                            <Upload size={24} className="mb-2" aria-hidden="true" />
                            <span className="text-xs font-semibold">Drop file here</span>
                          </div>
                        ) : fData.file ? (
                          <div className="flex flex-col items-center text-green-700">
                            <CheckCircle size={24} className="mb-2" aria-hidden="true" />
                            <span className="text-xs font-medium truncate w-full px-2" title={fData.name}>{fData.name}</span>
                          </div>
                        ) : (
                          <div className="flex flex-col items-center text-slate-400 group-hover:text-blue-500">
                            <Upload size={24} className="mb-2" aria-hidden="true" />
                            <span className="text-xs font-medium">Click or Drag Excel File</span>
                            <span className="text-[10px] mt-1">Supports .xls and .xlsx</span>
                          </div>
                        )}
                      </div>
                    </div>

                    {/* File validation error */}
                    {fileErrors[key] && (
                      <p className="text-red-500 text-xs mb-2 flex items-center gap-1">
                        <AlertCircle size={12} aria-hidden="true" /> {fileErrors[key]}
                      </p>
                    )}

                    {/* Label selector */}
                    <label htmlFor={`label-${key}`} className="block text-xs font-semibold text-slate-500 uppercase tracking-wider mb-1">
                      Identify Source
                    </label>
                    <select
                      id={`label-${key}`}
                      value={fData.label || ""}
                      onChange={(e) => handleLabelChange(e.target.value, key)}
                      className="w-full text-sm p-2 rounded border border-slate-300 focus:ring-2 focus:ring-blue-500 outline-none bg-white cursor-pointer"
                      aria-label={`Select source entity for Attachment ${num}`}
                    >
                      <option value="" disabled>-- Select Entity --</option>
                      {OPTIONS.map(opt => (
                        <option key={opt} value={opt} disabled={opt === otherLabel}>
                          {opt}{opt === otherLabel ? ' (used by other file)' : ''}
                        </option>
                      ))}
                    </select>
                    {fData.label && (
                      <p className="text-xs text-green-600 mt-1 flex items-center gap-1">
                        <CheckCircle size={10} aria-hidden="true" /> {fData.label}
                      </p>
                    )}
                  </div>
                );
              })}
            </div>

            {/* ── Action Area ── */}
            <div className="border-t border-slate-100 pt-6 flex flex-col items-center gap-4">
              {!processedFileUrl ? (
                <button
                  onClick={handleMerge}
                  disabled={isProcessing || !pyodide || currentStep < 2}
                  className={`flex items-center gap-2 px-8 py-3 rounded-full font-bold shadow-lg transition-all ${
                    (isProcessing || !pyodide || currentStep < 2)
                      ? 'bg-slate-300 text-slate-500 cursor-not-allowed'
                      : 'bg-blue-600 hover:bg-blue-700 text-white hover:-translate-y-1'
                  }`}
                  aria-busy={isProcessing}
                >
                  {isProcessing ? <><RefreshCw className="animate-spin" size={20} aria-hidden="true" /> Processing...</> : "Merge & Process Files"}
                </button>
              ) : (
                <>
                  {/* ── Preview Panel ── */}
                  {previewData && (
                    <div className="w-full bg-slate-50 rounded-lg border border-slate-200 overflow-hidden">
                      <button
                        onClick={() => setShowPreview(!showPreview)}
                        className="flex items-center justify-between w-full p-4 hover:bg-slate-100 transition-colors"
                        aria-expanded={showPreview}
                      >
                        <h3 className="font-semibold text-slate-700 flex items-center gap-2 text-sm">
                          <FileSpreadsheet size={16} aria-hidden="true" /> Merge Summary
                        </h3>
                        {showPreview ? <ChevronUp size={16} aria-hidden="true" /> : <ChevronDown size={16} aria-hidden="true" />}
                      </button>

                      {showPreview && (
                        <div className="px-4 pb-4 space-y-3">
                          {/* Totals cards */}
                          <div className="grid grid-cols-3 gap-2 sm:gap-3">
                            <div className="bg-green-50 p-2.5 sm:p-3 rounded-lg border border-green-200 text-center">
                              <div className="text-[10px] sm:text-xs text-green-600 font-medium">Medical</div>
                              <div className="text-sm sm:text-lg font-bold text-green-700 truncate" title={formatIndianNumber(previewData.medicalTotal)}>
                                ₹{formatIndianNumber(previewData.medicalTotal)}
                              </div>
                            </div>
                            <div className="bg-orange-50 p-2.5 sm:p-3 rounded-lg border border-orange-200 text-center">
                              <div className="text-[10px] sm:text-xs text-orange-600 font-medium">Surgical</div>
                              <div className="text-sm sm:text-lg font-bold text-orange-700 truncate" title={formatIndianNumber(previewData.surgicalTotal)}>
                                ₹{formatIndianNumber(previewData.surgicalTotal)}
                              </div>
                            </div>
                            <div className="bg-blue-50 p-2.5 sm:p-3 rounded-lg border border-blue-200 text-center">
                              <div className="text-[10px] sm:text-xs text-blue-600 font-medium">Grand Total</div>
                              <div className="text-sm sm:text-lg font-bold text-blue-700 truncate" title={formatIndianNumber(previewData.grandTotal)}>
                                ₹{formatIndianNumber(previewData.grandTotal)}
                              </div>
                            </div>
                          </div>

                          {/* Party count */}
                          <p className="text-sm text-slate-600">
                            <span className="font-semibold text-blue-700">{previewData.parties.length}</span> parties matched across files
                          </p>

                          {/* Party list */}
                          <div className="max-h-48 overflow-y-auto divide-y divide-slate-100 rounded border border-slate-200 bg-white">
                            {previewData.parties.map((party, i) => (
                              <div key={i} className="flex justify-between py-2 px-3 text-sm hover:bg-slate-50">
                                <span className="text-slate-600 truncate mr-3">{party.name}</span>
                                <span className="text-slate-800 font-medium whitespace-nowrap">₹{formatIndianNumber(party.total)}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  {/* ── Download Buttons ── */}
                  <div className="flex flex-col sm:flex-row gap-3">
                    <a
                      href={processedFileUrl}
                      download={`${outputBaseName}.xlsx`}
                      className="flex items-center justify-center gap-2 px-6 py-2.5 rounded-full font-bold shadow-md bg-green-600 hover:bg-green-700 text-white text-sm transition-all hover:-translate-y-0.5"
                    >
                      <FileSpreadsheet size={18} aria-hidden="true" /> Download Excel
                    </a>
                    <a
                      href={processedPdfUrl}
                      download={`${outputBaseName}.pdf`}
                      className="flex items-center justify-center gap-2 px-6 py-2.5 rounded-full font-bold shadow-md bg-red-500 hover:bg-red-600 text-white text-sm transition-all hover:-translate-y-0.5"
                    >
                      <FileText size={18} aria-hidden="true" /> Download PDF
                    </a>
                    <button onClick={resetApp} className="flex items-center justify-center gap-2 px-5 py-2.5 rounded-full font-bold border-2 border-slate-200 text-slate-600 text-sm hover:bg-slate-50 transition-all">
                      <RotateCcw size={16} aria-hidden="true" /> Start Over
                    </button>
                  </div>
                </>
              )}
            </div>

            {/* ── System Logs ── */}
            <div className="mt-2">
              <button
                onClick={() => setShowLogs(!showLogs)}
                className="flex items-center gap-2 text-xs font-semibold text-slate-400 hover:text-slate-600 transition-colors mx-auto"
                aria-expanded={showLogs}
                aria-controls="log-panel"
              >
                <Terminal size={14} aria-hidden="true" /> {showLogs ? "Hide System Logs" : "Show System Logs"} {showLogs ? <ChevronUp size={14} aria-hidden="true" /> : <ChevronDown size={14} aria-hidden="true" />}
              </button>
              {showLogs && (
                <div id="log-panel" role="log" aria-live="polite" className="bg-slate-900 rounded-lg p-4 font-mono text-xs text-green-400 h-48 overflow-y-auto mt-3 shadow-inner">
                  <div className="text-slate-500 border-b border-slate-800 pb-2 mb-2 flex justify-between"><span>System Logs</span></div>
                  {logs.length === 0 && <span className="text-slate-600 italic">Waiting for input...</span>}
                  {logs.map((log, i) => <div key={i} className="mb-1 leading-relaxed">{log}</div>)}
                </div>
              )}
            </div>

          </div>
        )}
      </div>
    </div>
  );
};

export default App;