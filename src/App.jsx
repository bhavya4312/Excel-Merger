import React, { useState, useEffect, useRef, useMemo } from 'react';
import { Upload, FileSpreadsheet, RefreshCw, AlertCircle, CheckCircle, ChevronDown, ChevronUp, Terminal, RotateCcw, FileText, AlertTriangle, FolderArchive } from 'lucide-react';
import { initPyodideWorker, autoDetectLabelWorker, runMergeWorker } from './pyodideClient';

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



/** Extract summary info from processed data for the preview panel */
const extractPreviewInfo = (data) => {
  const parties = [];
  let medicalTotal = 0;
  let surgicalTotal = 0;

  data.forEach(row => {
    const valA = String(row[0] || '').trim();
    const valB = String(row[1] || '').trim();

    // Party header detection
    const isPartyHeader = (valA && !valB &&
      !valA.toLowerCase().startsWith("rcpt") &&
      !valA.startsWith("Outstanding Summary") &&
      !valA.startsWith("---") &&
      !valA.startsWith("--") &&
      valA !== "Bill" &&
      valA !== "Type");

    if (isPartyHeader) {
      const partyTotal = Number(row[5] !== "" && row[5] !== undefined ? row[5] : (row[6] || 0)) || 0;
      parties.push({ name: valA, total: partyTotal });
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

/** Sanitize filename for ZIP archive extraction across OS filesystems */
const sanitizeFilename = (name) => {
  return String(name || 'Customer')
    .replace(/[<>:"/\\|?*]/g, '_')
    .replace(/\s+/g, ' ')
    .trim();
};

/** Group flat merged rows by individual party for card image generation */
const groupRowsByParty = (data) => {
  const parties = [];
  let currentParty = null;
  let currentSection = null;

  (data || []).forEach((row) => {
    const valA = String(row[0] || '').trim();
    const valB = String(row[1] || '').trim();

    // Party Header detection
    const isPartyHeader = (valA && !valB &&
      !valA.toLowerCase().startsWith("rcpt") &&
      !valA.startsWith("Outstanding Summary") &&
      !valA.startsWith("---") &&
      !valA.startsWith("--") &&
      valA !== "Bill" &&
      valA !== "Type");

    if (isPartyHeader) {
      const partyTotal = Number(row[5] !== "" && row[5] !== undefined ? row[5] : (row[6] || 0)) || 0;
      currentParty = {
        displayName: valA,
        totalAmount: partyTotal,
        medicalTransactions: [],
        surgicalTransactions: [],
        medicalTotal: 0,
        surgicalTotal: 0,
      };
      parties.push(currentParty);
      currentSection = null;
      return;
    }

    if (!currentParty) return;

    if (valA.startsWith("--- Medical") || valA.startsWith("-- Medical")) {
      currentSection = "medical";
      return;
    }
    if (valA.startsWith("--- Surgical") || valA.startsWith("-- Surgical")) {
      currentSection = "surgical";
      return;
    }
    if (valA.startsWith("Outstanding Summary")) {
      const medMatch = valA.match(/Medical:\s*([\d,.]+)/);
      const surgMatch = valA.match(/Surgical:\s*([\d,.]+)/);
      if (medMatch) currentParty.medicalTotal = parseFloat(medMatch[1].replace(/,/g, ''));
      if (surgMatch) currentParty.surgicalTotal = parseFloat(surgMatch[1].replace(/,/g, ''));
      currentSection = null;
      return;
    }
    if (valA === "Type" || valA === "Bill" || valB === "Bill") {
      return;
    }

    // Has valid transaction content
    const hasData = row.some(cell => cell !== "" && cell !== null && cell !== undefined);
    if (hasData && currentSection) {
      if (currentSection === "medical") {
        currentParty.medicalTransactions.push(row);
      } else if (currentSection === "surgical") {
        currentParty.surgicalTransactions.push(row);
      }
    }
  });

  return parties;
};

/** Render a single party's high-DPI card to a PNG Blob */
const renderPartyCardCanvas = (party, formattedDate) => {
  return new Promise((resolve) => {
    const dpr = 2; // High-DPI 2x scale for sharp typography on mobile screens
    const width = 880;
    const padX = 24;
    const padY = 24;
    const contentWidth = width - padX * 2;

    const headerHeight = 92;
    const sectionGap = 16;
    const bannerHeight = 28;
    const tableHeaderHeight = 30;
    const rowHeight = 28;
    const summaryHeight = 38;

    const medRows = party.medicalTransactions.length;
    const surgRows = party.surgicalTransactions.length;

    let tablesHeight = 0;
    if (medRows > 0) {
      tablesHeight += bannerHeight + tableHeaderHeight + (medRows * rowHeight) + sectionGap;
    }
    if (surgRows > 0) {
      tablesHeight += bannerHeight + tableHeaderHeight + (surgRows * rowHeight) + sectionGap;
    }

    const totalHeight = padY + headerHeight + sectionGap + tablesHeight + summaryHeight + padY;

    const canvas = document.createElement('canvas');
    canvas.width = width * dpr;
    canvas.height = totalHeight * dpr;

    const ctx = canvas.getContext('2d');
    ctx.scale(dpr, dpr);

    // Background
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, width, totalHeight);

    // Subtle card border
    ctx.strokeStyle = '#E2E8F0';
    ctx.lineWidth = 1;
    ctx.strokeRect(10, 10, width - 20, totalHeight - 20);

    let y = padY;

    // --- 1. Top Header Box ---
    ctx.fillStyle = '#F8FAFC';
    ctx.fillRect(padX, y, contentWidth, headerHeight);
    ctx.strokeStyle = '#CBD5E1';
    ctx.lineWidth = 1;
    ctx.strokeRect(padX, y, contentWidth, headerHeight);

    // Top Badge & Date
    ctx.fillStyle = '#1D4ED8';
    ctx.font = 'bold 11px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'left';
    ctx.fillText('OUTSTANDING STATEMENT', padX + 16, y + 22);

    ctx.fillStyle = '#64748B';
    ctx.font = '500 12px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'right';
    ctx.fillText(`Statement Date: ${formattedDate}`, padX + contentWidth - 16, y + 22);

    // Party Name & City
    ctx.fillStyle = '#0F172A';
    ctx.font = 'bold 18px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'left';
    const partyNameText = party.displayName || 'Customer';
    ctx.fillText(partyNameText, padX + 16, y + 50);

    // Companies Subtitle
    const hasMed = medRows > 0;
    const hasSurg = surgRows > 0;
    let companiesText = '';
    if (hasMed && hasSurg) companiesText = 'Rahul Medical & Surgicals  •  Rahul Surgical';
    else if (hasMed) companiesText = 'Rahul Medical & Surgicals';
    else if (hasSurg) companiesText = 'Rahul Surgical';

    ctx.fillStyle = '#475569';
    ctx.font = '500 12px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.fillText(companiesText, padX + 16, y + 72);

    // Total Outstanding on Header Right
    ctx.fillStyle = '#1E3A8A';
    ctx.font = 'bold 20px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'right';
    ctx.fillText(`₹${formatIndianNumber(party.totalAmount || 0)}`, padX + contentWidth - 16, y + 52);

    ctx.fillStyle = '#64748B';
    ctx.font = '500 11px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.fillText('Total Outstanding', padX + contentWidth - 16, y + 70);
    ctx.textAlign = 'left';

    y += headerHeight + sectionGap;

    // --- Table Column Layout (6 columns, Type removed) ---
    const cols = [
      { title: 'Bill', width: 0.20, align: 'left' },
      { title: 'Date', width: 0.15, align: 'center' },
      { title: 'Debit (₹)', width: 0.17, align: 'right' },
      { title: 'Credit (₹)', width: 0.17, align: 'right' },
      { title: 'Balance (₹)', width: 0.19, align: 'right' },
      { title: 'Days', width: 0.12, align: 'center' },
    ];

    let curX = padX;
    cols.forEach(col => {
      col.x = curX;
      col.colWidth = col.width * contentWidth;
      curX += col.colWidth;
    });

    const drawTable = (sectionTitle, bannerBg, bannerTextColor, transactions) => {
      // 1. Section Banner (BOLD and UPPERCASE)
      ctx.fillStyle = bannerBg;
      ctx.fillRect(padX, y, contentWidth, bannerHeight);
      ctx.strokeStyle = '#94A3B8';
      ctx.lineWidth = 0.5;
      ctx.strokeRect(padX, y, contentWidth, bannerHeight);

      ctx.fillStyle = bannerTextColor;
      ctx.font = 'bold 13px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText(sectionTitle, padX + contentWidth / 2, y + 19);

      y += bannerHeight;

      // 2. Table Column Headers
      ctx.fillStyle = '#F1F5F9';
      ctx.fillRect(padX, y, contentWidth, tableHeaderHeight);
      ctx.strokeStyle = '#CBD5E1';
      ctx.lineWidth = 0.5;
      ctx.strokeRect(padX, y, contentWidth, tableHeaderHeight);

      ctx.fillStyle = '#334155';
      ctx.font = 'bold 12px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';

      cols.forEach(col => {
        let textX = col.x + 8;
        if (col.align === 'center') textX = col.x + col.colWidth / 2;
        if (col.align === 'right') textX = col.x + col.colWidth - 8;
        ctx.textAlign = col.align;
        ctx.fillText(col.title, textX, y + 19);
      });

      y += tableHeaderHeight;

      // 3. Transactions Rows
      transactions.forEach((tx, idx) => {
        const isEven = idx % 2 === 0;
        ctx.fillStyle = isEven ? '#FFFFFF' : '#F8FAFC';
        ctx.fillRect(padX, y, contentWidth, rowHeight);
        ctx.strokeStyle = '#E2E8F0';
        ctx.lineWidth = 0.5;
        ctx.strokeRect(padX, y, contentWidth, rowHeight);

        const valA = String(tx[0] || '').trim();
        const isRcpt = valA.toLowerCase().startsWith('rcpt');

        if (isRcpt) {
          ctx.fillStyle = '#EFF6FF';
          ctx.fillRect(padX + 1, y + 1, contentWidth - 2, rowHeight - 2);
          ctx.fillStyle = '#1D4ED8';
          ctx.font = 'italic 12px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
          ctx.textAlign = 'left';
          ctx.fillText(`📄  ${valA}`, padX + 12, y + 18);
        } else {
          const is6Col = tx.length === 6 || (tx[0] && tx[0] !== "");
          const bill = String((is6Col ? tx[0] : tx[1]) || '').trim();
          const date = String((is6Col ? tx[1] : tx[2]) || '').trim();
          const rawDebit = is6Col ? tx[2] : tx[3];
          const rawCredit = is6Col ? tx[3] : tx[4];
          const rawBalance = is6Col ? tx[4] : tx[5];
          const days = String((is6Col ? tx[5] : tx[6]) || '').trim();

          const debitVal = rawDebit !== '' && rawDebit !== null && rawDebit !== undefined && !isNaN(rawDebit) ? Number(rawDebit) : null;
          const creditVal = rawCredit !== '' && rawCredit !== null && rawCredit !== undefined && !isNaN(rawCredit) ? Number(rawCredit) : null;
          const balanceVal = rawBalance !== '' && rawBalance !== null && rawBalance !== undefined && !isNaN(rawBalance) ? Number(rawBalance) : null;

          // Bill
          ctx.fillStyle = '#0F172A';
          ctx.font = bill.startsWith('*') ? 'bold 12px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' : '12px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
          ctx.textAlign = 'left';
          ctx.fillText(bill, cols[0].x + 8, y + 18);

          // Date
          ctx.fillStyle = '#475569';
          ctx.font = '12px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
          ctx.textAlign = 'center';
          ctx.fillText(date, cols[1].x + cols[1].colWidth / 2, y + 18);

          // Debit
          ctx.fillStyle = '#0F172A';
          ctx.textAlign = 'right';
          ctx.fillText(debitVal !== null ? formatIndianNumber(debitVal) : '', cols[2].x + cols[2].colWidth - 8, y + 18);

          // Credit
          ctx.fillStyle = creditVal && creditVal > 0 ? '#15803D' : '#0F172A';
          ctx.fillText(creditVal !== null ? formatIndianNumber(creditVal) : '', cols[3].x + cols[3].colWidth - 8, y + 18);

          // Balance
          ctx.fillStyle = '#0F172A';
          ctx.font = 'bold 12px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
          ctx.fillText(balanceVal !== null ? formatIndianNumber(balanceVal) : '', cols[4].x + cols[4].colWidth - 8, y + 18);

          // Days
          ctx.fillStyle = '#64748B';
          ctx.font = '12px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
          ctx.textAlign = 'center';
          ctx.fillText(days, cols[5].x + cols[5].colWidth / 2, y + 18);
        }

        y += rowHeight;
      });

      y += sectionGap;
    };

    if (medRows > 0) {
      drawTable('--- RAHUL MEDICAL & SURGICALS ---', '#C6E0B4', '#1B4D1B', party.medicalTransactions);
    }
    if (surgRows > 0) {
      drawTable('--- RAHUL SURGICAL ---', '#F8CBAD', '#7C2D12', party.surgicalTransactions);
    }

    // --- 3. Summary Footer Bar ---
    ctx.fillStyle = '#4472C4';
    ctx.fillRect(padX, y, contentWidth, summaryHeight);
    ctx.strokeStyle = '#2B579A';
    ctx.lineWidth = 1;
    ctx.strokeRect(padX, y, contentWidth, summaryHeight);

    ctx.fillStyle = '#FFFFFF';
    ctx.font = 'bold 13px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'center';

    const medStr = `Medical: ₹${formatIndianNumber(party.medicalTotal || 0)}`;
    const surgStr = `Surgical: ₹${formatIndianNumber(party.surgicalTotal || 0)}`;
    const totalStr = `Total: ₹${formatIndianNumber(party.totalAmount || 0)}`;
    ctx.fillText(`Outstanding Summary   |   ${medStr}   |   ${surgStr}   |   ${totalStr}`, padX + contentWidth / 2, y + 23);

    canvas.toBlob((blob) => {
      resolve(blob);
    }, 'image/png');
  });
};


// ============================================================================
// App Component
// ============================================================================
const App = () => {
  // --- Application State ---
  const [isWorkerReady, setIsWorkerReady] = useState(false);
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

  // --- Raw Merged Data & ZIP Export State ---
  const [parsedRows, setParsedRows] = useState(null);
  const [isGeneratingZip, setIsGeneratingZip] = useState(false);
  const [zipProgress, setZipProgress] = useState({ current: 0, total: 0 });

  // --- Computed Values ---
  const currentStep = useMemo(() => {
    if (processedFileUrl) return 3;
    if (files.file1.file && files.file2.file && files.file1.label && files.file2.label && files.file1.label !== files.file2.label) return 2;
    return 1;
  }, [files, processedFileUrl]);

  // ----------------------------------------------------------------------
  // 1. Initialize Pyodide via Background Web Worker
  // ----------------------------------------------------------------------
  const doInitPyodide = async () => {
    try {
      setInitError(null);
      setLoading(true);
      setLoadingStep(0);

      await initPyodideWorker(
        (step) => setLoadingStep(step),
        (msg) => addLog(msg)
      );

      setIsWorkerReady(true);
      setLoading(false);
      addLog("Python Web Worker ready. Ready to process files.");
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

    // Auto-detect label via Web Worker
    if (isWorkerReady) {
      try {
        const detectedLabel = await autoDetectLabelWorker(uploadedFile);
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
    setParsedRows(null);
    setIsGeneratingZip(false);
    setZipProgress({ current: 0, total: 0 });
    setLogs([]);
    setError(null);
    setWarningCount(0);
    setShowLogs(false);
    setShowPreview(true);
    addLog("App reset. Ready for new files.");
  };

  // ----------------------------------------------------------------------
  // 3. PDF Generation (Single Page — preserved, dynamically loaded)
  // ----------------------------------------------------------------------
  const generatePDF = async (data) => {
    addLog("Generating formatted single-page PDF document...");
    const { jsPDF } = await import('jspdf');
    const autoTableModule = await import('jspdf-autotable');
    const autoTable = autoTableModule.default || autoTableModule;

    const body = [];

    data.forEach(row => {
      const valA = String(row[0] || '').trim();
      const valB = String(row[1] || '').trim();

      // Handle empty spacing rows between parties
      if (row.every(c => c === "" || c === null || c === undefined)) {
        body.push([{ content: '', colSpan: 6, styles: { minCellHeight: 14, fillColor: [255, 255, 255], lineWidth: 0 } }]);
        return;
      }

      const isPartyHeader = (valA && !valB &&
        !valA.toLowerCase().startsWith("rcpt") &&
        !valA.startsWith("Outstanding Summary") &&
        !valA.startsWith("---") &&
        !valA.startsWith("--") &&
        valA !== "Bill" &&
        valA !== "Type");

      if (isPartyHeader) {
        const partyTotal = row[5] !== "" && row[5] !== undefined ? row[5] : (row[6] || 0);
        body.push([
          { content: valA, colSpan: 5, styles: { fillColor: [217, 225, 242], fontStyle: 'bold', halign: 'center', lineWidth: 0.5, lineColor: [180, 195, 225] } },
          { content: formatIndianNumber(partyTotal), styles: { fillColor: [217, 225, 242], fontStyle: 'bold', halign: 'right', lineWidth: 0.5, lineColor: [180, 195, 225] } }
        ]);
      } else if (valA === "Bill" || (valA === "Type" && valB === "Bill")) {
        body.push([
          { content: 'Bill', styles: { fillColor: [242, 242, 242], fontStyle: 'bold', halign: 'center', lineWidth: 0.5 } },
          { content: 'Date', styles: { fillColor: [242, 242, 242], fontStyle: 'bold', halign: 'center', lineWidth: 0.5 } },
          { content: 'Debit', styles: { fillColor: [242, 242, 242], fontStyle: 'bold', halign: 'center', lineWidth: 0.5 } },
          { content: 'Credit', styles: { fillColor: [242, 242, 242], fontStyle: 'bold', halign: 'center', lineWidth: 0.5 } },
          { content: 'Balance', styles: { fillColor: [242, 242, 242], fontStyle: 'bold', halign: 'center', lineWidth: 0.5 } },
          { content: 'Days', styles: { fillColor: [242, 242, 242], fontStyle: 'bold', halign: 'center', lineWidth: 0.5 } },
        ]);
      } else if (valA.startsWith("---") || valA.startsWith("--")) {
        const isMed = valA.toLowerCase().includes("medical");
        const fill = isMed ? [198, 224, 180] : [248, 203, 173];
        const text = isMed ? "-- Medical --" : "-- Surgical --";
        body.push([{ content: text, colSpan: 6, styles: { fillColor: fill, fontStyle: 'bold', halign: 'center', lineWidth: 0.5, minCellHeight: 14 } }]);
      } else if (valA.startsWith("Outstanding Summary")) {
        body.push([{ content: valA, colSpan: 6, styles: { fillColor: [68, 114, 196], textColor: [255, 255, 255], fontStyle: 'bold', halign: 'center', lineWidth: 0.5, minCellHeight: 16 } }]);
      } else if (valA.toLowerCase().startsWith("rcpt")) {
        body.push([{ content: valA, colSpan: 6, styles: { fontStyle: 'italic', halign: 'left', fontSize: 7, fillColor: [248, 250, 252], lineWidth: 0.5, lineColor: [200, 200, 200] } }]);
      } else {
        const is6Col = row.length === 6 || (row[0] && row[0] !== "");
        const bill = String((is6Col ? row[0] : row[1]) || '');
        const date = String((is6Col ? row[1] : row[2]) || '');
        const debit = is6Col ? row[2] : row[3];
        const credit = is6Col ? row[3] : row[4];
        const balance = is6Col ? row[4] : row[5];
        const days = String((is6Col ? row[5] : row[6]) || '');

        const isBoldRow = bill.startsWith('*');
        body.push([
          { content: bill, styles: { fontStyle: isBoldRow ? 'bold' : 'normal', halign: 'center', lineWidth: 0.5, lineColor: [200, 200, 200] } },
          { content: date, styles: { fontStyle: 'normal', halign: 'center', lineWidth: 0.5, lineColor: [200, 200, 200] } },
          { content: formatIndianNumber(debit), styles: { fontStyle: 'normal', halign: 'right', lineWidth: 0.5, lineColor: [200, 200, 200] } },
          { content: formatIndianNumber(credit), styles: { fontStyle: 'normal', halign: 'right', lineWidth: 0.5, lineColor: [200, 200, 200] } },
          { content: formatIndianNumber(balance), styles: { fontStyle: isBoldRow ? 'bold' : 'normal', halign: 'right', lineWidth: 0.5, lineColor: [200, 200, 200] } },
          { content: days, styles: { fontStyle: 'normal', halign: 'center', lineWidth: 0.5, lineColor: [200, 200, 200] } },
        ]);
      }
    });

    const A4_WIDTH_PT = 595.28;
    const TOP_MARGIN = 35;
    const MARGIN_X = 25;
    const TABLE_WIDTH = A4_WIDTH_PT - MARGIN_X * 2; // 545.28 pt

    // 6-Column layout across 545.28 pt:
    const colStyles = {
      0: { cellWidth: 95, halign: 'center' },   // Bill (e.g. *RM-4498)
      1: { cellWidth: 80, halign: 'center' },   // Date (e.g. 06-03-26)
      2: { cellWidth: 95, halign: 'right' },    // Debit
      3: { cellWidth: 85, halign: 'right' },    // Credit
      4: { cellWidth: 105, halign: 'right' },   // Balance
      5: { cellWidth: 85.28, halign: 'center' }, // Days
    };

    // Calculate dynamic single-page height safely without multi-page break
    const maxSafeHeight = Math.max(841.89, body.length * 35 + 300);
    const dummyDoc = new jsPDF({
      orientation: 'portrait',
      unit: 'pt',
      format: [A4_WIDTH_PT, maxSafeHeight]
    });

    autoTable(dummyDoc, {
      body: body,
      theme: 'grid',
      styles: { fontSize: 7.5, cellPadding: 2.5, textColor: [0, 0, 0], font: 'helvetica' },
      columnStyles: colStyles,
      margin: { top: TOP_MARGIN, left: MARGIN_X, right: MARGIN_X, bottom: 0 },
      tableWidth: TABLE_WIDTH,
    });

    const finalPageHeight = Math.max(841.89, dummyDoc.lastAutoTable.finalY + 40);

    // Second pass: Render exact single-page document
    const finalDoc = new jsPDF({
      orientation: 'portrait',
      unit: 'pt',
      format: [A4_WIDTH_PT, finalPageHeight]
    });

    // Add centered title
    finalDoc.setFontSize(12);
    finalDoc.setFont('helvetica', 'bold');
    finalDoc.setTextColor(40, 40, 40);
    finalDoc.text(`Outstanding Report  —  ${getDisplayDate()}`, A4_WIDTH_PT / 2, 22, { align: 'center' });

    autoTable(finalDoc, {
      body: body,
      theme: 'grid',
      styles: { fontSize: 7.5, cellPadding: 2.5, textColor: [0, 0, 0], font: 'helvetica' },
      columnStyles: colStyles,
      margin: { top: TOP_MARGIN, left: MARGIN_X, right: MARGIN_X, bottom: 0 },
      tableWidth: TABLE_WIDTH,
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

            if is_rcpt_row:
                master_data[current_party_key][f'transactions_{file_label}'].append([col_a_val, "", "", "", "", ""])
            else:
                bill_no = str(row_data[1]).strip()
                date_val = str(row_data[2]).strip()
                debit_val = row_data[3] if len(row_data) > 3 else ""
                credit_val = row_data[4] if len(row_data) > 4 else ""
                bal_val = row_data[5] if len(row_data) > 5 else ""
                days_val = str(row_data[6]).strip() if len(row_data) > 6 else ""
                master_data[current_party_key][f'transactions_{file_label}'].append([bill_no, date_val, debit_val, credit_val, bal_val, days_val])

def style_excel_file(filename):
    wb = load_workbook(filename)
    ws = wb.active
    ws.title = "Outstanding Report"
    
    # Ensure grid lines are visible
    ws.views.sheetView[0].showGridLines = True

    # Fonts
    party_header_font = Font(name="Calibri", bold=True, size=11, color="000000")
    sub_header_font = Font(name="Calibri", bold=True, size=10, color="000000")
    section_font = Font(name="Calibri", bold=True, size=10, color="000000")
    summary_font = Font(name="Calibri", bold=True, size=10.5, color="FFFFFF")
    regular_font = Font(name="Calibri", size=10)
    bold_regular_font = Font(name="Calibri", bold=True, size=10)
    rcpt_font = Font(name="Calibri", italic=True, size=9.5, color="333333")

    # Fills
    party_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    sub_header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    medical_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    surgical_fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
    summary_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")

    # Borders
    thin = Side(style='thin', color="B0B0B0")
    no_side = Side(border_style=None)
    thin_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    indian_num_format = '#,##,##0.00'

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        val_a = str(row[0].value).strip() if row[0].value else ""
        val_b = str(row[1].value).strip() if len(row) > 1 and row[1].value else ""

        # Blank spacer row
        if not any(c.value for c in row):
            ws.row_dimensions[row[0].row].height = 14
            for cell in row:
                cell.border = Border(left=no_side, right=no_side, top=no_side, bottom=no_side)
            continue

        is_party_header = (val_a and not val_b and not val_a.lower().startswith("rcpt") and not val_a.startswith("Outstanding Summary") and not val_a.startswith("---") and not val_a.startswith("--") and val_a != "Bill")

        if is_party_header:
            ws.row_dimensions[row[0].row].height = 24
            if len(row) >= 6:
                for i in range(6):
                    row[i].fill = party_fill
                    row[i].border = Border(top=thin, left=thin if i in (0, 5) else no_side, right=thin if i in (4, 5) else no_side, bottom=thin)
                row[0].font, row[0].alignment = party_header_font, Alignment(horizontal='center', vertical='center')
                row[5].font, row[5].alignment = party_header_font, Alignment(horizontal='right', vertical='center')
                if isinstance(row[5].value, (int, float)):
                    row[5].number_format = indian_num_format
                ws.merge_cells(start_row=row[0].row, start_column=1, end_row=row[0].row, end_column=5)

        elif val_a.startswith("---") or val_a.startswith("--"):
            ws.row_dimensions[row[0].row].height = 20
            if len(row) >= 6:
                ws.merge_cells(start_row=row[0].row, start_column=1, end_row=row[0].row, end_column=6)
                row[0].font, row[0].alignment = section_font, Alignment(horizontal='center', vertical='center')
                fill = medical_fill if "medical" in val_a.lower() else surgical_fill
                for cell in row:
                    cell.fill = fill
                    cell.border = thin_border

        elif val_a == "Bill":
            ws.row_dimensions[row[0].row].height = 20
            if len(row) >= 6:
                for cell in row:
                    cell.font, cell.fill, cell.border = sub_header_font, sub_header_fill, thin_border
                    cell.alignment = Alignment(horizontal='center', vertical='center')

        elif val_a.startswith("Outstanding Summary"):
            ws.row_dimensions[row[0].row].height = 22
            if len(row) >= 6:
                for cell in row:
                    cell.fill, cell.border = summary_fill, thin_border
                row[0].font, row[0].alignment = summary_font, Alignment(horizontal='center', vertical='center')
                ws.merge_cells(start_row=row[0].row, start_column=1, end_row=row[0].row, end_column=6)

        elif val_a.lower().startswith("rcpt"):
            ws.row_dimensions[row[0].row].height = 19
            if len(row) >= 6:
                ws.merge_cells(start_row=row[0].row, start_column=1, end_row=row[0].row, end_column=6)
                for cell in row:
                    cell.border = thin_border
                row[0].font = rcpt_font
                row[0].alignment = Alignment(horizontal='left', vertical='center')

        else:
            ws.row_dimensions[row[0].row].height = 19
            is_bold_bill = val_a.startswith('*')
            for idx_c, cell in enumerate(row):
                cell.border = thin_border
                if idx_c in (0, 1, 5):
                    cell.font = bold_regular_font if (idx_c == 0 and is_bold_bill) else regular_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                elif idx_c in (2, 3, 4):
                    cell.font = bold_regular_font if (idx_c == 4 and is_bold_bill) else regular_font
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = indian_num_format

    # Set polished column widths (6 columns)
    column_widths = {
        'A': 20, # Bill No
        'B': 14, # Date
        'C': 16, # Debit
        'D': 14, # Credit
        'E': 16, # Balance
        'F': 12, # Days
    }
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    wb.save(filename)

# Execution
combined_data = {}
if os.path.exists(FILE_1): process_file(FILE_1, combined_data, 'file1')
if os.path.exists(FILE_2): process_file(FILE_2, combined_data, 'file2')

final_output_rows = []
transaction_header = ["Bill", "Date", "Debit", "Credit", "Balance", "Days"]

for key, data in combined_data.items():
    header_row = [f"{data['original_name']}  ({data['city']})", "", "", "", "", data['total_combined']]
    final_output_rows.append(header_row)
    final_output_rows.append(transaction_header)

    if data['transactions_file1']:
        final_output_rows.append(["--- Medical ---", "", "", "", "", ""])
        for trans in data['transactions_file1']:
            padded = list(trans)
            while len(padded) < 6: padded.append("")
            final_output_rows.append(padded[:6])

    if data['transactions_file2']:
        final_output_rows.append(["--- Surgical ---", "", "", "", "", ""])
        for trans in data['transactions_file2']:
            padded = list(trans)
            while len(padded) < 6: padded.append("")
            final_output_rows.append(padded[:6])

    t1, t2, total = data['totals_breakdown']['file1'], data['totals_breakdown']['file2'], data['total_combined']
    final_output_rows.append([f"Outstanding Summary   |   Medical: {format_indian(t1)}   |   Surgical: {format_indian(t2)}   |   Total: {format_indian(total)}", "", "", "", "", ""])
    final_output_rows.append([""] * 6)

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
    if (!isWorkerReady) { setError("Python environment is not ready."); return; }

    setIsProcessing(true);
    setLogs([]);
    setError(null);
    setWarningCount(0);
    setPreviewData(null);
    setParsedRows(null);
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
      const f1Ext = file1.file.name.split('.').pop();
      const f2Ext = file2.file.name.split('.').pop();

      const f1Name = file1.label === "RAHUL MEDICAL & SURGICAL" ? `medical_input.${f1Ext}` : `surgical_input.${f1Ext}`;
      const f2Name = file2.label === "RAHUL MEDICAL & SURGICAL" ? `medical_input.${f2Ext}` : `surgical_input.${f2Ext}`;

      const medicalVfsName = file1.label === "RAHUL MEDICAL & SURGICAL" ? f1Name : f2Name;
      const surgicalVfsName = file1.label === "RAHUL MEDICAL & SURGICAL" ? f2Name : f1Name;

      const f1Buffer = await file1.file.arrayBuffer();
      const f2Buffer = await file2.file.arrayBuffer();

      const workerFiles = [
        { name: f1Name, buffer: f1Buffer },
        { name: f2Name, buffer: f2Buffer }
      ];

      addLog(`Saved "${f1Name}" and "${f2Name}" in worker memory.`);
      addLog("Executing Python logic in background worker...");

      const finalScript = getPythonScript(medicalVfsName, surgicalVfsName);

      // Execute Python logic in Web Worker
      const { excelBuffer, jsonResult } = await runMergeWorker(workerFiles, finalScript);
      const parsedData = jsonResult.data;
      const warnCount = jsonResult.warnings || 0;
      const warnDetails = jsonResult.warning_details || [];

      // Save raw rows for ZIP party card generation
      setParsedRows(parsedData);

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
      const pdfDoc = await generatePDF(parsedData);
      const pdfBlob = pdfDoc.output('blob');
      setProcessedPdfUrl(URL.createObjectURL(pdfBlob));

      // Fetch Excel
      if (excelBuffer) {
        const blob = new Blob([excelBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
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

  // ----------------------------------------------------------------------
  // 6. Logic: Generate Party Image Cards and Download ZIP
  // ----------------------------------------------------------------------
  const handleDownloadZip = async () => {
    if (!parsedRows || parsedRows.length === 0) {
      setError("No processed party data available. Please merge files first.");
      return;
    }

    setIsGeneratingZip(true);
    setZipProgress({ current: 0, total: 0 });
    addLog("Starting individual party card generation...");

    try {
      const JSZipModule = await import('jszip');
      const JSZip = JSZipModule.default || JSZipModule;
      const parties = groupRowsByParty(parsedRows);
      if (parties.length === 0) {
        setError("No parties found to generate images.");
        setIsGeneratingZip(false);
        return;
      }

      const zip = new JSZip();
      const formattedDate = getDisplayDate();
      const totalParties = parties.length;
      setZipProgress({ current: 0, total: totalParties });

      for (let i = 0; i < totalParties; i++) {
        const party = parties[i];
        setZipProgress({ current: i + 1, total: totalParties });

        // Yield thread so React can render progress bar
        await new Promise(resolve => setTimeout(resolve, 0));

        const blob = await renderPartyCardCanvas(party, formattedDate);
        const safeName = sanitizeFilename(party.displayName);
        zip.file(`${safeName}.png`, blob);
      }

      addLog(`Packing ${totalParties} party cards into ZIP archive...`);
      const zipBlob = await zip.generateAsync({
        type: "blob",
        compression: "DEFLATE",
        compressionOptions: { level: 6 }
      });

      const zipFilename = `Party_Wise_Reports_${getDateString()}.zip`;
      const url = URL.createObjectURL(zipBlob);
      const link = document.createElement('a');
      link.href = url;
      link.download = zipFilename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      setTimeout(() => URL.revokeObjectURL(url), 10000);

      addLog(`Successfully downloaded "${zipFilename}"!`);
    } catch (err) {
      console.error("ZIP Generation error:", err);
      setError(`Failed to generate images ZIP: ${err.message}`);
      addLog(`[ERROR] Image ZIP: ${err.message}`);
    } finally {
      setIsGeneratingZip(false);
    }
  };

  // --- Output filename with date ---
  const outputBaseName = `Merged_Report_${getDateString()}`;

  // ====================================================================
  // RENDER
  // ====================================================================
  return (
    <main id="main-content" className="min-h-screen bg-slate-50 text-slate-800 font-sans p-4 sm:p-6">
      <div className="max-w-3xl mx-auto bg-white shadow-xl rounded-xl overflow-hidden border border-slate-200">

        {/* ── Header ── */}
        <header className="bg-blue-600 p-6 text-white">
          <h1 className="text-2xl font-bold flex items-center gap-2">
            <FileSpreadsheet aria-hidden="true" /> Excel Merge Tool
          </h1>
          <p className="text-blue-50 mt-2 text-sm font-normal">
            Securely merge and export Surgical and Medical reports to Excel, PDF, and image cards.
          </p>
        </header>

        {/* ── Progress Stepper ── */}
        {!loading && (
          <nav className="px-6 pt-5 pb-1" aria-label="Progress steps">
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
                        'bg-slate-200 text-slate-700'
                      }`}>
                        {isCompleted ? <CheckCircle size={16} aria-hidden="true" /> : step.num}
                      </div>
                      <span className={`text-xs font-medium ${isCurrent ? 'text-blue-600' : isCompleted ? 'text-green-600' : 'text-slate-600'}`}>
                        {step.label}
                      </span>
                    </div>
                  </React.Fragment>
                );
              })}
            </div>
          </nav>
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
                <p className="text-slate-700 font-medium text-sm">{LOADING_STEPS[loadingStep]}</p>
                <p className="text-slate-600 text-xs mt-2">This usually takes 10–20 seconds on first visit</p>
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
                          <div className="flex flex-col items-center text-slate-600 group-hover:text-blue-600">
                            <Upload size={24} className="mb-2" aria-hidden="true" />
                            <span className="text-xs font-semibold">Click or Drag Excel File</span>
                            <span className="text-xs text-slate-500 mt-1">Supports .xls and .xlsx</span>
                          </div>
                        )}
                      </div>
                    </div>

                    {/* File validation error */}
                    {fileErrors[key] && (
                      <p className="text-red-600 text-xs mb-2 flex items-center gap-1">
                        <AlertCircle size={12} aria-hidden="true" /> {fileErrors[key]}
                      </p>
                    )}

                    {/* Label selector */}
                    <label htmlFor={`label-${key}`} className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
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
                  disabled={isProcessing || !isWorkerReady || currentStep < 2}
                  className={`flex items-center gap-2 px-8 py-3 rounded-full font-bold shadow-lg transition-all ${
                    (isProcessing || !isWorkerReady || currentStep < 2)
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
                  <div className="flex flex-wrap items-center justify-center gap-3">
                    <a
                      href={processedFileUrl}
                      download={`${outputBaseName}.xlsx`}
                      className="flex items-center justify-center gap-2 px-5 py-2.5 rounded-full font-bold shadow-md bg-green-600 hover:bg-green-700 text-white text-sm transition-all hover:-translate-y-0.5"
                    >
                      <FileSpreadsheet size={18} aria-hidden="true" /> Download Excel
                    </a>
                    <a
                      href={processedPdfUrl}
                      download={`${outputBaseName}.pdf`}
                      className="flex items-center justify-center gap-2 px-5 py-2.5 rounded-full font-bold shadow-md bg-red-500 hover:bg-red-600 text-white text-sm transition-all hover:-translate-y-0.5"
                    >
                      <FileText size={18} aria-hidden="true" /> Download PDF
                    </a>
                    <button
                      onClick={handleDownloadZip}
                      disabled={isGeneratingZip || !parsedRows}
                      className="flex items-center justify-center gap-2 px-5 py-2.5 rounded-full font-bold shadow-md bg-amber-600 hover:bg-amber-700 text-white text-sm transition-all hover:-translate-y-0.5 disabled:bg-amber-300 disabled:cursor-not-allowed"
                      aria-label="Download party report images as ZIP"
                    >
                      {isGeneratingZip ? (
                        <>
                          <RefreshCw className="animate-spin" size={18} aria-hidden="true" />
                          <span>Generating ({zipProgress.current}/{zipProgress.total})...</span>
                        </>
                      ) : (
                        <>
                          <FolderArchive size={18} aria-hidden="true" />
                          <span>Download Images (ZIP)</span>
                        </>
                      )}
                    </button>
                    <button onClick={resetApp} className="flex items-center justify-center gap-2 px-4 py-2.5 rounded-full font-bold border-2 border-slate-200 text-slate-600 text-sm hover:bg-slate-50 transition-all">
                      <RotateCcw size={16} aria-hidden="true" /> Start Over
                    </button>
                  </div>

                  {/* ── ZIP Generation Progress Bar ── */}
                  {isGeneratingZip && (
                    <div className="w-full max-w-md bg-amber-50 border border-amber-200 rounded-lg p-3 text-center">
                      <div className="flex justify-between text-xs text-amber-900 font-medium mb-1.5">
                        <span>Rendering party card images...</span>
                        <span>{zipProgress.current} of {zipProgress.total}</span>
                      </div>
                      <div className="w-full bg-amber-200 rounded-full h-2 overflow-hidden">
                        <div
                          className="bg-amber-600 h-2 rounded-full transition-all duration-100"
                          style={{ width: `${zipProgress.total > 0 ? Math.round((zipProgress.current / zipProgress.total) * 100) : 0}%` }}
                        />
                      </div>
                    </div>
                  )}
                </>
              )}
            </div>

            {/* ── System Logs ── */}
            <div className="mt-2">
              <button
                onClick={() => setShowLogs(!showLogs)}
                className="flex items-center gap-2 text-xs font-semibold text-slate-600 hover:text-slate-900 transition-colors mx-auto p-1 rounded focus:outline-none focus:ring-2 focus:ring-blue-500"
                aria-expanded={showLogs}
                aria-controls="log-panel"
              >
                <Terminal size={14} aria-hidden="true" /> {showLogs ? "Hide System Logs" : "Show System Logs"} {showLogs ? <ChevronUp size={14} aria-hidden="true" /> : <ChevronDown size={14} aria-hidden="true" />}
              </button>
              {showLogs && (
                <div id="log-panel" role="log" aria-live="polite" className="bg-slate-900 rounded-lg p-4 font-mono text-xs text-green-400 h-48 overflow-y-auto mt-3 shadow-inner">
                  <div className="text-slate-400 border-b border-slate-800 pb-2 mb-2 flex justify-between"><span>System Logs</span></div>
                  {logs.length === 0 && <span className="text-slate-500 italic">Waiting for input...</span>}
                  {logs.map((log, i) => <div key={i} className="mb-1 leading-relaxed">{log}</div>)}
                </div>
              )}
            </div>

          </div>
        )}
      </div>
    </main>
  );
};

export default App;