import React, { useState, useEffect, useRef } from 'react';
import { Upload, FileSpreadsheet, Download, RefreshCw, AlertCircle, CheckCircle, ChevronDown, ChevronUp, Terminal, RotateCcw, FileText } from 'lucide-react';
import { jsPDF } from 'https://esm.sh/jspdf@2.5.1';
import autoTable from 'https://esm.sh/jspdf-autotable@3.8.2';

/**
 * INSTRUCTIONS FOR DEPLOYMENT:
 * 1. Ensure you add this script tag to your public/index.html <head>:
 * <script src="https://cdn.jsdelivr.net/pyodide/v0.23.4/full/pyodide.js"></script>
 * 2. npm install lucide-react jspdf jspdf-autotable
 */

const App = () => {
  // Application State
  const [pyodide, setPyodide] = useState(null);
  const [loading, setLoading] = useState(true);
  const [loadingText, setLoadingText] = useState("Initializing Python Environment...");
  const [logs, setLogs] = useState([]);
  const [showLogs, setShowLogs] = useState(false);
  
  const pyodideInitializing = useRef(false);
  
  // File State
  const [files, setFiles] = useState({
    file1: { file: null, label: null, name: null },
    file2: { file: null, label: null, name: null }
  });
  
  const [processedFileUrl, setProcessedFileUrl] = useState(null);
  const [processedPdfUrl, setProcessedPdfUrl] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState(null);

  const OPTIONS = ["RAHUL SURGICAL", "RAHUL MEDICAL & SURGICAL"];

  // ----------------------------------------------------------------------
  // 1. Initialize Pyodide
  // ----------------------------------------------------------------------
  useEffect(() => {
    if (pyodideInitializing.current) return;
    pyodideInitializing.current = true;

    const initPyodide = async () => {
      try {
        if (!window.loadPyodide) {
          throw new Error("Pyodide script not found. Please add the CDN to index.html");
        }
        const py = await window.loadPyodide();
        setLoadingText("Setting up package manager...");
        await py.loadPackage("micropip");
        const micropip = py.pyimport("micropip");
        setLoadingText("Installing Python libraries (pandas, openpyxl, xlrd)...");
        await py.loadPackage("pandas");
        await micropip.install("xlrd");
        await micropip.install("openpyxl");
        setPyodide(py);
        setLoading(false);
        addLog("Python environment ready. Ready to process files.");
      } catch (err) {
        setError(`Failed to load Python: ${err.message}`);
        setLoading(false);
        pyodideInitializing.current = false;
      }
    };
    initPyodide();
  }, []);

  const addLog = (msg) => setLogs(prev => [...prev, `[${new Date().toLocaleTimeString()}] ${msg}`]);

  // ----------------------------------------------------------------------
  // 2. File Handling & Reset
  // ----------------------------------------------------------------------
  const handleFileUpload = (e, fileKey) => {
    const uploadedFile = e.target.files[0];
    if (!uploadedFile) return;
    setFiles(prev => ({ ...prev, [fileKey]: { ...prev[fileKey], file: uploadedFile, name: uploadedFile.name } }));
    
    if (processedFileUrl) URL.revokeObjectURL(processedFileUrl);
    if (processedPdfUrl) URL.revokeObjectURL(processedPdfUrl);
    setProcessedFileUrl(null); 
    setProcessedPdfUrl(null);
    setError(null);
  };

  const handleLabelChange = (val, fileKey) => {
    setFiles(prev => ({ ...prev, [fileKey]: { ...prev[fileKey], label: val } }));
  };

  const resetApp = () => {
    setFiles({
      file1: { file: null, label: null, name: null },
      file2: { file: null, label: null, name: null }
    });
    if (processedFileUrl) URL.revokeObjectURL(processedFileUrl);
    if (processedPdfUrl) URL.revokeObjectURL(processedPdfUrl);
    setProcessedFileUrl(null);
    setProcessedPdfUrl(null);
    setLogs([]);
    setError(null);
    setShowLogs(false);
    addLog("App reset. Ready for new files.");
  };

  // ----------------------------------------------------------------------
  // 3. JS PDF Generation (Single Page Logic)
  // ----------------------------------------------------------------------
  const generatePDF = (data) => {
    addLog("Generating formatted single-page PDF document...");
    const body = [];

    data.forEach(row => {
      const valA = String(row[0] || '').trim();
      const valB = String(row[1] || '').trim();
      
      // Handle empty spacing rows
      if (row.every(c => c === "" || c === null)) {
        body.push([{ content: '', colSpan: 7, styles: { minCellHeight: 15, fillColor: [255,255,255], lineWidth: 0 } }]);
        return;
      }

      const isPartyHeader = (valA && !valB && !valA.toLowerCase().startsWith("rcpt") && !valA.startsWith("Outstanding Summary") && !valA.startsWith("---") && valA !== "Type");

      if (isPartyHeader) {
        body.push([
          { content: valA, colSpan: 6, styles: { fillColor: [217, 225, 242], fontStyle: 'bold', halign: 'center', lineWidth: 0.5 } },
          { content: String(row[6]), styles: { fillColor: [217, 225, 242], fontStyle: 'bold', halign: 'right', lineWidth: 0.5 } }
        ]);
      } else if (valA === "Type") {
        body.push(row.map(cell => ({ content: String(cell), styles: { fillColor: [242, 242, 242], fontStyle: 'bold', halign: 'center', lineWidth: 0.5 } })));
      } else if (valA.startsWith("---")) {
        const fill = valA.includes("Medical") ? [198, 224, 180] : [248, 203, 173];
        body.push([{ content: valA, colSpan: 7, styles: { fillColor: fill, fontStyle: 'bold', halign: 'center', lineWidth: 0.5 } }]);
      } else if (valA.startsWith("Outstanding Summary")) {
        body.push([{ content: valA, colSpan: 7, styles: { fillColor: [68, 114, 196], textColor: [255, 255, 255], fontStyle: 'bold', halign: 'center', lineWidth: 0.5 } }]);
      } else {
        // Normal data rows
        body.push(row.map((cell, i) => ({ 
          content: cell !== "" && cell !== null ? String(cell) : "", 
          styles: { halign: (i > 2 && cell !== "") ? 'right' : 'left', lineWidth: 0.1, lineColor: [200, 200, 200] } 
        })));
      }
    });

    // ---------------------------------------------------------
    // DYNAMIC HEIGHT CALCULATION:
    // 1. Draw table on a fake, infinitely tall document
    // ---------------------------------------------------------
    const A4_WIDTH_PT = 595.28;
    const dummyDoc = new jsPDF('p', 'pt', [A4_WIDTH_PT, 99999]); 
    
    autoTable(dummyDoc, {
      body: body,
      theme: 'grid',
      styles: { fontSize: 8, cellPadding: 3, font: 'helvetica' },
      margin: { top: 30, left: 30, right: 30 },
      tableWidth: 'auto',
    });

    // 2. Extract exactly how tall the table ended up being
    const totalContentHeight = dummyDoc.lastAutoTable.finalY + 30; // Add 30pt for bottom margin padding
    // Make sure the document is at least A4 height (841.89pt) so it doesn't look weird if the table is tiny
    const finalPageHeight = Math.max(841.89, totalContentHeight); 

    // ---------------------------------------------------------
    // FINAL PDF GENERATION:
    // Create the actual document using the exactly measured height
    // ---------------------------------------------------------
    const finalDoc = new jsPDF('p', 'pt', [A4_WIDTH_PT, finalPageHeight]);

    autoTable(finalDoc, {
      body: body,
      theme: 'grid',
      styles: { fontSize: 8, cellPadding: 3, textColor: [0, 0, 0], font: 'helvetica' },
      margin: { top: 30, left: 30, right: 30 },
      tableWidth: 'auto',
    });

    return finalDoc;
  };

  // ----------------------------------------------------------------------
  // 4. The Python Script (Template String)
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

def process_file(filepath, master_data, file_label):
    if not os.path.exists(filepath): return
    try:
        if filepath.lower().endswith('.xls'):
            df = pd.read_excel(filepath, header=None, engine='xlrd')
        else:
            df = pd.read_excel(filepath, header=None, engine='openpyxl')
    except Exception as e: return

    current_party_key = None
    for index, row in df.iterrows():
        row_data = ["" if pd.isna(x) else x for x in row.tolist()]
        if len(row_data) < 2: continue

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
            if any(str(x).strip() for x in row_data):
                if len(row_data) > 3:
                    val_d = row_data[3]
                    if isinstance(val_d, str) and " " in val_d.strip():
                        parts = val_d.strip().split()
                        if len(parts) == 2:
                            try:
                                debit, credit = float(parts[0]), float(parts[1])
                                row_data[3] = debit
                                row_data.insert(4, credit)
                            except ValueError: pass
                try:
                    if len(row_data) > 3:
                        current_debit = float(row_data[3])
                        if current_debit < 0:
                            row_data[3] = 0.0
                            if len(row_data) > 4: row_data[4] = abs(current_debit)
                            else: row_data.append(abs(current_debit))
                except (ValueError, TypeError): pass

                col_a_val = str(row_data[0]).strip()
                if not col_a_val.lower().startswith("rcpt"): row_data[0] = ""
                master_data[current_party_key][f'transactions_{file_label}'].append(row_data)

def style_excel_file(filename):
    wb = load_workbook(filename)
    ws = wb.active
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
                    for cell in row: cell.font, cell.border = regular_font, thin_border

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
            final_output_rows.append(padded)

    if data['transactions_file2']:
        final_output_rows.append(["--- Surgical ---", "", "", "", "", "", ""])
        for trans in data['transactions_file2']:
            padded = list(trans)
            while len(padded) < 7: padded.append("")
            final_output_rows.append(padded)

    t1, t2, total = data['totals_breakdown']['file1'], data['totals_breakdown']['file2'], data['total_combined']
    final_output_rows.append([f"Outstanding Summary   |   Medical: {t1}   |   Surgical: {t2}   |   Total: {total}", "", "", "", "", "", ""])
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

json.dumps(clean_rows)
`;

  // ----------------------------------------------------------------------
  // 5. Logic: Execute Pipeline
  // ----------------------------------------------------------------------
  const handleMerge = async () => {
    if (!pyodide) { setError("Python environment is not ready."); return; }

    setIsProcessing(true);
    setLogs([]);
    setError(null);
    if (processedFileUrl) URL.revokeObjectURL(processedFileUrl);
    if (processedPdfUrl) URL.revokeObjectURL(processedPdfUrl);
    setProcessedFileUrl(null);
    setProcessedPdfUrl(null);

    const { file1, file2 } = files;
    if (!file1.file || !file2.file) { setError("Please upload both files."); setIsProcessing(false); return; }
    if (!file1.label || !file2.label) { setError("Please select options for both files."); setIsProcessing(false); return; }
    if (file1.label === file2.label) { setError("Files must have different labels."); setIsProcessing(false); return; }

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
      const parsedData = JSON.parse(jsonResult);

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
      setError(err.message);
      addLog(`Error: ${err.message}`);
      setShowLogs(true);
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans p-6">
      <div className="max-w-3xl mx-auto bg-white shadow-xl rounded-xl overflow-hidden border border-slate-200">
        
        {/* Header */}
        <div className="bg-blue-600 p-6 text-white">
          <h1 className="text-2xl font-bold flex items-center gap-2">
            <FileSpreadsheet /> Excel Merge Tool
          </h1>
          <p className="text-blue-100 mt-2 text-sm">
            Securely merge and export Surgical and Medical reports to Excel and PDF.
          </p>
        </div>

        {loading ? (
          <div className="p-12 flex flex-col items-center justify-center text-center">
            <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mb-4"></div>
            <p className="text-slate-500 font-medium">{loadingText}</p>
          </div>
        ) : (
          <div className="p-6 space-y-8">
            {error && (
              <div className="bg-red-50 text-red-700 p-4 rounded-lg flex items-start gap-3 border border-red-200">
                <AlertCircle className="shrink-0 mt-0.5" size={20} />
                <div className="text-sm font-medium">{error}</div>
              </div>
            )}

            {/* Inputs */}
            <div className="grid md:grid-cols-2 gap-6">
              {[1, 2].map((num) => {
                const key = `file${num}`;
                const fData = files[key];
                return (
                  <div key={key} className="bg-slate-50 p-4 rounded-lg border border-slate-200">
                    <h3 className="font-semibold text-slate-700 mb-3 flex items-center gap-2">
                      <div className="bg-blue-100 text-blue-700 w-6 h-6 rounded-full flex items-center justify-center text-xs">{num}</div>
                      Attachment {num}
                    </h3>
                    <div className="relative group mb-4">
                      <input type="file" accept=".xls,.xlsx" onChange={(e) => handleFileUpload(e, key)} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"/>
                      <div className={`border-2 border-dashed rounded-lg p-6 text-center transition-colors ${fData.file ? 'border-green-400 bg-green-50' : 'border-slate-300 hover:border-blue-400 bg-white'}`}>
                        {fData.file ? (
                          <div className="flex flex-col items-center text-green-700">
                            <CheckCircle size={24} className="mb-2" />
                            <span className="text-xs font-medium truncate w-full px-2" title={fData.name}>{fData.name}</span>
                          </div>
                        ) : (
                          <div className="flex flex-col items-center text-slate-400 group-hover:text-blue-500">
                            <Upload size={24} className="mb-2" />
                            <span className="text-xs font-medium">Click or Drag Excel File</span>
                            <span className="text-[10px] mt-1">Supports .xls and .xlsx</span>
                          </div>
                        )}
                      </div>
                    </div>
                    <label className="block text-xs font-semibold text-slate-500 uppercase tracking-wider mb-1">Identify Source</label>
                    <select value={fData.label || ""} onChange={(e) => handleLabelChange(e.target.value, key)} className="w-full text-sm p-2 rounded border border-slate-300 focus:ring-2 focus:ring-blue-500 outline-none bg-white cursor-pointer">
                      <option value="" disabled>-- Select Party --</option>
                      {OPTIONS.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                    </select>
                  </div>
                );
              })}
            </div>

            {/* Action Area */}
            <div className="border-t border-slate-100 pt-6 flex flex-col items-center gap-4">
              {!processedFileUrl ? (
                <button onClick={handleMerge} disabled={isProcessing || !pyodide} className={`flex items-center gap-2 px-8 py-3 rounded-full font-bold shadow-lg transition-all ${(isProcessing || !pyodide) ? 'bg-slate-300 text-slate-500 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-700 text-white hover:-translate-y-1'}`}>
                  {isProcessing ? <><RefreshCw className="animate-spin" size={20} /> Processing...</> : "Merge & Process Files"}
                </button>
              ) : (
                <div className="flex flex-col sm:flex-row gap-3">
                  <a href={processedFileUrl} download="Final_Merged_Report.xlsx" className="flex items-center justify-center gap-2 px-6 py-2.5 rounded-full font-bold shadow-md bg-green-600 hover:bg-green-700 text-white text-sm transition-all hover:-translate-y-0.5">
                    <FileSpreadsheet size={18} /> Download Excel
                  </a>
                  <a href={processedPdfUrl} download="Final_Merged_Report.pdf" className="flex items-center justify-center gap-2 px-6 py-2.5 rounded-full font-bold shadow-md bg-red-500 hover:bg-red-600 text-white text-sm transition-all hover:-translate-y-0.5">
                    <FileText size={18} /> Download PDF
                  </a>
                  <button onClick={resetApp} className="flex items-center justify-center gap-2 px-5 py-2.5 rounded-full font-bold border-2 border-slate-200 text-slate-600 text-sm hover:bg-slate-50 transition-all">
                    <RotateCcw size={16} /> Start Over
                  </button>
                </div>
              )}
            </div>

            {/* Logs */}
            <div className="mt-2">
              <button onClick={() => setShowLogs(!showLogs)} className="flex items-center gap-2 text-xs font-semibold text-slate-400 hover:text-slate-600 transition-colors mx-auto">
                <Terminal size={14} /> {showLogs ? "Hide System Logs" : "Show System Logs"} {showLogs ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
              </button>
              {showLogs && (
                <div className="bg-slate-900 rounded-lg p-4 font-mono text-xs text-green-400 h-48 overflow-y-auto mt-3 shadow-inner">
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