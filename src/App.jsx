import React, { useState, useEffect, useRef } from 'react';
import { Upload, FileSpreadsheet, Download, RefreshCw, AlertCircle, CheckCircle } from 'lucide-react';

/**
 * INSTRUCTIONS FOR DEPLOYMENT:
 * 1. Ensure you add this script tag to your public/index.html <head>:
 * <script src="https://cdn.jsdelivr.net/pyodide/v0.23.4/full/pyodide.js"></script>
 * * 2. npm install lucide-react
 */

const App = () => {
  // Application State
  const [pyodide, setPyodide] = useState(null);
  const [loading, setLoading] = useState(true);
  const [loadingText, setLoadingText] = useState("Initializing Python Environment...");
  const [logs, setLogs] = useState([]);
  
  // File State
  const [files, setFiles] = useState({
    file1: { file: null, label: null, name: null },
    file2: { file: null, label: null, name: null }
  });
  
  const [processedFileUrl, setProcessedFileUrl] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState(null);

  // Constants
  const OPTIONS = ["RAHUL SURGICAL", "RAHUL MEDICAL & SURGICAL"];
  const TARGET_FILENAMES = {
    "RAHUL SURGICAL": "RAHUL SURGICAL.xls",
    "RAHUL MEDICAL & SURGICAL": "RAHUL MEDICAL & SURGICAL.xls"
  };

  // ----------------------------------------------------------------------
  // 1. Initialize Pyodide (Python in Browser)
  // ----------------------------------------------------------------------
  useEffect(() => {
    const initPyodide = async () => {
      try {
        if (!window.loadPyodide) {
          throw new Error("Pyodide script not found. Please add the CDN to index.html");
        }

        const py = await window.loadPyodide();
        setLoadingText("Installing Python libraries (pandas, openpyxl, xlrd)...");
        
        // Load required packages
        await py.loadPackage("pandas");
        await py.loadPackage("openpyxl");
        await py.loadPackage("xlrd");

        setPyodide(py);
        setLoading(false);
        addLog("Python environment ready.");
      } catch (err) {
        setError(`Failed to load Python: ${err.message}`);
        setLoading(false);
      }
    };

    initPyodide();
  }, []);

  const addLog = (msg) => {
    setLogs(prev => [...prev, `[${new Date().toLocaleTimeString()}] ${msg}`]);
  };

  // ----------------------------------------------------------------------
  // 2. File Handling
  // ----------------------------------------------------------------------
  const handleFileUpload = (e, fileKey) => {
    const uploadedFile = e.target.files[0];
    if (!uploadedFile) return;

    setFiles(prev => ({
      ...prev,
      [fileKey]: { ...prev[fileKey], file: uploadedFile, name: uploadedFile.name }
    }));
    
    // Reset output if new file uploaded
    setProcessedFileUrl(null); 
    setError(null);
  };

  const handleLabelChange = (val, fileKey) => {
    setFiles(prev => ({
      ...prev,
      [fileKey]: { ...prev[fileKey], label: val }
    }));
  };

  // ----------------------------------------------------------------------
  // 3. The Python Script (Embedded as a String)
  // ----------------------------------------------------------------------
  // We modify the input paths to match where we will write them in the virtual file system
  const pythonScript = `
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# ================= CONFIGURATION =================
# MODIFIED: Paths map to the root of the Pyodide Virtual File System
FILE_1 = "RAHUL MEDICAL & SURGICAL.xls" 
FILE_2 = "RAHUL SURGICAL.xls" 
OUTPUT_FILE = "Final_Merged_Report.xlsx"
# =================================================

def process_file(filepath, master_data, file_label):
    if not os.path.exists(filepath):
        print(f"Error: File not found - {filepath}")
        return

    print(f"Processing: {filepath}...")

    try:
        # Determine engine based on extension or content
        if filepath.lower().endswith('.xls'):
            df = pd.read_excel(filepath, header=None, engine='xlrd')
        else:
            df = pd.read_excel(filepath, header=None, engine='openpyxl')
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return

    current_party_key = None

    for index, row in df.iterrows():
        row_data = ["" if pd.isna(x) else x for x in row.tolist()]
        if len(row_data) < 2: continue

        col_b_value = str(row_data[1]).strip()

        if col_b_value.startswith('-'):
            party_name = col_b_value.lstrip('-').strip()
            city = str(row_data[3]).strip() if len(row_data) > 3 else ""
            unique_key = (party_name + city).lower().replace(" ", "")

            try:
                amount = float(row_data[4]) if len(row_data) > 4 else 0.0
            except ValueError:
                amount = 0.0

            if unique_key not in master_data:
                master_data[unique_key] = {
                    'original_name': party_name,
                    'city': city,
                    'transactions_file1': [],
                    'transactions_file2': [],
                    'total_combined': 0.0,
                    'totals_breakdown': {'file1': 0.0, 'file2': 0.0}
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
                                debit = float(parts[0])
                                credit = float(parts[1])
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

                list_key = f'transactions_{file_label}'
                master_data[current_party_key][list_key].append(row_data)

def style_excel_file(filename):
    print("Applying styles...")
    wb = load_workbook(filename)
    ws = wb.active
    
    # Styles
    HEADER_SIZE = 14
    REGULAR_SIZE = 12
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
        cell_a = row[0]
        cell_b = row[1]
        val_a = str(cell_a.value).strip() if cell_a.value else ""
        val_b = str(cell_b.value).strip() if cell_b.value else ""

        is_party_header = (val_a and not val_b and not val_a.lower().startswith("rcpt") 
                           and not val_a.startswith("Outstanding Summary") and not val_a.startswith("---"))

        if is_party_header:
            if len(row) >= 7:
                for i in range(7):
                    cell = row[i]
                    cell.fill = party_fill
                    b_right = thin if i == 5 or i == 6 else no_side
                    b_left = thin if i == 6 else no_side
                    cell.border = Border(top=thin, left=b_left, right=b_right, bottom=no_side)
                cell_a.font = party_header_font
                cell_a.alignment = Alignment(horizontal='center', vertical='center')
                row[6].font = party_header_font
                row[6].alignment = Alignment(horizontal='right', vertical='center')
                ws.merge_cells(start_row=cell_a.row, start_column=1, end_row=cell_a.row, end_column=6)

        elif val_a.startswith("---"):
            if len(row) >= 7:
                ws.merge_cells(start_row=cell_a.row, start_column=1, end_row=cell_a.row, end_column=7)
                cell_a.font = section_font
                cell_a.alignment = Alignment(horizontal='center', vertical='center')
                fill_color = medical_fill if "Medical" in val_a else surgical_fill
                for cell in row:
                    cell.fill = fill_color
                    cell.border = thin_border

        elif val_b == "Bill":
            if len(row) >= 7:
                for cell in row:
                    cell.font = sub_header_font
                    cell.fill = sub_header_fill
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal='center')

        elif val_a.startswith("Outstanding Summary"):
            if len(row) >= 7:
                for i in range(7):
                    cell = row[i]
                    cell.fill = summary_fill
                    cell.border = Border(top=thin, left=no_side, right=no_side, bottom=no_side)
                cell_a.font = summary_font
                cell_a.alignment = Alignment(horizontal='center', vertical='center')
                ws.merge_cells(start_row=cell_a.row, start_column=1, end_row=cell_a.row, end_column=7)

        else:
            if any(cell.value for cell in row):
                if len(row) >= 7:
                    for cell in row:
                        cell.font = regular_font
                        cell.border = thin_border

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
                        if val.startswith("Outstanding Summary") or val.startswith("---") or val_b_check == "Bill": continue
                        cell_len = len(val)
                        if cell_len > max_length: max_length = cell_len
                except: pass
            ws.column_dimensions[col_letter].width = max(max_length + 2, 8) if i == 1 else min((max_length + 2), 60)

    wb.save(filename)

# --- MAIN EXECUTION BLOCK ---
combined_data = {}

# IMPORTANT: These files must exist in the VFS before this runs
if os.path.exists(FILE_1):
    process_file(FILE_1, combined_data, 'file1')
else:
    print(f"Skipping File 1: Not found at {FILE_1}")

if os.path.exists(FILE_2):
    process_file(FILE_2, combined_data, 'file2')
else:
    print(f"Skipping File 2: Not found at {FILE_2}")

print("Merging data...")
final_output_rows = []
transaction_header = ["Type", "Bill", "Date", "Debit", "Credit", "Balance", "Days"]

for key, data in combined_data.items():
    display_name = f"{data['original_name']}  ({data['city']})"
    header_row = [display_name, "", "", "", "", "", data['total_combined']]
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
    summary_text = f"Outstanding Summary   |   Medical: {t1}   |   Surgical: {t2}   |   Total: {total}"
    final_output_rows.append([summary_text, "", "", "", "", "", ""])
    final_output_rows.append([""] * 7)

print(f"Saving to {OUTPUT_FILE}...")
df = pd.DataFrame(final_output_rows)
df.to_excel(OUTPUT_FILE, index=False, header=False)
style_excel_file(OUTPUT_FILE)
print("Processing Complete.")
`;

  // ----------------------------------------------------------------------
  // 4. Logic: Write Files -> Run Python -> Read Result
  // ----------------------------------------------------------------------
  const handleMerge = async () => {
    setIsProcessing(true);
    setLogs([]);
    setError(null);
    setProcessedFileUrl(null);

    // Validation
    const { file1, file2 } = files;
    if (!file1.file || !file2.file) {
      setError("Please upload both files.");
      setIsProcessing(false);
      return;
    }
    if (!file1.label || !file2.label) {
      setError("Please select options for both files.");
      setIsProcessing(false);
      return;
    }
    if (file1.label === file2.label) {
      setError("Please ensure files have different labels (One Medical, One Surgical).");
      setIsProcessing(false);
      return;
    }

    try {
      addLog("Reading uploaded files...");
      
      // Helper to read file and write to Pyodide VFS
      const writeFileToPyodide = async (fileObj, label) => {
        const arrayBuffer = await fileObj.file.arrayBuffer();
        const uint8Array = new Uint8Array(arrayBuffer);
        
        // Use the label to determine the filename the Python script expects
        const targetName = TARGET_FILENAMES[label];
        if(!targetName) throw new Error(`Unknown label mapping: ${label}`);

        // Write to root of VFS
        pyodide.FS.writeFile(targetName, uint8Array);
        addLog(`Saved ${fileObj.file.name} as "${targetName}" in virtual memory.`);
      };

      await writeFileToPyodide(file1, file1.label);
      await writeFileToPyodide(file2, file2.label);

      addLog("Executing Python logic...");
      
      // Capture Python print statements
      pyodide.setStdout({ batched: (msg) => addLog(`[PY] ${msg}`) });

      // Run the script
      await pyodide.runPythonAsync(pythonScript);

      // Check if output exists
      if (pyodide.FS.analyzePath("Final_Merged_Report.xlsx").exists) {
        const fileContent = pyodide.FS.readFile("Final_Merged_Report.xlsx");
        const blob = new Blob([fileContent], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const url = URL.createObjectURL(blob);
        setProcessedFileUrl(url);
        addLog("Merge successful! File ready for download.");
      } else {
        throw new Error("Output file was not generated by Python script.");
      }

    } catch (err) {
      console.error(err);
      setError(err.message);
      addLog(`Error: ${err.message}`);
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
            Securely merge your Surgical and Medical reports entirely in your browser.
          </p>
        </div>

        {/* Loading State for Python */}
        {loading ? (
          <div className="p-12 flex flex-col items-center justify-center text-center">
            <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mb-4"></div>
            <p className="text-slate-500 font-medium">{loadingText}</p>
            <p className="text-xs text-slate-400 mt-2">This happens only once per session.</p>
          </div>
        ) : (
          <div className="p-6 space-y-8">
            
            {/* Error Banner */}
            {error && (
              <div className="bg-red-50 text-red-700 p-4 rounded-lg flex items-start gap-3 border border-red-200">
                <AlertCircle className="shrink-0 mt-0.5" size={20} />
                <div className="text-sm font-medium">{error}</div>
              </div>
            )}

            {/* File Inputs Grid */}
            <div className="grid md:grid-cols-2 gap-6">
              {[1, 2].map((num) => {
                const key = `file${num}`;
                const fData = files[key];
                
                return (
                  <div key={key} className="bg-slate-50 p-4 rounded-lg border border-slate-200">
                    <h3 className="font-semibold text-slate-700 mb-3 flex items-center gap-2">
                      <div className="bg-blue-100 text-blue-700 w-6 h-6 rounded-full flex items-center justify-center text-xs">
                        {num}
                      </div>
                      Attachment {num}
                    </h3>

                    {/* File Drop Area */}
                    <div className="relative group mb-4">
                      <input 
                        type="file" 
                        accept=".xls,.xlsx"
                        onChange={(e) => handleFileUpload(e, key)}
                        className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                      />
                      <div className={`
                        border-2 border-dashed rounded-lg p-6 text-center transition-colors
                        ${fData.file ? 'border-green-400 bg-green-50' : 'border-slate-300 hover:border-blue-400 hover:bg-white'}
                      `}>
                        {fData.file ? (
                          <div className="flex flex-col items-center text-green-700">
                            <CheckCircle size={24} className="mb-2" />
                            <span className="text-xs font-medium truncate w-full px-2">{fData.name}</span>
                          </div>
                        ) : (
                          <div className="flex flex-col items-center text-slate-400">
                            <Upload size={24} className="mb-2" />
                            <span className="text-xs">Click to upload Excel</span>
                          </div>
                        )}
                      </div>
                    </div>

                    {/* Dropdown */}
                    <label className="block text-xs font-semibold text-slate-500 uppercase tracking-wider mb-1">
                      Identify Source
                    </label>
                    <select 
                      className="w-full text-sm p-2 rounded border border-slate-300 focus:ring-2 focus:ring-blue-500 outline-none bg-white"
                      value={fData.label || ""}
                      onChange={(e) => handleLabelChange(e.target.value, key)}
                    >
                      <option value="" disabled>-- Select Party --</option>
                      {OPTIONS.map(opt => (
                        <option key={opt} value={opt}>{opt}</option>
                      ))}
                    </select>
                  </div>
                );
              })}
            </div>

            {/* Action Area */}
            <div className="border-t border-slate-100 pt-6 flex flex-col items-center">
              {!processedFileUrl ? (
                <button
                  onClick={handleMerge}
                  disabled={isProcessing}
                  className={`
                    flex items-center gap-2 px-8 py-3 rounded-full font-bold shadow-lg transition-all
                    ${isProcessing 
                      ? 'bg-slate-300 text-slate-500 cursor-not-allowed' 
                      : 'bg-blue-600 hover:bg-blue-700 text-white hover:shadow-xl hover:-translate-y-1'
                    }
                  `}
                >
                  {isProcessing ? (
                    <><RefreshCw className="animate-spin" size={20} /> Processing...</>
                  ) : (
                    "Merge & Process Files"
                  )}
                </button>
              ) : (
                <a
                  href={processedFileUrl}
                  download="Final_Merged_Report.xlsx"
                  className="flex items-center gap-2 px-8 py-3 rounded-full font-bold shadow-lg bg-green-600 hover:bg-green-700 text-white transition-all animate-bounce"
                >
                  <Download size={20} /> Download Final Report
                </a>
              )}
            </div>

            {/* Logs Console */}
            <div className="bg-slate-900 rounded-lg p-4 font-mono text-xs text-green-400 h-48 overflow-y-auto">
              <div className="text-slate-500 border-b border-slate-800 pb-2 mb-2">System Logs</div>
              {logs.length === 0 && <span className="text-slate-600 italic">Waiting for input...</span>}
              {logs.map((log, i) => (
                <div key={i} className="mb-1">{log}</div>
              ))}
            </div>

          </div>
        )}
      </div>
      
      <div className="text-center mt-6 text-slate-400 text-xs">
        Powered by Pyodide • Runs locally in browser
      </div>
    </div>
  );
};

export default App;