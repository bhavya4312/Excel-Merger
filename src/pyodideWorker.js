/* eslint-disable no-undef */
// Web Worker for isolated Pyodide WASM runtime
importScripts('https://cdn.jsdelivr.net/pyodide/v0.23.4/full/pyodide.js');

let pyodide = null;

const sendLog = (msg) => {
  self.postMessage({ type: 'LOG', message: msg });
};

self.onmessage = async (e) => {
  const { id, type, payload } = e.data;

  try {
    if (type === 'INIT') {
      self.postMessage({ type: 'INIT_STEP', step: 0 });

      if (!pyodide) {
        pyodide = await loadPyodide({
          indexURL: 'https://cdn.jsdelivr.net/pyodide/v0.23.4/full/'
        });
      }
      self.postMessage({ type: 'INIT_STEP', step: 1 });

      await pyodide.loadPackage('micropip');
      const micropip = pyodide.pyimport('micropip');
      self.postMessage({ type: 'INIT_STEP', step: 2 });

      await pyodide.loadPackage('pandas');
      await micropip.install('xlrd');
      await micropip.install('openpyxl');
      self.postMessage({ type: 'INIT_STEP', step: 3 });

      self.postMessage({ id, type: 'INIT_SUCCESS' });
      sendLog('Python Web Worker initialized successfully.');
    } else if (type === 'AUTO_DETECT') {
      if (!pyodide) throw new Error('Python environment is not ready.');

      const { buffer, ext } = payload;
      const tempFilename = `temp_detect_${id}_${Date.now()}.${ext}`;
      pyodide.FS.writeFile(tempFilename, new Uint8Array(buffer));

      const detectScript = `
import pandas as pd
import json

def detect():
    try:
        # Read header rows
        df = pd.read_excel('${tempFilename}', header=None, nrows=10)
        full_text = ' '.join(df.astype(str).values.flatten()).upper()
        
        # Check matching keywords
        if 'RAHUL MEDICAL' in full_text or 'MEDICAL & SURGICAL' in full_text:
            return 'RAHUL MEDICAL & SURGICAL'
        elif 'RAHUL SURGICAL' in full_text or 'SURGICAL' in full_text:
            return 'RAHUL SURGICAL'
        return None
    except Exception as e:
        return None

json.dumps({'label': detect()})
`;
      const resStr = await pyodide.runPythonAsync(detectScript);
      const res = JSON.parse(resStr);
      try {
        pyodide.FS.unlink(tempFilename);
      } catch {
        // ignore cleanup error
      }

      self.postMessage({ id, type: 'AUTO_DETECT_SUCCESS', label: res.label });
    } else if (type === 'MERGE') {
      if (!pyodide) throw new Error('Python environment is not ready.');

      const { files, script } = payload;

      // Write files to virtual FS
      for (const f of files) {
        pyodide.FS.writeFile(f.name, new Uint8Array(f.buffer));
      }

      // Configure stdout streaming
      pyodide.setStdout({
        batched: (msg) => sendLog(`[PY] ${msg}`)
      });

      // Run merge script
      const jsonResultStr = await pyodide.runPythonAsync(script);
      const jsonResult = JSON.parse(jsonResultStr);

      let excelBuffer = null;
      if (pyodide.FS.analyzePath('Final_Merged_Report.xlsx').exists) {
        excelBuffer = pyodide.FS.readFile('Final_Merged_Report.xlsx').buffer;
      }

      self.postMessage({
        id,
        type: 'MERGE_SUCCESS',
        payload: {
          excelBuffer,
          jsonResult
        }
      }, excelBuffer ? [excelBuffer] : []);
    }
  } catch (err) {
    self.postMessage({ id, type: 'ERROR', error: err.message || String(err) });
  }
};
