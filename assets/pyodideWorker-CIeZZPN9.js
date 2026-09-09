(function(){"use strict";importScripts("https://cdn.jsdelivr.net/pyodide/v0.23.4/full/pyodide.js");let e=null;const d=o=>{self.postMessage({type:"LOG",message:o})};self.onmessage=async o=>{const{id:r,type:l,payload:c}=o.data;try{if(l==="INIT"){self.postMessage({type:"INIT_STEP",step:0}),e||(e=await loadPyodide({indexURL:"https://cdn.jsdelivr.net/pyodide/v0.23.4/full/"})),self.postMessage({type:"INIT_STEP",step:1}),await e.loadPackage("micropip");const t=e.pyimport("micropip");self.postMessage({type:"INIT_STEP",step:2}),await e.loadPackage("pandas"),await t.install("xlrd"),await t.install("openpyxl"),self.postMessage({type:"INIT_STEP",step:3}),self.postMessage({id:r,type:"INIT_SUCCESS"}),d("Python Web Worker initialized successfully.")}else if(l==="AUTO_DETECT"){if(!e)throw new Error("Python environment is not ready.");const{buffer:t,ext:i}=c,a=`temp_detect_${r}_${Date.now()}.${i}`;e.FS.writeFile(a,new Uint8Array(t));const p=`
import pandas as pd
import json

def detect():
    try:
        # Read header rows
        df = pd.read_excel('${a}', header=None, nrows=10)
        full_text = ' '.join(df.astype(str).values.flatten()).upper()
        
        # Check matching keywords
        if 'RAHUL MEDICAL' in full_text or 'MEDICAL & SURGICAL' in full_text:
            return 'RAHUL MEDICAL & SURGICAL'
        elif 'RAHUL SURGICAL' in full_text or 'SURGICAL' in full_text:
            return 'RAHUL SURGICALS'
        return None
    except Exception as e:
        return None

json.dumps({'label': detect()})
`,s=await e.runPythonAsync(p),n=JSON.parse(s);try{e.FS.unlink(a)}catch{}self.postMessage({id:r,type:"AUTO_DETECT_SUCCESS",label:n.label})}else if(l==="MERGE"){if(!e)throw new Error("Python environment is not ready.");const{files:t,script:i}=c;for(const n of t)e.FS.writeFile(n.name,new Uint8Array(n.buffer));e.setStdout({batched:n=>d(`[PY] ${n}`)});const a=await e.runPythonAsync(i),p=JSON.parse(a);let s=null;e.FS.analyzePath("Final_Merged_Report.xlsx").exists&&(s=e.FS.readFile("Final_Merged_Report.xlsx").buffer),self.postMessage({id:r,type:"MERGE_SUCCESS",payload:{excelBuffer:s,jsonResult:p}},s?[s]:[])}}catch(t){self.postMessage({id:r,type:"ERROR",error:t.message||String(t)})}}})();
