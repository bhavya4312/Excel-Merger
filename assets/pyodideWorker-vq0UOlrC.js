(function(d){"use strict";let e=null;const f=o=>{self.postMessage({type:"LOG",message:o})};self.onmessage=async o=>{const{id:n,type:l,payload:c}=o.data;try{if(l==="INIT"){self.postMessage({type:"INIT_STEP",step:0}),e||(e=await d.loadPyodide()),self.postMessage({type:"INIT_STEP",step:1}),await e.loadPackage("micropip");const t=e.pyimport("micropip");self.postMessage({type:"INIT_STEP",step:2}),await e.loadPackage("pandas"),await t.install("xlrd"),await t.install("openpyxl"),self.postMessage({type:"INIT_STEP",step:3}),self.postMessage({id:n,type:"INIT_SUCCESS"}),f("Python Web Worker initialized successfully.")}else if(l==="AUTO_DETECT"){if(!e)throw new Error("Python environment is not ready.");const{buffer:t,ext:i}=c,r=`temp_detect_${n}_${Date.now()}.${i}`;e.FS.writeFile(r,new Uint8Array(t));const p=`
import pandas as pd
import json

def detect():
    try:
        # Read header rows
        df = pd.read_excel('${r}', header=None, nrows=10)
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
`,s=await e.runPythonAsync(p),a=JSON.parse(s);try{e.FS.unlink(r)}catch{}self.postMessage({id:n,type:"AUTO_DETECT_SUCCESS",label:a.label})}else if(l==="MERGE"){if(!e)throw new Error("Python environment is not ready.");const{files:t,script:i}=c;for(const a of t)e.FS.writeFile(a.name,new Uint8Array(a.buffer));e.setStdout({batched:a=>f(`[PY] ${a}`)});const r=await e.runPythonAsync(i),p=JSON.parse(r);let s=null;e.FS.analyzePath("Final_Merged_Report.xlsx").exists&&(s=e.FS.readFile("Final_Merged_Report.xlsx").buffer),self.postMessage({id:n,type:"MERGE_SUCCESS",payload:{excelBuffer:s,jsonResult:p}},s?[s]:[])}}catch(t){self.postMessage({id:n,type:"ERROR",error:t.message||String(t)})}}})(pyodide_mjs);
