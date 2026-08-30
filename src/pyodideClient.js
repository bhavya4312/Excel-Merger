// Client interface for interacting with the Pyodide Web Worker

let worker = null;
let messageIdCounter = 0;
const pendingRequests = new Map();
let logCallback = null;
let stepCallback = null;

export const getWorker = () => {
  if (!worker) {
    worker = new Worker(new URL('./pyodideWorker.js', import.meta.url));

    worker.onmessage = (e) => {
      const { id, type, step, label, payload, error, message } = e.data;

      if (type === 'LOG' && logCallback) {
        logCallback(message);
        return;
      }

      if (type === 'INIT_STEP' && stepCallback) {
        stepCallback(step);
        return;
      }

      if (id && pendingRequests.has(id)) {
        const { resolve, reject } = pendingRequests.get(id);
        pendingRequests.delete(id);

        if (type === 'ERROR') {
          reject(new Error(error));
        } else if (type === 'INIT_SUCCESS') {
          resolve(true);
        } else if (type === 'AUTO_DETECT_SUCCESS') {
          resolve(label);
        } else if (type === 'MERGE_SUCCESS') {
          resolve(payload);
        }
      }
    };

    worker.onerror = (err) => {
      console.error('Pyodide Worker Error:', err);
      if (logCallback) logCallback(`[Worker Error] ${err.message || 'Unknown error'}`);
    };
  }
  return worker;
};

export const initPyodideWorker = (onStep, onLog) => {
  stepCallback = onStep;
  logCallback = onLog;

  return new Promise((resolve, reject) => {
    const w = getWorker();
    const id = ++messageIdCounter;
    pendingRequests.set(id, { resolve, reject });
    w.postMessage({ id, type: 'INIT' });
  });
};

export const autoDetectLabelWorker = async (file) => {
  const buffer = await file.arrayBuffer();
  const ext = file.name.split('.').pop().toLowerCase();

  return new Promise((resolve, reject) => {
    const w = getWorker();
    const id = ++messageIdCounter;
    pendingRequests.set(id, { resolve, reject });
    w.postMessage({
      id,
      type: 'AUTO_DETECT',
      payload: { buffer, ext }
    }, [buffer]);
  });
};

export const runMergeWorker = (files, script) => {
  return new Promise((resolve, reject) => {
    const w = getWorker();
    const id = ++messageIdCounter;
    pendingRequests.set(id, { resolve, reject });

    const transferables = files.map(f => f.buffer).filter(Boolean);
    w.postMessage({
      id,
      type: 'MERGE',
      payload: { files, script }
    }, transferables);
  });
};

