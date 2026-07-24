/* ================================================================
   excelWorker.js — Web Worker de parseo de Excel (SheetJS)
   ────────────────────────────────────────────────────────────
   Se ejecuta en un hilo separado del hilo principal. Su ÚNICA
   responsabilidad es la parte cara en CPU: XLSX.read() del archivo
   binario y serializar cada hoja a un objeto plano de celdas.

   NO conoce nada de ExcelParser, DataStore, AccessManager ni del
   resto de la lógica de negocio — esa lógica sigue viviendo en
   app.js y se ejecuta igual que siempre, en el hilo principal,
   justo después de recibir el resultado de este worker. Esto
   mantiene el ciclo de vida y las reglas de negocio 100% intactos.

   Formato de mensajes:
     Entrada:  { id, arrayBuffer }               (Transferable)
     Salida:   { id, ok:true,  sheetNames, sheets }
           ó:  { id, ok:false, error }
================================================================ */

/* Carga SheetJS dentro del worker (mismo CDN/versión que index.html) */
importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');

self.onmessage = function (e) {
  const { id, arrayBuffer } = e.data || {};

  try {
    const data = new Uint8Array(arrayBuffer);

    /* Mismas opciones que se usaban en el hilo principal */
    const workbook = XLSX.read(data, { type: 'array', cellDates: false });

    /* Serializa cada hoja a un objeto plano { "A1": {v, w, t}, "!ref": "..." }
       — exactamente la misma forma que usa internamente SheetJS, así que
       XLSX.utils.sheet_to_json() y las funciones cellText()/parseMainSheet()
       etc. de ExcelParser (en app.js) funcionan SIN NINGÚN CAMBIO sobre
       este objeto reconstruido en el hilo principal. */
    const sheets = {};
    workbook.SheetNames.forEach((name) => {
      const ws = workbook.Sheets[name];
      const plainWs = {};
      Object.keys(ws).forEach((addr) => {
        if (addr.startsWith('!')) {
          plainWs[addr] = ws[addr]; // metadatos de hoja (!ref, !merges, etc.)
          return;
        }
        const cell = ws[addr];
        plainWs[addr] = { v: cell.v, w: cell.w, t: cell.t };
      });
      sheets[name] = plainWs;
    });

    self.postMessage({ id, ok: true, sheetNames: workbook.SheetNames, sheets });
  } catch (err) {
    self.postMessage({ id, ok: false, error: (err && err.message) || String(err) });
  }
};
