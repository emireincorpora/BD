/*******************************************************
  * 18_BRAIN_READERS.gs — Lectura de diccionaris (CATALEGS/SINONIMS/FORMACIONS) · V1
  * Soluciona: ReferenceError: readCatalogs_ is not defined
  *
  * Llegeix aquests fulls del MASTER:
  *   - CATALEGS: tipus | codi | valor | etiqueta | actiu | ordre
  *   - SINONIMS: tipus | codi_ref | sinonim | pes | actiu | notes | id_sinonim
  *   - FORMACIONS_CATALOG: codi_formacio | tipus | nom | familia | nivell | durada_hores | sinonims | actiu | ordre | notes | updated_at
  *******************************************************/

function readCatalogs_(ss) {
   ss = ss || (typeof getMasterSs_ === 'function' ? getMasterSs_() : SpreadsheetApp.getActiveSpreadsheet());
   const sheetName = (typeof SHEETS !== 'undefined' && SHEETS.CATALEGS) ? SHEETS.CATALEGS : 'CATALEGS';
   const sh = ss.getSheetByName(sheetName);
   if (!sh || sh.getLastRow() < 2) return {};

   const lastCol = Math.max(6, sh.getLastColumn());
   const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(x => String(x || '').trim());
   const idx = _rd_indexMap_(headers);

   const lr = sh.getLastRow();
   const data = sh.getRange(2, 1, lr - 1, lastCol).getValues();

   const out = {}; // key = tipus||codi
   for (let i = 0; i < data.length; i++) {
      const row = data[i];

      const tipus = _rd_cell_(row, idx, 'tipus');
      const codi   = _rd_cell_(row, idx, 'codi');
      const valor = _rd_cell_(row, idx, 'valor');
      const etiqueta = _rd_cell_(row, idx, 'etiqueta');
      const actiu = _rd_cellRaw_(row, idx, 'actiu');

      if (!tipus || !codi || !valor) continue;
      if (!_rd_isYes_(actiu)) continue;

      out[`${tipus}||${codi}`] = {
         tipus: String(tipus).trim(),
         codi: String(codi).trim(),
         valor: String(valor).trim(),
         etiqueta: String(etiqueta || '').trim(),
         syns: [] // s’afegeixen a buildBrainDictionaries_ via SINONIMS i tokenització
      };
   }
   return out;
}

function readSinonims_(ss) {
   ss = ss || (typeof getMasterSs_ === 'function' ? getMasterSs_() : SpreadsheetApp.getActiveSpreadsheet());
   const sheetName = (typeof SHEETS !== 'undefined' && SHEETS.SINONIMS) ? SHEETS.SINONIMS : 'SINONIMS';
   const sh = ss.getSheetByName(sheetName);
   if (!sh || sh.getLastRow() < 2) return [];

   const lastCol = Math.max(7, sh.getLastColumn());
   const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(x => String(x || '').trim());
   const idx = _rd_indexMap_(headers);

   const lr = sh.getLastRow();
   const data = sh.getRange(2, 1, lr - 1, lastCol).getValues();

   const out = [];
   for (let i = 0; i < data.length; i++) {
      const row = data[i];

      const tipus    = _rd_cell_(row, idx, 'tipus');
      const codiRef = _rd_cell_(row, idx, 'codi_ref');
      const sinonim = _rd_cell_(row, idx, 'sinonim');
      const pesRaw   = _rd_cellRaw_(row, idx, 'pes');
      const actiu    = _rd_cellRaw_(row, idx, 'actiu');
      const notes    = _rd_cell_(row, idx, 'notes');

      if (!tipus || !codiRef || !sinonim) continue;
      if (!_rd_isYes_(actiu)) continue;

      const pes = (pesRaw === '' || pesRaw === null || pesRaw === undefined) ? 1 : Number(pesRaw);
      out.push({
         tipus: String(tipus).trim(),
         codi_ref: String(codiRef).trim(),
         sinonim: String(sinonim).trim(),
         pes: isNaN(pes) ? 1 : pes,
         notes: String(notes || '')
      });
   }
   return out;
}

function readFormacions_(ss) {
   ss = ss || (typeof getMasterSs_ === 'function' ? getMasterSs_() : SpreadsheetApp.getActiveSpreadsheet());
   const sh = ss.getSheetByName('FORMACIONS_CATALOG');
   if (!sh || sh.getLastRow() < 2) return [];

   const lastCol = Math.max(11, sh.getLastColumn());
   const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(x => String(x || '').trim());
   const idx = _rd_indexMap_(headers);

   const lr = sh.getLastRow();
   const data = sh.getRange(2, 1, lr - 1, lastCol).getValues();

   const out = [];
   for (let i = 0; i < data.length; i++) {
      const row = data[i];

      const codi = _rd_cell_(row, idx, 'codi_formacio');
      const nom   = _rd_cell_(row, idx, 'nom');
      const actiu = _rd_cellRaw_(row, idx, 'actiu');

      if (!codi || !nom) continue;
      if (!_rd_isYes_(actiu)) continue;

      const familia = _rd_cell_(row, idx, 'familia');
      const nivell   = _rd_cell_(row, idx, 'nivell');
      const synRaw   = _rd_cell_(row, idx, 'sinonims');

      const syns = _rd_splitSyns_(synRaw);

      // Important: mantenim etiqueta “FORMACIO” perquè el motor generi tags tipus FORMACIO:XXXX
      out.push({
         tipus: 'FORMACIO',
         codi: String(codi).trim(),
         valor: String(nom).trim(),
         etiqueta: [familia, nivell].filter(Boolean).join(' | '),
         syns: syns
      });
   }
   return out;
}

/** -------------------------
  * Helpers interns
  * ------------------------- */

function _rd_indexMap_(headers) {
   const m = {};
   for (let i = 0; i < headers.length; i++) {
      const k = String(headers[i] || '').trim();
      if (k) m[k] = i + 1;
   }
   return m;
}

function _rd_cell_(row, idx, colName) {
   const c = idx[colName];
   if (!c) return '';
   const v = row[c - 1];
   if (v === null || v === undefined) return '';
   return String(v).trim();
}

function _rd_cellRaw_(row, idx, colName) {
   const c = idx[colName];
   if (!c) return '';
   return row[c - 1];
}

function _rd_isYes_(v) {
   if (v === true) return true;
   if (v === false) return false;
   const s = String(v === null || v === undefined ? '' : v).trim().toUpperCase();
   return (s === 'SI' || s === 'SÍ' || s === 'YES' || s === 'TRUE' || s === '1');
}

function _rd_splitSyns_(s) {
   const raw = String(s || '').trim();
   if (!raw) return [];
   const parts = raw
      .replace(/\r/g, '\n')
      .replace(/[•·|]/g, ';')
      .split(/[\n;,]+/)
      .map(x => String(x || '').trim())
      .filter(Boolean);

   // dedupe
   const seen = {};
   const out = [];
   for (let i = 0; i < parts.length; i++) {
      const key = _rd_norm_(parts[i]);
      if (!key) continue;
      if (seen[key]) continue;
      seen[key] = true;
      out.push(parts[i]);
   }
   return out.slice(0, 80);
}

function _rd_norm_(s) {
   s = String(s || '').toLowerCase().trim();
   try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) {}
   s = s.replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
   return s;
}