/*******************************************************
  * 17_TAGS_AUTO_LABELS.gs
  * - Omple tags_persona_auto i tags_vacant_auto (en català, llegible)
  * - A partir de tags_auto (intern: TIPUS:CODI) i el diccionari (CATALEGS)
  *
  * Executa: EMI_refreshAutoTagsLabelsNow()
  *******************************************************/

function EMI_refreshAutoTagsLabelsNow() {
   withLock_('AUTO_TAGS_LABELS', () => {
      const ss = getMasterSs_();
      const dict = buildBrainDictionaries_(); // { catalogs, reverse }

      const res1 = fillAutoTagLabelsForSheet_(ss, SHEETS.PERSONES, 'tags_persona_auto', dict.catalogs);
      const res2 = fillAutoTagLabelsForSheet_(ss, SHEETS.VACANTS,   'tags_vacant_auto',   dict.catalogs);

      SpreadsheetApp.flush();
      log_('INFO','AUTO_TAGS_LABELS_OK','Auto tags (labels) actualitzats', { persones: res1, vacants: res2 });
   });
}

function fillAutoTagLabelsForSheet_(ss, sheetName, outColName, catalogs) {
   const sh = ss.getSheetByName(sheetName);
   if (!sh || sh.getLastRow() < 2) return { updated: 0, rows: 0 };

   const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
   const idx = indexMap2_(headers);

   const colTagsAuto = (idx.tags_auto||0) - 1;
   const colOut         = (idx[outColName]||0) - 1;

   if (colTagsAuto < 0) return { updated: 0, rows: 0, warn: 'No hi ha tags_auto' };

   // si la columna no existeix, la creem al final
   if (colOut < 0) {
      sh.insertColumnAfter(sh.getLastColumn());
      const newCol = sh.getLastColumn();
      sh.getRange(1,newCol).setValue(outColName);
      idx[outColName] = newCol;
   }

   const lastRow = sh.getLastRow();
   const lastCol = sh.getLastColumn();
   const values = sh.getRange(2,1,lastRow-1,lastCol).getValues();

   let updated = 0;

   values.forEach(r => {
      const tagsAuto = String(r[colTagsAuto] || '').trim();
      if (!tagsAuto) return;

      const labels = tagsAutoToLabels_(tagsAuto, catalogs);
      const newVal = labels.join(' | '); // format “llista de paraules”
      const oldVal = String(r[(idx[outColName]-1)] || '').trim();

      if (newVal && newVal !== oldVal) {
         r[idx[outColName]-1] = newVal;
         updated++;
      }
   });

   if (updated > 0) {
      sh.getRange(2,1,values.length,lastCol).setValues(values);
   }

   return { updated, rows: values.length };
}

function tagsAutoToLabels_(tagsAuto, catalogs) {
   // tags_auto: "TIPUS:CODI ; TIPUS:CODI ; ..."
   const parts = tagsAuto.split(/[;,\n]+/).map(s => String(s).trim()).filter(Boolean);
   const out = [];
   const seen = new Set();

   parts.forEach(p => {
      const m = p.match(/^([^:]+):(.+)$/);
      if (!m) return;
      const tipus = m[1].trim();
      const codi   = m[2].trim();
      const key = tipus + '||' + codi;

      const entry = catalogs[key];
      const label = entry && entry.valor ? String(entry.valor).trim() : codi;

      if (!label) return;
      const norm = normalizeToken_(label);
      if (seen.has(norm)) return;
      seen.add(norm);
      out.push(label);
   });

   out.sort((a,b) => a.localeCompare(b, 'ca'));
   return out;
}