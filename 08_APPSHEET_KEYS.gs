/*******************************************************
  * 08_APPSHEET_KEYS.gs — Keys AppSheet (SINONIMS id)
  *******************************************************/

function EMI_appsSheetFixKeys() {
   withLock_('APPSHEET_FIX_KEYS', () => {
      const ss = getMasterSs_();

      // 1) SINONIMS: garantir id_sinonim
      ensureIdColumnAndFill_(ss, SHEETS.SINONIMS, 'id_sinonim', 'SYN');

      // (Opcional futur) si més endavant vols també ID a CATALEGS/FORMACIONS:
      // ensureIdColumnAndFill_(ss, SHEETS.CATALEGS, 'id_catalog', 'CAT');
      // ensureIdColumnAndFill_(ss, SHEETS.FORMACIONS_CATALOG, 'id_formacio', 'FOR');

      SpreadsheetApp.flush();
      log_('INFO','APPSHEET_KEYS_OK','Keys AppSheet arreglades', {});
   });
}

function ensureIdColumnAndFill_(ss, sheetName, idHeader, prefix) {
   const sh = ss.getSheetByName(sheetName);
   if (!sh) throw new Error('No existeix la sheet: ' + sheetName);

   // headers
   const lastCol = sh.getLastColumn();
   let headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);

   let col = headers.indexOf(idHeader) + 1;
   if (!col) {
      // afegeix columna al final
      sh.insertColumnAfter(lastCol);
      sh.getRange(1, lastCol + 1).setValue(idHeader);
      col = lastCol + 1;
      headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      log_('INFO','ADD_ID_COLUMN','Afegida columna ID', { sheet: sheetName, idHeader, col });
   }

   const lastRow = sh.getLastRow();
   if (lastRow <= 1) return;

   // llegeix ids existents
   const rng = sh.getRange(2, col, lastRow - 1, 1);
   const vals = rng.getValues();

   let filled = 0;
   for (let i = 0; i < vals.length; i++) {
      const v = String(vals[i][0] || '').trim();
      if (v) continue;
      // id determinista per fila (únic): prefix + uuid curt
      const id = prefix + '-' + Utilities.getUuid();
      vals[i][0] = id;
      filled++;
   }

   if (filled) rng.setValues(vals);
   log_('INFO','FILL_IDS_DONE','IDs omplerts', { sheet: sheetName, idHeader, filled, total: vals.length });
}