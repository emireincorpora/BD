/*******************************************************
  * 19_DIAGNOSE_AND_FIX.gs
  * - Diagnòstic clar al LOG (què hi ha i què falta)
  * - Fix robust: assegura VACANTS.text_index (sense perdre’s en columnes fantasma)
  * - Pipeline “com toca”: reindex -> match 2 sentits
  * - Test connexió ESCO amb log (per saber si pot nodrir-se de web)
  *******************************************************/

function EMI_DIAGNOSE_masterNow() {
   const ss = (typeof getMasterSs_ === 'function') ? getMasterSs_() : SpreadsheetApp.getActiveSpreadsheet();

   const info = {};

   // Comptes per CATALEGS
   info.cat = EMI_DIAGNOSE_countByTipus_(ss, 'CATALEGS', ['tipus','actiu']);

   // Comptes per SINONIMS
   info.syn = EMI_DIAGNOSE_countByTipus_(ss, 'SINONIMS', ['tipus','actiu']);

   // ESCO present?
   info.hasEsco = Boolean(
      (info.cat.counts && (info.cat.counts.ESCO_OCCUPATION || info.cat.counts.ESCO_SKILL))
   );

   // VACANTS té text_index?
   info.vacants_has_text_index = EMI_DIAGNOSE_hasColumn_(ss, 'VACANTS', 'text_index');

   // CV_RAW camps “raspables”
   info.cvraw_nonempty = EMI_DIAGNOSE_nonEmptyCounts_(ss, 'CV_RAW', [
      'cv_sector_objectiu','cv_competencies','cv_experiencia','cv_formacio','cv_text_index','raw_json'
   ]);

   // PERSONES tags_auto?
   info.persones_nonempty = EMI_DIAGNOSE_nonEmptyCounts_(ss, 'PERSONES', [
      'tags_auto','text_index','cv_sector_objectiu','cv_competencies','cv_experiencia','cv_formacio'
   ]);

   EMI_LOG_('INFO','DIAGNOSE_MASTER','Diagnòstic MASTER', info);
   uiAlert_(
      'Diagnòstic fet ✅\n\n' +
      '- Mira el full LOG per veure el detall.\n' +
      '- Punts típics: VACANTS.text_index, ESCO present, CV_RAW buit.\n'
   );
}

function EMI_TEST_escoConnectionNow() {
   // Test molt simple: si això falla, el “harvest web” no podrà funcionar.
   const url = 'https://ec.europa.eu/esco/api/search?text=soldador&language=es&type=occupation&limit=1&offset=0&full=false&selectedVersion=latest&viewObsolete=false';
   try {
      const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, headers: { 'Accept': 'application/json' }});
      const code = res.getResponseCode();
      const body = res.getContentText();
      EMI_LOG_('INFO','ESCO_TEST', 'Resposta ESCO', { code, sample: body.slice(0, 200) });
      uiAlert_('Test ESCO fet ✅\n\nCodi resposta: ' + code + '\nMira el LOG per veure el detall.');
   } catch (e) {
      EMI_LOG_('ERROR','ESCO_TEST_FAIL','No puc connectar a ESCO', { error: String(e) });
      uiAlert_('Test ESCO FALLA ❌\n\nMira el LOG (ERROR) i m’ho enganxes aquí.');
   }
}

function EMI_FIX_ensureVacantsTextIndex_() {
   const ss = getMasterSs_();
   const sh = ss.getSheetByName('VACANTS');
   if (!sh) return;

   const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x => String(x||'').trim());
   if (headers.indexOf('text_index') !== -1) return;

   // Inserim després de l’última capçalera real (no al final “fantasma”)
   const lastHeaderCol = EMI_findLastHeaderCol_(headers);
   sh.insertColumnAfter(lastHeaderCol);
   sh.getRange(1, lastHeaderCol + 1).setValue('text_index');

   EMI_LOG_('INFO','FIX_VACANTS_TEXT_INDEX','Afegida columna text_index a VACANTS', { col: lastHeaderCol + 1 });
}

function EMI_findLastHeaderCol_(headers) {
   for (let i = headers.length; i >= 1; i--) {
      if (String(headers[i-1]||'').trim() !== '') return i;
   }
   return headers.length || 1;
}

/** ---------------- Helpers diagnòstic ---------------- */

function EMI_DIAGNOSE_countByTipus_(ss, sheetName, neededCols) {
   const sh = ss.getSheetByName(sheetName);
   if (!sh || sh.getLastRow() < 2) return { sheetName, counts: {}, rows: 0 };

   const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x => String(x||'').trim());
   const idx = {};
   headers.forEach((h,i)=>{ if(h) idx[h]=i; });

   const tipusI = idx[neededCols[0]];
   const actiuI = idx[neededCols[1]];

   const data = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
   const counts = {};

   data.forEach(r => {
      const tipus = (tipusI!=null) ? String(r[tipusI]||'').trim() : '';
      if (!tipus) return;

      const actiu = (actiuI!=null) ? r[actiuI] : 'SI';
      if (!isYes_(actiu)) return;

      counts[tipus] = (counts[tipus] || 0) + 1;
   });

   return { sheetName, counts, rows: data.length };
}

function EMI_DIAGNOSE_nonEmptyCounts_(ss, sheetName, cols) {
   const sh = ss.getSheetByName(sheetName);
   if (!sh || sh.getLastRow() < 2) return { sheetName, counts: {}, rows: 0 };

   const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x => String(x||'').trim());
   const idx = {};
   headers.forEach((h,i)=>{ if(h) idx[h]=i; });

   const data = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();

   const out = {};
   cols.forEach(c => out[c] = 0);

   data.forEach(r => {
      cols.forEach(c => {
         const i = idx[c];
         if (i == null) return;
         const v = r[i];
         if (v === null || v === undefined) return;
         const s = String(v).trim();
         if (s !== '' && s.toLowerCase() !== 'none') out[c]++;
      });
   });

   return { sheetName, counts: out, rows: data.length };
}

function EMI_DIAGNOSE_hasColumn_(ss, sheetName, colName) {
   const sh = ss.getSheetByName(sheetName);
   if (!sh) return false;
   const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x => String(x||'').trim());
   return headers.indexOf(colName) !== -1;
}

/** --------- LOG wrapper (si no existeix log_) ---------- */

function EMI_LOG_(nivell, accio, missatge, obj) {
   try {
      if (typeof log_ === 'function') {
         log_(nivell, accio, missatge, obj || {});
         return;
      }
   } catch(e) {}

   // fallback: escrivim al full LOG si existeix
   try {
      const ss = (typeof getMasterSs_ === 'function') ? getMasterSs_() : SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName('LOG');
      if (!sh) return;
      sh.appendRow([new Date(), nivell, accio, missatge, JSON.stringify(obj || {})]);
   } catch(e2) {}
}
