/*******************************************************
  * 07_DIAG.gs — smoke test
  *******************************************************/

function EMI_smokeTest_V2() {
   withLock_('SMOKE_V2', () => {
      const issues = [];
      const ss = getMasterSs_();

      function fail_(name, data) {
         issues.push({ name, data: data || {} });
         log_('ERROR','SMOKE_FAIL', name, data || {});
      }
      function ok_(name, data) {
         log_('INFO','SMOKE_OK', name, data || {});
      }

      // Sheets requerides
      Object.keys(HEADERS).forEach(n => {
         if (!ss.getSheetByName(n)) fail_('SHEET_MISSING', { sheet: n });
      });

      // Root folder
      try {
         const root = getOrCreateRootFolder_();
         ok_('ROOT_OK', { rootId: root.getId(), name: root.getName() });
      } catch (e) {
         fail_('ROOT_FAIL', { err: String(e) });
      }

      // Counts mínims
      try {
         const t = getSheet_(SHEETS.TECNICS);
         const c = getSheet_(SHEETS.CATALEGS);
         const s = getSheet_(SHEETS.SINONIMS);
         ok_('COUNTS', {
            tecnics: Math.max(0, t.getLastRow()-1),
            catalegs: Math.max(0, c.getLastRow()-1),
            sinonims: Math.max(0, s.getLastRow()-1)
         });
      } catch (e) {
         fail_('COUNT_FAIL', { err: String(e) });
      }

      // Healthcheck row
      const shH = getSheet_(SHEETS.HEALTH);
      const status = issues.length ? 'FAIL' : 'OK';
      const message = issues.length ? ('Errors: ' + issues.length) : 'Tot correcte (V2)';
      shH.appendRow([toIso_(new Date()), 'SMOKE_TEST_V2', status, message, JSON.stringify({ issues })]);

      if (issues.length) uiAlert_('Smoke test: FAIL ❌ (revisa LOG + HEALTHCHECK)');
      else uiAlert_('Smoke test: OK ✅');
   });
}


/**
 * EMI_DIAG_auditInputs_V7
 * - Verifica que les taules d'entrada (AppSheet) i la ingesta CV tenen els headers esperats.
 * - NO modifica dades; només escriu al LOG.
 */
function EMI_DIAG_auditInputs_V7() {
   withLock_('AUDIT_INPUTS_V7', () => {
      const ss = getMasterSs_();

      function auditSheet_(sheetName, required) {
         const sh = ss.getSheetByName(sheetName);
         if (!sh) {
            log_('ERROR','AUDIT_MISSING_SHEET','No existeix sheet', { sheet: sheetName });
            return;
         }
         const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(v => String(v||'').trim());
         const seen = {};
         const dup = [];
         header.forEach(h => { if (!h) return; if (seen[h]) dup.push(h); seen[h]=true; });

         const missing = required.filter(h => !seen[h]);
         log_('INFO','AUDIT_HEADERS', 'Audit headers', {
            sheet: sheetName,
            headerCount: header.filter(Boolean).length,
            missing: missing,
            duplicates: dup
         });
      }

      auditSheet_(SHEETS.PERSONES, [
         'id_persona','emi_id_persona','nom','cognom1','email','telefon_mobil',
         'municipi','situacio_laboral_actual','cv_sector_objectiu',
         'tags_persona','tags_persona_auto','tags_auto','text_index','cv_text_index',
         'cv_experiencia','cv_formacio','cv_idiomes','carnet_conduir','vehicle_propi'
      ]);

      auditSheet_(SHEETS.VACANTS, [
         'id_vacant','emi_id_vacant','estat','empresa_id','titol_vacant',
         'descripcio','requisits','sector','horari','jornada',
         'tags_vacant','tags_vacant_auto','tags_auto','text_index'
      ]);

      auditSheet_('CV_RAW', [
         'id_cv','source_row','source_ts','ingested_ts','email','nom','cognoms',
         'cv_text_index','carnet_conduir','vehicle_propi','raw_json','linked_id_persona','status'
      ]);

      // Audit CV SOURCE headers (read-only)
      try {
         if (typeof EMI_CV_getSourceSpreadsheetId_ === 'function') {
            const srcId = EMI_CV_getSourceSpreadsheetId_();
            const src = SpreadsheetApp.openById(srcId);
            const sh = src.getSheets()[0];
            const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(v => String(v||'').trim());
            const must = ['Marca temporal','Nom i Cognoms','Correu electrònic','Telèfon'];
            const miss = must.filter(x => h.indexOf(x) === -1);
            log_('INFO','AUDIT_CV_SOURCE','Audit CV source headers', { srcId: srcId, sheet: sh.getName(), missing: miss, headerCount: h.filter(Boolean).length });
         }
      } catch (e) {
         log_('WARN','AUDIT_CV_SOURCE_FAIL','No he pogut auditar el source del CV (pot ser permís/ID)', { err: String(e) });
      }
   });
}
