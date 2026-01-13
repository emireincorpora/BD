/*******************************************************
  * 09_DROPDOWNS_VALIDATION.gs — EMI BD MASTER (AppSheet) · Dropdowns via Sheets Data Validation
  * - Crea/actualitza la sheet interna _DV_LISTS amb llistes canòniques
  * - Aplica Data Validation (dropdown) a columnes clau
  *
  * IMPORTANT:
  * - AppSheet pot crear dropdowns a partir de Data Validation de Sheets,
  *    però quan canviïs les validacions, a AppSheet hauràs de fer "Regenerate"
  *    de les taules perquè ho detecti.
  *******************************************************/

const EMI_DV = {
   SHEET_NAME: '_DV_LISTS',
   MAX_ROWS_APPLY: 5000, // fins on apliquem validació (files)
};

function EMI_refreshDropdowns() {
   withLock_('REFRESH_DROPDOWNS', () => {
      const ss = getMasterSs_();

      const lists = buildDvLists_(ss); // { sheet, ranges: {name: Range} }
      applyDvToCoreTables_(ss, lists.ranges);

      // Amaguem la sheet interna sempre
      try {
         const sh = ss.getSheetByName(EMI_DV.SHEET_NAME);
         if (sh) sh.hideSheet();
      } catch (_) {}

      SpreadsheetApp.flush();
      log_('INFO', 'DV_REFRESH_OK', 'Dropdowns/validacions actualitzats', {
         dvSheet: EMI_DV.SHEET_NAME,
         ranges: Object.keys(lists.ranges),
      });

      uiAlert_(
         "Dropdowns actualitzats ✅\n\n" +
         "NOTA AppSheet: ves a Data→Tables i fes Regenerate de PERSONES/EMPRESES/VACANTS " +
         "si vols que l'app detecti els nous desplegables."
      );
   });
}

/** =========================
  *   Build list sheet
  *   ========================= */
function buildDvLists_(ss) {
   let dvSh = ss.getSheetByName(EMI_DV.SHEET_NAME);
   if (!dvSh) dvSh = ss.insertSheet(EMI_DV.SHEET_NAME);
   dvSh.clearContents();

   // Capçaleres (una columna per llista)
   const cols = [
      { key: 'TECNICS_ACTIUS', header: 'TECNICS_ACTIUS (id_tecnic)' },
      { key: 'SECTOR_EMPRESA', header: 'SECTOR_EMPRESA (valor)' },
      { key: 'AMBIT_FORMACIO', header: 'AMBIT_FORMACIO (valor)' },
      { key: 'PROGRAMA_EMI', header: 'PROGRAMA_EMI (valor)' },
      { key: 'TAG_PERSONA', header: 'TAG_PERSONA (valor)' },
      { key: 'TAG_EMPRESA', header: 'TAG_EMPRESA (valor)' },
   ];

   dvSh.getRange(1, 1, 1, cols.length).setValues([cols.map(c => c.header)]);

   // Valors
   const tecnics = getActiveTecnicsIds_();                            // ids
   const sectors = getCatalogValuesByTipus_('SECTOR_EMPRESA'); // valors
   const ambits   = getCatalogValuesByTipus_('AMBIT_FORMACIO');
   const progr    = getCatalogValuesByTipus_('PROGRAMA_EMI');
   const tagsP    = getCatalogValuesByTipus_('TAG_PERSONA');
   const tagsE    = getCatalogValuesByTipus_('TAG_EMPRESA');

   const valuesByKey = {
      TECNICS_ACTIUS: tecnics,
      SECTOR_EMPRESA: sectors,
      AMBIT_FORMACIO: ambits,
      PROGRAMA_EMI: progr,
      TAG_PERSONA: tagsP,
      TAG_EMPRESA: tagsE,
   };

   const ranges = {};
   cols.forEach((c, idx) => {
      const col = idx + 1;
      const vals = valuesByKey[c.key] || [];
      ranges[c.key] = writeListColumn_(dvSh, col, vals);
   });

   // Deixa-la amagada
   try { dvSh.hideSheet(); } catch (_) {}

   log_('INFO', 'DV_LISTS_BUILT', 'Llistes EMI_DV generades', {
      tecnics: tecnics.length,
      sectors: sectors.length,
      ambits: ambits.length,
      programes: progr.length,
      tagPersona: tagsP.length,
      tagEmpresa: tagsE.length,
   });

   return { sheet: dvSh, ranges };
}

function writeListColumn_(sh, col, values) {
   // Escriu valors a partir de la fila 2
   const clean = (values || [])
      .map(v => String(v || '').trim())
      .filter(v => v);

   // Si està buit, posem un blanc a A2 perquè no peti requireValueInRange
   if (!clean.length) {
      sh.getRange(2, col, 1, 1).setValue('');
      return sh.getRange(2, col, 1, 1);
   }

   const out = clean.map(v => [v]);
   sh.getRange(2, col, out.length, 1).setValues(out);
   return sh.getRange(2, col, out.length, 1);
}

/** =========================
  *   Read sources
  *   ========================= */
function getActiveTecnicsIds_() {
   const sh = getSheet_(SHEETS.TECNICS);
   const map = headerMap_(sh);
   const colId = map['id_tecnic'];
   const colActiu = map['actiu'];
   if (!colId || !colActiu) {
      log_('WARN', 'DV_TECNICS_HEADERS_MISSING', 'Falten headers a TECNICS', { colId, colActiu });
      return [];
   }

   const last = sh.getLastRow();
   if (last <= 1) return [];

   const data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
   const out = [];
   data.forEach(r => {
      const id = String(r[colId - 1] || '').trim();
      const act = String(r[colActiu - 1] || '').trim().toUpperCase();
      if (id && act === 'SI') out.push(id);
   });

   // ordre estable
   return out;
}

function getCatalogValuesByTipus_(tipus) {
   const sh = getSheet_(SHEETS.CATALEGS);
   const map = headerMap_(sh);
   const colTipus = map['tipus'];
   const colValor = map['valor'];
   const colActiu = map['actiu'];
   const colOrdre = map['ordre'];

   if (!colTipus || !colValor || !colActiu) {
      log_('WARN', 'DV_CATALEGS_HEADERS_MISSING', 'Falten headers a CATALEGS', { colTipus, colValor, colActiu, colOrdre });
      return [];
   }

   const last = sh.getLastRow();
   if (last <= 1) return [];

   const data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();

   const rows = [];
   data.forEach(r => {
      const t = String(r[colTipus - 1] || '').trim();
      const v = String(r[colValor - 1] || '').trim();
      const a = String(r[colActiu - 1] || '').trim().toUpperCase();
      const o = colOrdre ? Number(r[colOrdre - 1] || 9999) : 9999;
      if (t === tipus && a === 'SI' && v) rows.push({ v, o });
   });

   rows.sort((a, b) => (a.o - b.o) || a.v.localeCompare(b.v, 'ca'));
   return rows.map(x => x.v);
}

/** =========================
  *   Apply validations
  *   ========================= */
function applyDvToCoreTables_(ss, ranges) {
   // Single-value (strict): allowInvalid=false
   // Multi-value (EnumList-ish): allowInvalid=true (per no bloquejar AppSheet)
   const specs = [
      // PERSONES
      { sheet: SHEETS.PERSONES, col: 'tecnic_referent', list: ranges.TECNICS_ACTIUS, allowInvalid: false, help: 'Selecciona id_tecnic (actiu)' },
      { sheet: SHEETS.PERSONES, col: 'ambits_interes', list: ranges.AMBIT_FORMACIO, allowInvalid: true,   help: 'Suggeriments (multi). Separar amb coma si cal.' },
      { sheet: SHEETS.PERSONES, col: 'programes_interes', list: ranges.PROGRAMA_EMI, allowInvalid: true, help: 'Suggeriments (multi). Separar amb coma si cal.' },
      { sheet: SHEETS.PERSONES, col: 'tags_persona', list: ranges.TAG_PERSONA, allowInvalid: true, help: 'Suggeriments (multi). Separar amb coma si cal.' },

      // EMPRESES
      { sheet: SHEETS.EMPRESES, col: 'tecnic_referent', list: ranges.TECNICS_ACTIUS, allowInvalid: false, help: 'Selecciona id_tecnic (actiu)' },
      { sheet: SHEETS.EMPRESES, col: 'sector_principal', list: ranges.SECTOR_EMPRESA, allowInvalid: false, help: 'Sector principal (catàleg)' },
      { sheet: SHEETS.EMPRESES, col: 'altres_sectors', list: ranges.SECTOR_EMPRESA, allowInvalid: true, help: 'Altres sectors (multi). Separar amb coma si cal.' },
      { sheet: SHEETS.EMPRESES, col: 'tags_empresa', list: ranges.TAG_EMPRESA, allowInvalid: true, help: 'Suggeriments (multi). Separar amb coma si cal.' },

      // VACANTS
      { sheet: SHEETS.VACANTS, col: 'tecnic_referent', list: ranges.TECNICS_ACTIUS, allowInvalid: false, help: 'Selecciona id_tecnic (actiu)' },
      { sheet: SHEETS.VACANTS, col: 'sector', list: ranges.SECTOR_EMPRESA, allowInvalid: false, help: 'Sector (catàleg)' },
      { sheet: SHEETS.VACANTS, col: 'tags_vacant', list: ranges.SECTOR_EMPRESA, allowInvalid: true, help: 'Suggeriments (multi). Separar amb coma si cal.' },
   ];

   const applied = [];
   specs.forEach(sp => {
      const sh = ss.getSheetByName(sp.sheet);
      if (!sh) {
         log_('WARN', 'DV_SHEET_MISSING', 'No existeix la sheet per aplicar EMI_DV', { sheet: sp.sheet });
         return;
      }
      const ok = setValidationForColumn_(sh, sp.col, sp.list, sp.allowInvalid, sp.help);
      if (ok) applied.push({ sheet: sp.sheet, col: sp.col, allowInvalid: sp.allowInvalid });
   });

   log_('INFO', 'DV_APPLIED', 'Validacions aplicades', { count: applied.length, applied });
}

function setValidationForColumn_(sh, colName, listRange, allowInvalid, helpText) {
   const map = headerMap_(sh);
   const col = map[colName];
   if (!col) {
      log_('WARN', 'DV_COL_MISSING', 'No trobo columna per aplicar EMI_DV', { sheet: sh.getName(), colName });
      return false;
   }

   // On aplicar-ho: des de fila 2 fins a MAX_ROWS_APPLY o fins a lastRow (el que sigui més gran)
   const lastRow = Math.max(sh.getLastRow(), 2);
   const endRow = Math.max(lastRow, 2 + EMI_DV.MAX_ROWS_APPLY);
   const numRows = endRow - 1;

   const target = sh.getRange(2, col, numRows, 1);

   // Builder
   let builder = SpreadsheetApp.newDataValidation()
      .requireValueInRange(listRange, true)
      .setAllowInvalid(Boolean(allowInvalid));

   if (helpText) builder = builder.setHelpText(String(helpText));

   target.setDataValidation(builder.build());
   return true;
}