/*******************************************************
  * 12_EMI_IDS.gs — Assignació EMI_ID anual seqüencial
  * - Escaneja PERSONES/EMPRESES/VACANTS i omple emi_id_* buits
  *******************************************************/

function EMI_assignMissingEmiIds() {
   return withLock_('ASSIGN_EMI_IDS', () => {
      const res = {
         persones: assignFor_(SHEETS.PERSONES, 'emi_id_persona', PROP_KEYS.SEQ_PERSONA_PREFIX, 'EMI-P'),
         empreses: assignFor_(SHEETS.EMPRESES, 'emi_id_empresa', PROP_KEYS.SEQ_EMPRESA_PREFIX, 'EMI-E'),
         vacants:   assignFor_(SHEETS.VACANTS,   'emi_id_vacant',   PROP_KEYS.SEQ_VACANT_PREFIX,   'EMI-V')
      };

      log_('INFO', 'ASSIGN_EMI_IDS_OK', 'EMI IDs assignats (si calia)', res);
      return res;
   });
}

function assignFor_(sheetName, emiColName, seqPropPrefix, prefixText) {
   const sh = getSheet_(sheetName);
   const lastRow = sh.getLastRow();
   if (lastRow <= 1) return { filled: 0, total: 0 };

   const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
   const colEmi = header.indexOf(emiColName) + 1;
   if (!colEmi) throw new Error(sheetName + ' no té columna ' + emiColName);

   const year = Number(Utilities.formatDate(new Date(), EMI_TZ, 'yyyy'));
   const rng = sh.getRange(2, colEmi, lastRow - 1, 1);
   const vals = rng.getValues();

   let filled = 0;
   for (let i = 0; i < vals.length; i++) {
      const cur = String(vals[i][0] || '').trim();
      if (cur) continue;

      const seq = nextSeqFromPrefix_(seqPropPrefix, year);
      vals[i][0] = buildEmiId_(prefixText, year, seq);
      filled++;
   }

   if (filled) rng.setValues(vals);
   return { filled, total: vals.length };
}

function nextSeqFromPrefix_(seqPropPrefix, year) {
   const k = String(seqPropPrefix) + String(year);
   const cur = Number(getProp_(k) || '0');
   const next = cur + 1;
   setProp_(k, String(next));
   return next;
}

function buildEmiId_(prefixText, year, seq) {
   const s = String(Math.max(0, Number(seq || 0)));
   const padded = ('000000' + s).slice(-6);
   return String(prefixText) + '-' + String(year) + '-' + padded;
}