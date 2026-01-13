/*******************************************************
  * 06_TRIGGERS_IDS.gs — EMI BD MASTER (AppSheet) · V2
  * - Triggers instal·lables (onOpen, onEdit, CRON 5m, CRON diari)
  * - Assignació EMI IDs:
  *      - onEdit (manual Sheets) via assignEmiIdIfMissing_
  *      - batch CRON via EMI_assignMissingEmiIdsBatch
  *
  * ARREL FIX:
  * - Eliminem dependència de nextEmiId_()
  * - Fem servir el generador anual seqüencial existent:
  *       nextSeqFromPrefix_(PROP_KEYS.SEQ_..., year) + buildEmiId_(prefix, year, seq)
  *******************************************************/

/** =========================
  *   TRIGGERS handlers
  *   ========================= */

/** Handler ON_OPEN (instal·lable). */
function EMI_onOpen_V2(e) {
   try {
      if (typeof EMI_buildMenu_ === 'function') EMI_buildMenu_();
   } catch (err) {
      try { log_('WARN', 'ONOPEN_FAIL', 'Error al trigger onOpen', { err: String(err) }); } catch (_) {}
   }
}

/**
  * Handler ON_EDIT (instal·lable).
  * IMPORTANT: AppSheet sovint NO dispara onEdit (writes server-side).
  * Això és només per edicions manuals dins Sheets.
  */
function EMI_onEdit_V2(e) {
   try {
      if (!e || !e.range) return;

      var sh = e.range.getSheet();
      var name = sh.getName();
      var row = e.range.getRow();
      if (row <= 1) return;

      // PERSONES / EMPRESES / VACANTS: assigna EMI ID si falta
      if (name === SHEETS.PERSONES) {
         assignEmiIdIfMissing_(sh, row, 'emi_id_persona', 'PERSONA', 'P', 'data_ultima_actualitzacio');
      } else if (name === SHEETS.EMPRESES) {
         if (typeof maybeNormalizeEmpresa_ === 'function') maybeNormalizeEmpresa_(sh, row);
         assignEmiIdIfMissing_(sh, row, 'emi_id_empresa', 'EMPRESA', 'E', 'data_ultima_actualitzacio');
      } else if (name === SHEETS.VACANTS) {
         assignEmiIdIfMissing_(sh, row, 'emi_id_vacant', 'VACANT', 'V', 'updated_at');
      }

   } catch (err) {
      try { log_('WARN','ONEDIT_FAIL','Error al trigger onEdit', { err: String(err) }); } catch (_) {}
   }
}

/**
  * CRON cada 5 minuts: punt “robust”.
  * - Assigna EMI IDs a files sense emi_id_*
  * - (Opcional) pot cridar ingestió CV si existeix (més endavant)
  */
function EMI_cron5m_V2() {
   try {
      if (typeof EMI_assignMissingEmiIdsBatch === 'function') EMI_assignMissingEmiIdsBatch();
      if (typeof EMI_CV_ingestNewResponses === 'function') EMI_CV_ingestNewResponses();
   } catch (err) {
      try {
         log_('ERROR', 'CRON5M_FAIL', 'Error al CRON 5m', {
            err: String(err),
            stack: (err && err.stack) ? String(err.stack) : ''
         });
      } catch (_) {}
   }
}

/** CRON diari: manteniment lleuger + healthcheck (si existeix). */
function EMI_cronDaily_V2() {
   try {
      if (typeof EMI_smokeTest_V2 === 'function') {
         EMI_smokeTest_V2();
      } else {
         try { log_('INFO', 'CRON_DAILY_OK', 'CRON diari OK (sense smoke test)', {}); } catch (_) {}
      }
   } catch (err) {
      try { log_('ERROR', 'CRON_DAILY_FAIL', 'Error al CRON diari', { err: String(err) }); } catch (_) {}
   }
}

/**
 * Wrapper per poder programar el pipeline (triggers no poden passar paràmetres).
 * Fa: reindex + auto-tags (labels) + matching (2 sentits) amb topN=25
 */
function EMI_PIPELINE_rebuildBrainNow_25() {
   try {
      if (typeof EMI_PIPELINE_rebuildBrainNow === 'function') {
         EMI_PIPELINE_rebuildBrainNow(25);
      } else {
         log_('WARN','PIPELINE_MISSING','No existeix EMI_PIPELINE_rebuildBrainNow()', {});
      }
   } catch (err) {
      try { log_('ERROR','PIPELINE_TRIGGER_FAIL','Error al pipeline programat', { err: String(err), stack: (err && err.stack) ? String(err.stack) : '' }); } catch (_) {}
   }
}

/** =========================
  *   INSTALL / REMOVE / LIST
  *   ========================= */

function EMI_installTriggers_V2() {
   withLock_('INSTALL_TRIGGERS_V2', function () {
      var ss = null;
      try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (_) { ss = null; }
      if (!ss) ss = getMasterSs_();

      // Esborra triggers antics d’aquests handlers per no duplicar
      var handlers = ['EMI_onOpen_V2', 'EMI_onEdit_V2', 'EMI_cron5m_V2', 'EMI_cronDaily_V2', 'EMI_DICT_runHarvestBatch', 'EMI_MASTER_trimGhostColumnsNow'];
      ScriptApp.getProjectTriggers().forEach(function (t) {
         var h = t.getHandlerFunction();
         if (handlers.indexOf(h) >= 0) ScriptApp.deleteTrigger(t);
      });

      ScriptApp.newTrigger('EMI_onOpen_V2').forSpreadsheet(ss).onOpen().create();
      ScriptApp.newTrigger('EMI_onEdit_V2').forSpreadsheet(ss).onEdit().create();
      ScriptApp.newTrigger('EMI_cron5m_V2').timeBased().everyMinutes(5).create();
      ScriptApp.newTrigger('EMI_cronDaily_V2').timeBased().atHour(2).nearMinute(10).everyDays(1).create();
    ScriptApp.newTrigger('EMI_DICT_runHarvestBatch').timeBased().everyMinutes(30).create();
    ScriptApp.newTrigger('EMI_MASTER_trimGhostColumnsNow').timeBased().atHour(3).nearMinute(5).everyDays(1).create();

      log_('INFO', 'TRIGGERS_INSTALLED', 'Triggers instal·lats (onOpen+onEdit+cron5m+daily+harvest+trimGhost)', { ssId: ss.getId() });
      uiAlert_('Triggers instal·lats ✅\n\n- onOpen (menú)\n- onEdit (manual)\n- CRON 5m (AppSheet robust)\n- CRON diari\n- Harvest diccionari (cada 30 min)\n- Neteja columnes fantasma (diari)');
   });
}


/**
 * Instal·lació recomanada (V3):
 * - Manté els triggers “mínims” per fer que el sistema vagi sol, sense bucles.
 * - IMPORTANT: AppSheet NO dispara onEdit, per això necessitem un cron periòdic.
 *
 * Inclou:
 * - onOpen (menú)
 * - onEdit (només edicions manuals al Sheets)
 * - cron 10 min: assigna EMI IDs + ingesta de CV (si n’hi ha)
 * - harvest diccionari: cada 30 min
 * - pipeline (reindex + matching): cada 2 hores (topN=25)
 * - trim columnes fantasma: diari
 * - smoke test: diari
 */
function EMI_installTriggers_V3() {
   withLock_('INSTALL_TRIGGERS_V3', function () {
      var ss = null;
      try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (_) { ss = null; }
      if (!ss) ss = getMasterSs_();

      // 1) Elimina triggers d’aquests handlers per evitar duplicats i bucles antics
      var handlers = [
         'EMI_onOpen_V2','EMI_onEdit_V2',
         'EMI_cron5m_V2','EMI_cronDaily_V2',
         'EMI_DICT_runHarvestBatch',
         'EMI_MASTER_trimGhostColumnsNow',
         'EMI_PIPELINE_rebuildBrainNow_25',
         // per si ha quedat algun trigger “dels vells”
         'EMI_recomputeMatchesBothSidesNow',
         'EMI_recomputeMatchesNow',
         'EMI_recomputeVacantMatchesNow',
         'EMI_PIPELINE_rebuildBrainNow'
      ];

      ScriptApp.getProjectTriggers().forEach(function (t) {
         var h = t.getHandlerFunction();
         if (handlers.indexOf(h) >= 0) ScriptApp.deleteTrigger(t);
      });

      // 2) Crea triggers “bons”
      ScriptApp.newTrigger('EMI_onOpen_V2').forSpreadsheet(ss).onOpen().create();
      ScriptApp.newTrigger('EMI_onEdit_V2').forSpreadsheet(ss).onEdit().create();

      // Cron “bot” (AppSheet)
      ScriptApp.newTrigger('EMI_cron5m_V2').timeBased().everyMinutes(10).create();

      // Manteniment diari (smoke test)
      ScriptApp.newTrigger('EMI_cronDaily_V2').timeBased().atHour(2).nearMinute(10).everyDays(1).create();

      // Diccionari (creixement progressiu)
      ScriptApp.newTrigger('EMI_DICT_runHarvestBatch').timeBased().everyMinutes(30).create();

      // Pipeline (matching) – regular però no constant
      ScriptApp.newTrigger('EMI_PIPELINE_rebuildBrainNow_25').timeBased().everyHours(2).create();

      // Neteja columnes fantasma
      ScriptApp.newTrigger('EMI_MASTER_trimGhostColumnsNow').timeBased().atHour(3).nearMinute(5).everyDays(1).create();

      log_('INFO', 'TRIGGERS_INSTALLED_V3', 'Triggers instal·lats V3', { ssId: ss.getId() });
      uiAlert_(
         'Triggers V3 instal·lats ✅\n\n' +
         '- onOpen (menú)\n' +
         '- onEdit (manual Sheets)\n' +
         '- cron 10 min (IDs + CV)\n' +
         '- Harvest diccionari (30 min)\n' +
         '- Pipeline (cada 2 hores)\n' +
         '- Trim columnes fantasma (diari)\n' +
         '- Smoke test (diari)'
      );
   });
}


function EMI_removeTriggers_V2() {
   withLock_('REMOVE_TRIGGERS_V2', function () {
      var handlers = ['EMI_onOpen_V2', 'EMI_onEdit_V2', 'EMI_cron5m_V2', 'EMI_cronDaily_V2', 'EMI_DICT_runHarvestBatch', 'EMI_MASTER_trimGhostColumnsNow'];
      var removed = 0;

      ScriptApp.getProjectTriggers().forEach(function (t) {
         if (handlers.indexOf(t.getHandlerFunction()) >= 0) {
            ScriptApp.deleteTrigger(t);
            removed++;
         }
      });

      log_('INFO', 'TRIGGERS_REMOVED', 'Triggers eliminats', { removed: removed });
      uiAlert_('Triggers eliminats ✅ (removed=' + removed + ')');
   });
}

function EMI_listTriggers_V2() {
   var out = [];
   ScriptApp.getProjectTriggers().forEach(function (t) {
      out.push({
         handler: t.getHandlerFunction(),
         type: String(t.getEventType ? t.getEventType() : ''),
         source: String(t.getTriggerSource ? t.getTriggerSource() : '')
      });
   });
   log_('INFO', 'TRIGGERS_LIST', 'Triggers actuals', { count: out.length, triggers: out });
   uiAlert_('Triggers llistats al LOG ✅ (count=' + out.length + ')');
}

/** =========================
  *   EMI ID ASSIGNMENT (root fix)
  *   ========================= */

/**
  * Assigna EMI ID a una fila concreta si està buit.
  * Firma que ja tens referenciada des de EMI_onEdit_V2:
  *    (sh, row, emiColName, seqType, prefixLetter, tsColName)
  */
function assignEmiIdIfMissing_(sh, row, emiColName, seqType, prefixLetter, tsColName) {
   if (!sh || !row || row <= 1) return;

   // Validem que existeix el generador anual
   if (typeof nextSeqFromPrefix_ !== 'function' || typeof buildEmiId_ !== 'function' || typeof PROP_KEYS === 'undefined') {
      try {
         log_('ERROR', 'EMI_ID_GEN_MISSING', 'Falten nextSeqFromPrefix_/buildEmiId_/PROP_KEYS (revisa mòdul IDs)', {
            nextSeqFromPrefix_: typeof nextSeqFromPrefix_,
            buildEmiId_: typeof buildEmiId_,
            hasPropKeys: (typeof PROP_KEYS !== 'undefined')
         });
      } catch (_) {}
      return;
   }

   var map = hdrMap1Based_(sh);
   var emiCol = map[emiColName];
   var tsCol = tsColName ? map[tsColName] : null;
   if (!emiCol) return;

   var cur = String(sh.getRange(row, emiCol).getValue() || '').trim();
   if (cur) return;

   var year = Number(Utilities.formatDate(new Date(), EMI_TZ, 'yyyy'));
   var gen = getEmiGenConfig_(seqType, prefixLetter);
   if (!gen) return;

   var seq = nextSeqFromPrefix_(gen.seqPropPrefix, year);
   var code = buildEmiId_(gen.prefixText, year, seq);

   sh.getRange(row, emiCol).setValue(code);
   if (tsCol) {
      try { sh.getRange(row, tsCol).setValue((typeof toIso_ === 'function') ? toIso_(new Date()) : new Date()); } catch (_) {}
   }

   try { log_('INFO', 'ASSIGN_EMI_ID_ROW', 'EMI ID assignat a fila', { sheet: sh.getName(), row: row, emi: code }); } catch (_) {}
}

/**
  * Batch robust per assignar emi_id_* (pensat per AppSheet).
  * - Processa fins a N assignacions per taula per execució (evita timeouts)
  * - No toca files amb emi_id ja omplert
  */
function EMI_assignMissingEmiIdsBatch() {
   withLock_('CRON_ASSIGN_EMI_IDS', function () {
      var ss = getMasterSs_();

      var maxPerSheet = 200; // ajustable
      var res = {
         persones: assignMissingInSheet_(ss, SHEETS.PERSONES, 'emi_id_persona', 'PERSONA', 'P', 'data_ultima_actualitzacio', maxPerSheet, null),
         empreses: assignMissingInSheet_(ss, SHEETS.EMPRESES, 'emi_id_empresa', 'EMPRESA', 'E', 'data_ultima_actualitzacio', maxPerSheet, 'nom_empresa_normalitzat'),
         vacants:   assignMissingInSheet_(ss, SHEETS.VACANTS,   'emi_id_vacant',   'VACANT',   'V', 'updated_at',                      maxPerSheet, null)
      };

      log_('INFO', 'CRON_EMI_IDS_DONE', 'Batch EMI IDs executat', res);
   });
}

/**
  * Assignació batch dins una sheet.
  * normalizeColName opcional: si existeix i està buit, el calcula (per EMPRESES).
  */
function assignMissingInSheet_(ss, sheetName, emiColName, seqType, prefixLetter, tsColName, maxRows, normalizeColName) {
   var sh = ss.getSheetByName(sheetName);
   if (!sh) return { sheet: sheetName, processed: 0, assigned: 0, note: 'missing_sheet' };

   var lastRow = sh.getLastRow();
   var lastCol = sh.getLastColumn();
   if (lastRow <= 1) return { sheet: sheetName, processed: 0, assigned: 0, note: 'empty' };

   // Validem generador anual (arrel)
   if (typeof nextSeqFromPrefix_ !== 'function' || typeof buildEmiId_ !== 'function' || typeof PROP_KEYS === 'undefined') {
      try {
         log_('ERROR', 'EMI_ID_GEN_MISSING', 'Falten nextSeqFromPrefix_/buildEmiId_/PROP_KEYS (revisa mòdul IDs)', {
            sheet: sheetName,
            nextSeqFromPrefix_: typeof nextSeqFromPrefix_,
            buildEmiId_: typeof buildEmiId_,
            hasPropKeys: (typeof PROP_KEYS !== 'undefined')
         });
      } catch (_) {}
      return { sheet: sheetName, processed: 0, assigned: 0, note: 'missing_id_generator' };
   }

   var map = hdrMap1Based_(sh);
   var emiCol = map[emiColName];
   var tsCol = map[tsColName]; // pot no existir
   var normCol = normalizeColName ? map[normalizeColName] : null;
   var nameCol = (sheetName === SHEETS.EMPRESES) ? map['nom_empresa'] : null;

   if (!emiCol) return { sheet: sheetName, processed: 0, assigned: 0, note: 'missing_emi_col', emiColName: emiColName };

   var data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
   var processed = 0;
   var assigned = 0;

   var year = Number(Utilities.formatDate(new Date(), EMI_TZ, 'yyyy'));
   var gen = getEmiGenConfig_(seqType, prefixLetter);
   if (!gen) return { sheet: sheetName, processed: 0, assigned: 0, note: 'missing_gen_cfg' };

   for (var i = 0; i < data.length; i++) {
      if (assigned >= maxRows) break;

      var rowIdx = i + 2;
      var curEmi = String(data[i][emiCol - 1] || '').trim();
      if (curEmi) continue;

      // opcional: normalitza empresa
      if (sheetName === SHEETS.EMPRESES && normCol && nameCol) {
         var curNorm = String(data[i][normCol - 1] || '').trim();
         var nm = String(data[i][nameCol - 1] || '').trim();
         if (!curNorm && nm) {
            var norm = (typeof normalizeText_ === 'function') ? normalizeText_(nm) : String(nm).toLowerCase();
            if (norm) {
               sh.getRange(rowIdx, normCol).setValue(norm);
               data[i][normCol - 1] = norm;
            }
         }
      }

      var seq = nextSeqFromPrefix_(gen.seqPropPrefix, year);
      var code = buildEmiId_(gen.prefixText, year, seq);

      sh.getRange(rowIdx, emiCol).setValue(code);
      if (tsCol) sh.getRange(rowIdx, tsCol).setValue((typeof toIso_ === 'function') ? toIso_(new Date()) : new Date());

      assigned++;
      processed++;
   }

   return {
      sheet: sheetName,
      processed: processed,
      assigned: assigned,
      tsColName: tsColName,
      tsColFound: Boolean(tsCol)
   };
}

/**
  * Mapeig de seqType -> (PROP_KEYS prefix + text prefix)
  * Exemple: PERSONA -> 'EMI_SEQ_PERSONA_' + 'EMI-P'
  */
function getEmiGenConfig_(seqType, prefixLetter) {
   var t = String(seqType || '').trim().toUpperCase();

   if (t === 'PERSONA') return { seqPropPrefix: PROP_KEYS.SEQ_PERSONA_PREFIX, prefixText: 'EMI-P' };
   if (t === 'EMPRESA') return { seqPropPrefix: PROP_KEYS.SEQ_EMPRESA_PREFIX, prefixText: 'EMI-E' };
   if (t === 'VACANT' || t === 'VACANTS') return { seqPropPrefix: PROP_KEYS.SEQ_VACANT_PREFIX, prefixText: 'EMI-V' };

   try { log_('ERROR', 'EMI_ID_BAD_TYPE', 'seqType desconegut per EMI ID', { seqType: seqType, prefixLetter: prefixLetter }); } catch (_) {}
   return null;
}

/**
  * Header map 1-based: { headerName -> colNumber }.
  * (Evita dependències i confusions entre versions).
  */
function hdrMap1Based_(sh) {
   var lastCol = Math.max(1, sh.getLastColumn());
   var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
   var m = {};
   for (var i = 0; i < hdr.length; i++) {
      var h = hdr[i];
      if (h) m[h] = i + 1;
   }
   return m;
}

/*******************************************************
 * HARD TRIGGER REMOVAL (maintenance/install only)
 * - Removes ALL installable triggers for current user.
 * - Use when EMI_removeTriggers_V2() leaves protected triggers.
 *******************************************************/
function EMI_removeTriggers_HARD() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  const failures = [];
  triggers.forEach(function(t) {
    try {
      ScriptApp.deleteTrigger(t);
      removed++;
    } catch (err) {
      failures.push({
        handler: (t && t.getHandlerFunction) ? t.getHandlerFunction() : null,
        source: (t && t.getTriggerSource) ? String(t.getTriggerSource()) : null,
        type: (t && t.getEventType) ? String(t.getEventType()) : null,
        err: String(err)
      });
    }
  });
  const remaining = ScriptApp.getProjectTriggers().map(function(t) {
    return {
      handler: t.getHandlerFunction(),
      source: String(t.getTriggerSource()),
      type: String(t.getEventType())
    };
  });
  if (typeof EMI_LOG_ === 'function') {
    EMI_LOG_('INFO', 'TRIGGERS_REMOVED_HARD', 'Triggers eliminats (HARD)', {
      removed: removed,
      failures: failures,
      remainingCount: remaining.length,
      remaining: remaining
    });
  } else {
    Logger.log(JSON.stringify({ removed: removed, failures: failures, remaining: remaining }));
  }
}
