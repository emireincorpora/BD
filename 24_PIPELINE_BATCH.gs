/*******************************************************
  * 24_PIPELINE_BATCH.gs
  * - Pipeline batch + resumable (reindex + auto-tags + matching)
  * - Estat persistent a PropertiesService + STATE sheet
  *******************************************************/

const EMI_PIPELINE_STATE_KEY = 'EMI_PIPELINE_STATE_V1';
const EMI_PIPELINE_MAX_ROWS_PER_RUN = 150;
const EMI_PIPELINE_MAX_VACANTS_PER_RUN = 15;
const EMI_PIPELINE_MAX_MS = 4.5 * 60 * 1000; // 4.5 min (safe margin)

function EMI_PIPELINE_rebuildBrainNow(topN) {
   return withLock_('PIPELINE_REBUILD', () => {
      const ss = getMasterSs_();
      ensureBrainColumns_(ss);
      ensureVacantSuggerimentsSheet_(ss);
      EMI_PIPELINE_ensureStateSheet_(ss);

      const state = {
         phase: 'reindex_persones',
         cursor: 2,
         topN: Number(topN || 25),
         updated_at: toIso_(new Date())
      };

      EMI_PIPELINE_setState_(state);
      EMI_PIPELINE_logState_('PIPELINE_START', state);

      return EMI_PIPELINE_runBatch_();
   });
}

function EMI_PIPELINE_runNext() {
   return withLock_('PIPELINE_RESUME', () => EMI_PIPELINE_runBatch_());
}

function EMI_PIPELINE_runBatch_() {
   const start = Date.now();
   const ss = getMasterSs_();
   const dict = buildBrainDictionaries_();
   const cfg = readMatchConfig_(dict.catalogs);

   let state = EMI_PIPELINE_getState_();
   if (!state || !state.phase) {
      EMI_PIPELINE_logState_('PIPELINE_NO_STATE', { state });
      return { status: 'idle' };
   }

   let done = false;

   while (!done && (Date.now() - start) < EMI_PIPELINE_MAX_MS) {
      switch (state.phase) {
         case 'reindex_persones':
            state = EMI_PIPELINE_runReindexPhase_(ss, dict, SHEETS.PERSONES, buildPersonaBaseText_, state, 'reindex_empreses');
            break;
         case 'reindex_empreses':
            state = EMI_PIPELINE_runReindexPhase_(ss, dict, SHEETS.EMPRESES, buildEmpresaBaseText_, state, 'reindex_vacants');
            break;
         case 'reindex_vacants':
            state = EMI_PIPELINE_runReindexPhase_(ss, dict, SHEETS.VACANTS, buildVacantBaseText_, state, 'labels_persones');
            break;
         case 'labels_persones':
            state = EMI_PIPELINE_runLabelsPhase_(ss, dict.catalogs, SHEETS.PERSONES, 'tags_persona_auto', state, 'labels_vacants');
            break;
         case 'labels_vacants':
            state = EMI_PIPELINE_runLabelsPhase_(ss, dict.catalogs, SHEETS.VACANTS, 'tags_vacant_auto', state, 'matching_vacants');
            break;
         case 'matching_vacants':
            state = EMI_PIPELINE_runMatchingVacantsPhase_(ss, dict, cfg, state, 'done');
            break;
         case 'done':
            done = true;
            break;
         default:
            EMI_PIPELINE_logState_('PIPELINE_UNKNOWN_PHASE', state);
            done = true;
            break;
      }
   }

   if (state.phase === 'done') {
      EMI_PIPELINE_setState_({ phase: 'done', cursor: 0, topN: state.topN, updated_at: toIso_(new Date()) });
      EMI_PIPELINE_logState_('PIPELINE_DONE', state);
      return { status: 'done' };
   }

   EMI_PIPELINE_setState_(state);
   EMI_PIPELINE_scheduleNextRun_();
   EMI_PIPELINE_logState_('PIPELINE_CONTINUE', state);

   return { status: 'scheduled', state };
}

function EMI_PIPELINE_runReindexPhase_(ss, dict, sheetName, baseFn, state, nextPhase) {
   const cursor = Number(state.cursor || 2);
   const res = reindexSheetBatch_(ss, sheetName, dict, baseFn, cursor, EMI_PIPELINE_MAX_ROWS_PER_RUN);

   const nextState = {
      phase: res.done ? nextPhase : state.phase,
      cursor: res.done ? 2 : res.nextCursor,
      topN: state.topN,
      updated_at: toIso_(new Date())
   };

   EMI_PIPELINE_logState_('PIPELINE_REINDEX', { sheetName, cursor, res, nextState });
   return nextState;
}

function EMI_PIPELINE_runLabelsPhase_(ss, catalogs, sheetName, outColName, state, nextPhase) {
   const cursor = Number(state.cursor || 2);
   const res = fillAutoTagLabelsBatch_(ss, sheetName, outColName, catalogs, cursor, EMI_PIPELINE_MAX_ROWS_PER_RUN);

   const nextState = {
      phase: res.done ? nextPhase : state.phase,
      cursor: res.done ? 2 : res.nextCursor,
      topN: state.topN,
      updated_at: toIso_(new Date())
   };

   EMI_PIPELINE_logState_('PIPELINE_LABELS', { sheetName, cursor, res, nextState });
   return nextState;
}

function EMI_PIPELINE_runMatchingVacantsPhase_(ss, dict, cfg, state, nextPhase) {
   const cursor = Number(state.cursor || 0);
   const topN = Number(state.topN || 25);

   const persones = loadEntitiesForMatch_(ss, SHEETS.PERSONES, 'id_persona', dict, buildPersonaBaseText_);
   const vacants = loadEntitiesForMatch_(ss, SHEETS.VACANTS, 'id_vacant', dict, buildVacantBaseText_);

   if (!vacants.length || !persones.length) {
      return { phase: nextPhase, cursor: 0, topN, updated_at: toIso_(new Date()) };
   }

   const res = computeVacantMatchesBatch_(vacants, persones, cfg, dict.catalogs, cursor, EMI_PIPELINE_MAX_VACANTS_PER_RUN, topN);

   upsertVacantSuggestionsBatch_(ss, res.vacantIds, res.rows);

   const nextState = {
      phase: res.done ? nextPhase : state.phase,
      cursor: res.done ? 0 : res.nextCursor,
      topN: topN,
      updated_at: toIso_(new Date())
   };

   EMI_PIPELINE_logState_('PIPELINE_MATCHING', { cursor, res, nextState });
   return nextState;
}

function reindexSheetBatch_(ss, sheetName, dict, baseTextFn, startRow, maxRows) {
   const sh = ss.getSheetByName(sheetName);
   if (!sh || sh.getLastRow() < 2) return { done: true, processed: 0, nextCursor: 2 };

   const lastCol = Math.min(Math.max(1, sh.getLastColumn()), EMI_BRAIN_MAX_HEADER_COLS);
   const headersRaw = sh.getRange(1, 1, 1, lastCol).getValues()[0];
   const headers = headersRaw.map(h => String(h || '').trim());
   const idx = indexMap2_(headers);

   const tagsAutoC = idx['tags_auto'];
   const textIndexC = idx['text_index'];

   if (!tagsAutoC || !textIndexC) {
      log_('WARN','REINDEX_BATCH_MISSING_COLS','Falten columnes tags_auto/text_index', { sheetName });
      return { done: true, processed: 0, nextCursor: 2 };
   }

   const lastHeaderCol = Math.max(
      EMI_lastHeaderCol_(headers),
      Number(tagsAutoC || 0),
      Number(textIndexC || 0)
   );

   const lr = sh.getLastRow();
   const start = Math.max(2, Number(startRow || 2));
   if (start > lr) return { done: true, processed: 0, nextCursor: 2 };

   const end = Math.min(lr, start + maxRows - 1);
   const rng = sh.getRange(start, 1, end - start + 1, lastHeaderCol);
   const values = rng.getValues();

   let changed = 0;
   for (let r = 0; r < values.length; r++) {
      const row = values[r];
      const base = baseTextFn(row, idx, headers);
      const out = computeIndexAndTags_(base, dict);

      const prevTags = String(row[tagsAutoC - 1] || '');
      const prevIdx  = String(row[textIndexC - 1] || '');

      if (prevTags !== out.tags_auto) { row[tagsAutoC - 1] = out.tags_auto; changed++; }
      if (prevIdx  !== out.text_index) { row[textIndexC - 1] = out.text_index; changed++; }
   }

   if (changed) rng.setValues(values);

   const done = end >= lr;
   return { done, processed: values.length, changed, nextCursor: done ? 2 : end + 1 };
}

function fillAutoTagLabelsBatch_(ss, sheetName, outColName, catalogs, startRow, maxRows) {
   const sh = ss.getSheetByName(sheetName);
   if (!sh || sh.getLastRow() < 2) return { done: true, processed: 0, nextCursor: 2 };

   const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
   const idx = indexMap2_(headers);

   const colTagsAuto = (idx.tags_auto || 0) - 1;
   let colOut = (idx[outColName] || 0) - 1;

   if (colTagsAuto < 0) return { done: true, processed: 0, nextCursor: 2, warn: 'No tags_auto' };

   if (colOut < 0) {
      sh.insertColumnAfter(sh.getLastColumn());
      colOut = sh.getLastColumn() - 1;
      sh.getRange(1, colOut + 1).setValue(outColName);
   }

   const lr = sh.getLastRow();
   const start = Math.max(2, Number(startRow || 2));
   if (start > lr) return { done: true, processed: 0, nextCursor: 2 };

   const end = Math.min(lr, start + maxRows - 1);
   const rng = sh.getRange(start, 1, end - start + 1, sh.getLastColumn());
   const values = rng.getValues();

   let updated = 0;

   values.forEach(r => {
      const tagsAuto = String(r[colTagsAuto] || '').trim();
      if (!tagsAuto) return;

      const labels = tagsAutoToLabels_(tagsAuto, catalogs);
      const newVal = labels.join(' | ');
      const oldVal = String(r[colOut] || '').trim();

      if (newVal && newVal !== oldVal) {
         r[colOut] = newVal;
         updated++;
      }
   });

   if (updated) rng.setValues(values);

   const done = end >= lr;
   return { done, processed: values.length, updated, nextCursor: done ? 2 : end + 1 };
}

function computeVacantMatchesBatch_(vacants, persones, cfg, catalogs, startIndex, maxVacants, topN) {
   const outRows = [];
   const vacantIds = [];
   const now = toIso_(new Date());

   const start = Math.max(0, Number(startIndex || 0));
   const end = Math.min(vacants.length, start + maxVacants);

   for (let i = start; i < end; i++) {
      const v = vacants[i];
      vacantIds.push(v.id);
      const scored = [];

      persones.forEach(p => {
         const scoreObj = scoreMatch_(v, p, cfg, catalogs);
         if (scoreObj.score <= 0) return;
         scored.push({ p, ...scoreObj });
      });

      scored.sort((a,b) => b.score - a.score);
      scored.slice(0, Number(topN || 25)).forEach(s => {
         outRows.push([
            Utilities.getUuid(),
            v.id,
            s.p.id,
            s.score,
            s.reasons,
            s.kw_hits,
            s.tag_hits,
            now
         ]);
      });
   }

   const done = end >= vacants.length;
   return { done, rows: outRows, vacantIds, nextCursor: done ? 0 : end };
}

function upsertVacantSuggestionsBatch_(ss, vacantIds, rows) {
   if (!vacantIds || !vacantIds.length) return;

   const sh = ss.getSheetByName(SHEETS.VACANT_SUGGERIMENTS);
   if (!sh) return;

   const lastCol = Math.min(Math.max(1, sh.getLastColumn()), 20);
   const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(String);
   const idx = indexMap2_(headers);
   const colVac = (idx.id_vacant || 0) - 1;

   if (colVac < 0) return;

   const set = new Set(vacantIds.map(v => String(v || '').trim()));
   const lr = sh.getLastRow();

   let kept = [];
   if (lr > 1) {
      const existing = sh.getRange(2,1,lr-1,lastCol).getValues();
      kept = existing.filter(r => !set.has(String(r[colVac] || '').trim()));
   }

   if (lr > 1) sh.getRange(2,1,lr-1,lastCol).clearContent();
   if (kept.length) sh.getRange(2,1,kept.length,lastCol).setValues(kept);
   if (rows.length) sh.getRange(kept.length + 2, 1, rows.length, rows[0].length).setValues(rows);
}

function EMI_PIPELINE_scheduleNextRun_() {
   const handler = 'EMI_PIPELINE_runNext';
   const tr = ScriptApp.getProjectTriggers();
   const existing = tr.some(t => t.getHandlerFunction() === handler);
   if (existing) return;
   ScriptApp.newTrigger(handler).timeBased().after(2 * 60 * 1000).create();
}

function EMI_PIPELINE_getState_() {
   const raw = PropertiesService.getScriptProperties().getProperty(EMI_PIPELINE_STATE_KEY);
   if (raw) {
      try { return JSON.parse(raw); } catch (_) {}
   }

   const ss = getMasterSs_();
   const sh = ss.getSheetByName(SHEETS.STATE);
   if (!sh || sh.getLastRow() < 2) return null;

   const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
   const idx = indexMap2_(headers);
   const colKey = (idx.key || 0) - 1;
   const colValue = (idx.value || 0) - 1;

   if (colKey < 0 || colValue < 0) return null;

   const data = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
   for (let i = 0; i < data.length; i++) {
      const key = String(data[i][colKey] || '').trim();
      if (key !== EMI_PIPELINE_STATE_KEY) continue;
      const val = String(data[i][colValue] || '').trim();
      try { return JSON.parse(val); } catch (_) { return null; }
   }

   return null;
}

function EMI_PIPELINE_setState_(state) {
   PropertiesService.getScriptProperties().setProperty(EMI_PIPELINE_STATE_KEY, JSON.stringify(state || {}));

   const ss = getMasterSs_();
   EMI_PIPELINE_ensureStateSheet_(ss);

   const sh = ss.getSheetByName(SHEETS.STATE);
   const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
   const idx = indexMap2_(headers);
   const colKey = (idx.key || 0) - 1;
   const colValue = (idx.value || 0) - 1;
   const colUpdated = (idx.updated_at || 0) - 1;
   const colJson = (idx.json || 0) - 1;

   if (colKey < 0 || colValue < 0) return;

   const lr = sh.getLastRow();
   if (lr <= 1) {
      sh.appendRow([EMI_PIPELINE_STATE_KEY, JSON.stringify(state || {}), toIso_(new Date()), '', JSON.stringify(state || {})]);
      return;
   }

   const data = sh.getRange(2,1,lr-1,sh.getLastColumn()).getValues();
   for (let i = 0; i < data.length; i++) {
      const key = String(data[i][colKey] || '').trim();
      if (key !== EMI_PIPELINE_STATE_KEY) continue;
      const row = i + 2;
      if (colValue >= 0) sh.getRange(row, colValue + 1).setValue(JSON.stringify(state || {}));
      if (colUpdated >= 0) sh.getRange(row, colUpdated + 1).setValue(toIso_(new Date()));
      if (colJson >= 0) sh.getRange(row, colJson + 1).setValue(JSON.stringify(state || {}));
      return;
   }

   sh.appendRow([EMI_PIPELINE_STATE_KEY, JSON.stringify(state || {}), toIso_(new Date()), '', JSON.stringify(state || {})]);
}

function EMI_PIPELINE_ensureStateSheet_(ss) {
   if (!ss.getSheetByName(SHEETS.STATE)) {
      const sh = ss.insertSheet(SHEETS.STATE);
      sh.appendRow(HEADERS[SHEETS.STATE]);
   }
}

function EMI_PIPELINE_logState_(action, payload) {
   try { log_('INFO', action, 'Pipeline state update', payload || {}); } catch (_) {}
}
