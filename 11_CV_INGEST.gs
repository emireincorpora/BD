/*******************************************************
  * 11_CV_INGEST.gs — CV import (READ-ONLY SOURCE) -> MASTER
  * -----------------------------------------------------
  * - Llegeix un Spreadsheet extern de respostes CV (NO s'escriu mai al CV)
  * - Escriu al MASTER:
  *      - CV_RAW: staging amb raw_json + camps extrets/enriquits
  *      - PERSONES: crea/actualitza camps cv_* i omple alguns camps bàsics si estan buits
  *      - DOCUMENTS: enllaça adjunts detectats (si existeix taula i headers compatibles)
  *
  * IMPORTANT:
  * - NO modifiquem el Sheets del CV (ni columnes, ni camps "Processat", ni "INSCRIT").
  * - Control d'incrementalitat via PropertiesService (cursor última fila importada).
  *******************************************************/

// ========= CONFIG CV (SOURCE READ-ONLY) =========
const EMI_CV_SOURCE_SPREADSHEET_ID_DEFAULT = '1nWBcO-GOmPYD2LYQepm-ijgmd8pftnY24eGAJZOyDZM';


/**
 * ID del Spreadsheet de respostes del CV (SOURCE READ-ONLY).
 * - Per defecte, fem servir EMI_CV_SOURCE_SPREADSHEET_ID_DEFAULT
 * - Si vols canviar-ho sense tocar codi: Script Properties -> EMI_CV_SOURCE_SPREADSHEET_ID
 */
function EMI_CV_getSourceSpreadsheetId_() {
   try {
      const p = PropertiesService.getScriptProperties().getProperty('EMI_CV_SOURCE_SPREADSHEET_ID');
      return p || EMI_CV_SOURCE_SPREADSHEET_ID_DEFAULT;
   } catch (e) {
      return EMI_CV_SOURCE_SPREADSHEET_ID_DEFAULT;
   }
}

/** Setter opcional (segur) */
function EMI_CV_setSourceSpreadsheetIdNow(id) {
   id = String(id || '').trim();
   if (!id || id.length < 20) throw new Error('ID de Spreadsheet CV no vàlid');
   PropertiesService.getScriptProperties().setProperty('EMI_CV_SOURCE_SPREADSHEET_ID', id);
   EMI_CV_log_('INFO','CV_SOURCE_SET','Actualitzat EMI_CV_SOURCE_SPREADSHEET_ID', { id: id });
}
const EMI_CV_SOURCE_TAB_NAME = 'Respostes al formulari 3.0';

// Cursor incremental (última fila importada al source)
const EMI_CV_PROP_LAST_ROW = 'EMI_CV_LAST_ROW_IMPORTED';

// Per evitar timeouts
const EMI_CV_MAX_ROWS_PER_RUN = 200;

// Backfill (des de CV_RAW cap a PERSONES)
const EMI_CV_BACKFILL_MAX_ROWS = 50;

// Auto-crear/actualitzar persona
const EMI_CV_AUTOCREATE_PERSONA = true;

// ========= TAULES MASTER =========
const EMI_CV_RAW_SHEET = 'CV_RAW';

// CV_RAW headers (no destructiu: si falten columnes, s'afegeixen al final)
const EMI_CV_RAW_HEADERS = [
   'id_cv',                         // key
   'source_row',                  // fila original (1-based)
   'source_ts',                   // timestamp form
   'ingested_ts',                // quan s'ha importat
   'email',
   'nom',
   'cognoms',
   'telefon',
   'doc',
   'municipi',
   'comarca',
   'file_urls',
   'raw_json',
   'linked_id_persona',
   'status',
   'notes',

   // Enriquits (noves columnes)
   'sector_objectiu',
   'cv_competencies',
   'cv_idiomes',
   'cv_experiencia',
   'cv_formacio',
   'cv_voluntariat',
   'cv_horari',
   'cv_incorporacio',
   'carnet_conduir',
   'vehicle_propi',
   'cv_text_index',
   'applied_to_persona_ts'
];

// ================================================
//   PUBLIC API
// ================================================

/**
  * Import incremental del CV (SOURCE read-only) -> CV_RAW + merge PERSONES.
  */
function EMI_CV_pullNewRows() {
   return EMI_CV_withLock_('CV_PULL', () => {
      EMI_CV_ensureCvRawSheet_();
      EMI_CV_ensurePersonesCvColumns_();

      const srcSS = SpreadsheetApp.openById(EMI_CV_getSourceSpreadsheetId_());
      const srcSh = srcSS.getSheetByName(EMI_CV_SOURCE_TAB_NAME);
      if (!srcSh) throw new Error('No trobo la fulla del CV: ' + EMI_CV_SOURCE_TAB_NAME);

      const lastRow = srcSh.getLastRow();
      const lastCol = srcSh.getLastColumn();

      if (lastRow <= 1 || lastCol <= 1) {
         EMI_CV_log_('INFO', 'CV_EMPTY', 'No hi ha dades CV (o només headers)', { lastRow, lastCol });
         return { imported: 0, lastRow };
      }

      const headers = srcSh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
      const prev = Number(EMI_CV_getProp_(EMI_CV_PROP_LAST_ROW) || '1');
      const start = Math.max(2, prev + 1);

      if (start > lastRow) {
         EMI_CV_log_('INFO', 'CV_NO_NEW', 'No hi ha noves files per importar', { prev, lastRow });
         return { imported: 0, lastRow };
      }

      const end = Math.min(lastRow, start + EMI_CV_MAX_ROWS_PER_RUN - 1);
      const values = srcSh.getRange(start, 1, end - start + 1, lastCol).getValues();

      const shRaw = EMI_CV_getSheet_(EMI_CV_RAW_SHEET);
      const outRows = [];
      let imported = 0;

      for (let i = 0; i < values.length; i++) {
         const sourceRowIndex = start + i;
         const row = values[i];

         const obj = EMI_CV_rowToObject_(headers, row);
         const extracted = EMI_CV_extractCvFields_(obj);

         let linkedIdPersona = '';
         let status = 'NEW';
         let notes = '';
         let appliedTs = '';

         try {
            if (EMI_CV_AUTOCREATE_PERSONA) {
               linkedIdPersona = EMI_CV_findOrCreateOrUpdatePersonaFromCv_(extracted);
               status = linkedIdPersona ? 'LINKED' : 'SKIPPED';
               if (linkedIdPersona) appliedTs = EMI_CV_toIso_(new Date());
            }

            if (linkedIdPersona && extracted.fileUrls && extracted.fileUrls.length) {
               EMI_CV_linkDocumentsFromCv_(linkedIdPersona, extracted.fileUrls);
            }
         } catch (e) {
            status = 'ERROR';
            notes = 'Error linking: ' + String(e);
            EMI_CV_log_('ERROR', 'CV_LINK_FAIL', 'Error processant fila CV', {
               sourceRowIndex,
               err: String(e),
               stack: (e && e.stack) ? String(e.stack) : ''
            });
         }

         outRows.push([
            'CV-' + EMI_CV_uuid_(),
            sourceRowIndex,
            EMI_CV_parseAnyDateToIso_(extracted.sourceTs) || '',
            EMI_CV_toIso_(new Date()),
            extracted.email || '',
            extracted.nom || '',
            extracted.cognoms || '',
            extracted.telefon || '',
            extracted.doc || '',
            extracted.municipi || '',
            extracted.comarca || '',
            (extracted.fileUrls || []).join(' | '),
            EMI_CV_safeJsonStringify_(obj),
            linkedIdPersona || '',
            status,
            notes,

            extracted.sectorObjectiu || '',
            extracted.competencies || '',
            extracted.idiomesText || '',
            extracted.experienciaText || '',
            extracted.formacioText || '',
            extracted.voluntariatText || '',
            extracted.horari || '',
            extracted.incorporacio || '',
            extracted.carnet || '',
            extracted.vehicle || '',
            extracted.cvTextIndex || '',
            appliedTs
         ]);

         imported++;
      }

      if (outRows.length) {
         shRaw.getRange(shRaw.getLastRow() + 1, 1, outRows.length, EMI_CV_RAW_HEADERS.length).setValues(outRows);
      }

      EMI_CV_setProp_(EMI_CV_PROP_LAST_ROW, String(end));

      EMI_CV_log_('INFO', 'CV_IMPORTED', 'CV importats a CV_RAW', {
         imported,
         fromRow: start,
         toRow: end,
         source: EMI_CV_SOURCE_SPREADSHEET_ID,
         tab: EMI_CV_SOURCE_TAB_NAME
      });

      return { imported, fromRow: start, toRow: end };
   });
}

/**
  * Backfill: reaplica info enriquida des de CV_RAW cap a PERSONES (sense tocar el CV source).
  * Executa-la repetidament fins que et torni updated: 0.
  */
function EMI_CV_backfillPersonesFromCvRaw(maxRows) {
   return EMI_CV_withLock_('CV_BACKFILL', () => {
      EMI_CV_ensureCvRawSheet_();
      EMI_CV_ensurePersonesCvColumns_();

      const limit = Number(maxRows || EMI_CV_BACKFILL_MAX_ROWS);
      const sh = EMI_CV_getSheet_(EMI_CV_RAW_SHEET);
      const lastRow = sh.getLastRow();
      if (lastRow <= 1) return { scanned: 0, updated: 0, limit };

      const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v || '').trim());
      const hm = EMI_CV_headerMap_(header);

      const colLinked = hm['linked_id_persona'];
      const colRaw = hm['raw_json'];
      const colApplied = hm['applied_to_persona_ts'];

      if (!colLinked || !colRaw) throw new Error('CV_RAW necessita linked_id_persona i raw_json');

      const rows = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

      let updated = 0;
      let scanned = 0;

      for (let i = 0; i < rows.length && updated < limit; i++) {
         scanned++;
         const r = rows[i];
         const idPersona = String(r[colLinked - 1] || '').trim();
         const appliedTs = colApplied ? String(r[colApplied - 1] || '').trim() : '';

         if (!idPersona) continue;
         if (colApplied && appliedTs) continue;

         const raw = String(r[colRaw - 1] || '').trim();
         const obj = EMI_CV_safeParseJson_(raw);
         const extracted = EMI_CV_extractCvFields_(obj);

         EMI_CV_updatePersonaFromCv_(idPersona, extracted);

         if (colApplied) {
            sh.getRange(i + 2, colApplied).setValue(EMI_CV_toIso_(new Date()));
         }

         updated++;
      }

      EMI_CV_log_('INFO', 'CV_BACKFILL_DONE', 'Backfill CV_RAW -> PERSONES completat', { scanned, updated, limit });
      return { scanned, updated, limit };
   });
}

// ================================================
//   INTERNALS
// ================================================

function EMI_CV_ensureCvRawSheet_() {
   const ss = EMI_CV_getMasterSs_();
   let sh = ss.getSheetByName(EMI_CV_RAW_SHEET);
   if (!sh) sh = ss.insertSheet(EMI_CV_RAW_SHEET);

   if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, EMI_CV_RAW_HEADERS.length).setValues([EMI_CV_RAW_HEADERS]);
      return;
   }

   const cur = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v || '').trim());
   const existing = {};
   cur.forEach(h => { if (h) existing[h] = true; });

   const missing = EMI_CV_RAW_HEADERS.filter(h => !existing[h]);
   if (missing.length) {
      const startCol = sh.getLastColumn() + 1;
      sh.insertColumnsAfter(sh.getLastColumn(), missing.length);
      sh.getRange(1, startCol, 1, missing.length).setValues([missing]);
   }
}

/**
  * Afegeix columnes cv_* a PERSONES si no existeixen (no destructiu).
  */
function EMI_CV_ensurePersonesCvColumns_() {
   const sh = (typeof SHEETS !== 'undefined' && SHEETS.PERSONES)
      ? EMI_CV_getSheet_(SHEETS.PERSONES)
      : EMI_CV_getSheet_('PERSONES');

   const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v || '').trim());
   const existing = {};
   header.forEach(h => { if (h) existing[h] = true; });

   const needed = [
      'cv_source_ts',
      'cv_sector_objectiu',
      'cv_competencies',
      'cv_idiomes',
      'cv_experiencia',
      'cv_formacio',
      'cv_voluntariat',
      'cv_horari',
      'cv_incorporacio',
      'cv_text_index'
   ].filter(h => !existing[h]);

   if (needed.length) {
      const startCol = sh.getLastColumn() + 1;
      sh.insertColumnsAfter(sh.getLastColumn(), needed.length);
      sh.getRange(1, startCol, 1, needed.length).setValues([needed]);
      EMI_CV_log_('INFO', 'CV_PERSONES_COLS_ADDED', 'Columnes cv_* afegides a PERSONES', { added: needed });
   }
}

function EMI_CV_rowToObject_(headers, row) {
   const o = {};
   for (let i = 0; i < headers.length; i++) {
      const k = String(headers[i] || '').trim();
      if (!k) continue;
      o[k] = row[i];
   }
   return o;
}

function EMI_CV_headerMap_(headerArr) {
   const m = {};
   for (let i = 0; i < headerArr.length; i++) {
      const k = String(headerArr[i] || '').trim();
      if (k) m[k] = i + 1; // 1-based
   }
   return m;
}

function EMI_CV_safeParseJson_(s) {
   try { return JSON.parse(String(s || '{}')); } catch (_) { return {}; }
}

function EMI_CV_safeJsonStringify_(o) {
   try { return JSON.stringify(o || {}); } catch (_) { return '{}'; }
}

function EMI_CV_normalizePhone_(p) {
   const s = String(p || '').trim();
   if (!s) return '';
   return s.replace(/[^\d+]/g, '');
}

function EMI_CV_splitCognoms_(cognoms) {
   const s = String(cognoms || '').trim();
   if (!s) return { c1: '', c2: '' };
   const parts = s.split(/\s+/).filter(Boolean);
   if (parts.length === 1) return { c1: parts[0], c2: '' };
   return { c1: parts.shift(), c2: parts.join(' ') };
}

function EMI_CV_asYesNo_(v) {
   const s = String(v || '').trim().toLowerCase();
   if (!s) return '';
   if (['sí', 'si', 'sì', 'yes', 'y', 'true'].indexOf(s) >= 0) return 'SI';
   if (['no', 'n', 'false'].indexOf(s) >= 0) return 'NO';
   return String(v || '').trim();
}


/**
 * Normalitza els permisos de conduir (DGT/UE) a un format estable per a PERSONES.carnet_conduir:
 * AM, A1, A2, A, B1, B, B+E, C1, C1+E, C, C+E, D1, D1+E, D, D+E
 * - Ignora soroll (carretó, països, altres certificats)
 * - No introdueix tokens d'1 lletra
 */
function EMI_CV_normalizeDrivingLicenses_(raw, teCarnet) {
   const s0 = String(raw || '');
   const te = String(teCarnet || '').toUpperCase().trim();
   const s = s0.toUpperCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/[\n\r\t,;|]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

   const found = new Set();
   const RX = /(AM|A1|A2|A|B1|B|C1|C|D1|D)\s*(?:\+?\s*E)?/g;
   let m;
   while ((m = RX.exec(s)) !== null) {
      let tok = String(m[0] || '').replace(/\s+/g, '').toUpperCase();
      // converteix "C+E" -> "CE"
      tok = tok.replace('+', '');
      const hasE = /E$/.test(tok);
      const base = tok.replace(/E$/, '');
      const allowE = (base === 'B' || base === 'C' || base === 'C1' || base === 'D' || base === 'D1');
      if (hasE && allowE) found.add(base + '+E');
      else found.add(base);
   }

   if (!found.size) {
      // Si sabem que té carnet però no concreta quin, ho guardem com "SI"
      if (te === 'SI') return 'SI';
      if (te === 'NO') return 'NO';
      return '';
   }

   const ORDER = ['AM','A1','A2','A','B1','B','B+E','C1','C1+E','C','C+E','D1','D1+E','D','D+E'];
   const arr = Array.from(found);
   arr.sort((a,b) => {
      const ia = ORDER.indexOf(a); const ib = ORDER.indexOf(b);
      if (ia !== -1 && ib !== -1) return ia - ib;
      if (ia !== -1) return -1;
      if (ib !== -1) return 1;
      return a.localeCompare(b);
   });
   return arr.join(', ');
}

/**
  * Extractor enriquit basat en heurístiques (robust davant canvis al formulari).
  * Treu: experiència, formació, idiomes, voluntariat, competències, horari/incorporació, carnets/vehicle.
  */
function EMI_CV_extractCvFields_(obj) {
   const keys = Object.keys(obj);

   function pick_(regexList) {
      for (let r = 0; r < regexList.length; r++) {
         const rx = regexList[r];
         for (let i = 0; i < keys.length; i++) {
            const k = keys[i];
            if (rx.test(String(k).toLowerCase())) return obj[k];
         }
      }
      return '';
   }

   function pickNumbered_(patternWithN, n) {
      const rx = new RegExp(patternWithN.replace('{n}', String(n)), 'i');
      for (let i = 0; i < keys.length; i++) {
         const k = keys[i];
         if (rx.test(String(k))) return obj[k];
      }
      return '';
   }

   const sourceTs = pick_([/marca de temps/, /marca temporal/, /timestamp/]);
   const email = String(pick_([/correu electr[oò]nic/, /\bemail\b/, /correu/]) || '').trim().toLowerCase();
   const telefon = String(pick_([/tel[eè]fon/, /m[oò]bil/, /mobil/]) || '').trim();
   const doc = String(pick_([/\bdni\b/, /\bnie\b/, /document/]) || '').trim();
   const municipi = String(pick_([/poble|ciutat.*resid[eè]ncia/, /municipi/, /poblaci[oó]/]) || '').trim();
   const comarca = String(pick_([/comarca/]) || '').trim();

   // Nom i cognoms
   let nom = '';
   let cognoms = '';
   const fullName = String(pick_([/nom i cognoms/, /nom complet/, /^nom$/]) || '').trim();
   if (fullName) {
      const parts = fullName.split(/\s+/);
      nom = parts.shift() || '';
      cognoms = parts.join(' ');
   }

   // Sector objectiu (el guardem, però NO el posem a ambits_interes)
   const sectorObjectiu = String(pick_([/sector laboral.*buscant/, /sector laboral/]) || '').trim();

   // Competències
   const competencies = String(pick_([/compet[eè]ncies.*desenvolupades/, /compet[eè]ncies/]) || '').trim();

   // Horari / incorporació
   const horari = String(pick_([/franja hor[aà]ria/, /horari/]) || '').trim();
   const incorporacio = String(pick_([/incorporar/, /quan et podries incorporar/]) || '').trim();

   // Carnets / vehicle
   const teCarnet = EMI_CV_asYesNo_(pick_([/tens carnet de conduir/]));
   const carnetsRaw = String(pick_([/quins carnets tens/]) || '').trim();
   const carnet = EMI_CV_normalizeDrivingLicenses_(carnetsRaw, teCarnet);
   const vehicle = EMI_CV_asYesNo_(pick_([/tens vehicle propi/]));

   // Idiomes (fins a 5)
   const idiomesParts = [];
   for (let n = 1; n <= 5; n++) {
      const idi = String(pickNumbered_('Idioma\\s*{n}\\b', n) || '').trim();
      const niv = String(pickNumbered_('Nivell\\s*Idioma\\s*{n}\\b', n) || '').trim();
      if (idi) idiomesParts.push(niv ? (idi + ' (' + niv + ')') : idi);
   }
   const idiomesText = idiomesParts.join(' · ');

   // Voluntariat (fins a 3)
   const volParts = [];
   for (let n = 1; n <= 3; n++) {
      const nomV = String(pickNumbered_('Nom\\s*Voluntariat\\s*{n}\\b', n) || '').trim();
      const expV = String(pickNumbered_('Breu\\s*Explicaci[oó].*Voluntariat\\s*{n}\\b', n) || '').trim();
      const anyV = String(pickNumbered_('Any\\s*Voluntariat\\s*{n}\\b', n) || '').trim();
      if (nomV || expV) {
         let line = nomV || 'Voluntariat';
         if (anyV) line += ' (' + anyV + ')';
         if (expV) line += ': ' + expV;
         volParts.push(line);
      }
   }
   const voluntariatText = volParts.join(' | ');

   // Formació (fins a 10)
   const formParts = [];
   for (let n = 1; n <= 10; n++) {
      const tit = String(pickNumbered_('Nom\\s*del\\s*t[ií]tol\\s*o\\s*formaci[oó]\\s*{n}\\b', n) || '').trim();
      const any = String(pickNumbered_('Any\\s*d\\’?obtenci[oó].*T[ií]tol\\s*{n}\\b', n) || '').trim();
      if (tit) formParts.push(any ? (tit + ' (' + any + ')') : tit);
   }
   const formacioText = formParts.join(' · ');

   // Experiència (fins a 20)
   const expParts = [];
   for (let n = 1; n <= 20; n++) {
      const emp = String(pickNumbered_("Nom\\s*de\\s*l\\'?Empresa\\s*{n}\\b", n) || '').trim();
      const carrec = String(pickNumbered_("C[aà]rrec\\s*dins\\s*l\\'?empresa\\s*{n}\\b", n) || '').trim();
      const func = String(pickNumbered_('Descriu\\s*3\\s*o\\s*4\\s*funcions.*empresa\\s*{n}\\b', n) || '').trim();
      if (emp || carrec || func) {
         let line = emp || 'Empresa';
         if (carrec) line += ' — ' + carrec;
         if (func) line += ': ' + func;
         expParts.push(line);
      }
   }
   const experienciaText = expParts.join(' | ');

   // Adjuntos: detecta headers amb "Adjunta/Fitxer/Arxiu/Upload"
   const fileKeys = keys.filter(k => /adjunta|adjunt|fitxer|arxiu|puja|upload|file/i.test(String(k)));
   const fileUrls = [];
   fileKeys.forEach(k => {
      const v = obj[k];
      if (v === null || v === undefined) return;
      const s = String(v).trim();
      if (!s) return;
      s.split(/\s*,\s*|\s*\n\s*/).forEach(x => {
         const t = String(x || '').trim();
         if (t) fileUrls.push(t);
      });
   });

   const normFiles = fileUrls.map(f => {
      if (/^https?:\/\//i.test(f)) return f;
      if (/^[a-zA-Z0-9_-]{20,}$/.test(f)) return 'https://drive.google.com/open?id=' + f;
      return f;
   });

   // Text index base (per cerca/match via reindex més endavant)
   const cvTextIndex = [
      sectorObjectiu,
      competencies,
      idiomesText,
      voluntariatText,
      formacioText,
      experienciaText,
      horari,
      incorporacio,
      carnet,
      vehicle,
      municipi
   ].filter(Boolean).join(' · ').slice(0, 5000);

   return {
      sourceTs: sourceTs ? String(sourceTs) : '',
      email,
      nom,
      cognoms,
      telefon,
      doc,
      municipi,
      comarca,

      sectorObjectiu,
      competencies,
      idiomesText,
      voluntariatText,
      formacioText,
      experienciaText,
      horari,
      incorporacio,
      carnet,
      vehicle,
      cvTextIndex,

      fileUrls: normFiles
   };
}

/**
  * Dedupe: email -> tel -> nom+cognoms
  * Si existeix: actualitza cv_* + omple bàsics si buit
  * Si no existeix: crea persona mínima + cv_*
  */
function EMI_CV_findOrCreateOrUpdatePersonaFromCv_(extracted) {
   const sh = (typeof SHEETS !== 'undefined' && SHEETS.PERSONES)
      ? EMI_CV_getSheet_(SHEETS.PERSONES)
      : EMI_CV_getSheet_('PERSONES');

   const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v || '').trim());
   const hm = EMI_CV_headerMap_(header);

   const colId = hm['id_persona'];
   if (!colId) throw new Error('PERSONES necessita id_persona');

   const email = String(extracted.email || '').trim().toLowerCase();
   const telNorm = EMI_CV_normalizePhone_(extracted.telefon);
   const nomNorm = (String(extracted.nom || '') + ' ' + String(extracted.cognoms || '')).replace(/\s+/g, ' ').trim().toLowerCase();

   const lastRow = sh.getLastRow();
   if (lastRow > 1) {
      const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

      const colEmail = hm['email'] || 0;
      const colTel = hm['telefon_mobil'] || 0;
      const colNom = hm['nom'] || 0;
      const colCog1 = hm['cognom1'] || 0;
      const colCog2 = hm['cognom2'] || 0;

      function rowEmail(i) { return colEmail ? String(data[i][colEmail - 1] || '').trim().toLowerCase() : ''; }
      function rowTel(i) { return colTel ? EMI_CV_normalizePhone_(data[i][colTel - 1]) : ''; }
      function rowFullName(i) {
         const n = colNom ? String(data[i][colNom - 1] || '').trim() : '';
         const c1 = colCog1 ? String(data[i][colCog1 - 1] || '').trim() : '';
         const c2 = colCog2 ? String(data[i][colCog2 - 1] || '').trim() : '';
         return (n + ' ' + c1 + ' ' + c2).replace(/\s+/g, ' ').trim().toLowerCase();
      }

      // 1) email
      if (email && colEmail) {
         for (let i = 0; i < data.length; i++) {
            if (rowEmail(i) === email) {
               const id = String(data[i][colId - 1] || '').trim();
               if (id) EMI_CV_updatePersonaFromCv_(id, extracted);
               return id;
            }
         }
      }

      // 2) tel
      if (telNorm && colTel) {
         for (let i = 0; i < data.length; i++) {
            if (rowTel(i) === telNorm) {
               const id = String(data[i][colId - 1] || '').trim();
               if (id) EMI_CV_updatePersonaFromCv_(id, extracted);
               return id;
            }
         }
      }

      // 3) nom+cognoms
      if (nomNorm && (colNom && colCog1)) {
         for (let i = 0; i < data.length; i++) {
            if (rowFullName(i) === nomNorm) {
               const id = String(data[i][colId - 1] || '').trim();
               if (id) EMI_CV_updatePersonaFromCv_(id, extracted);
               return id;
            }
         }
      }
   }

   // Crear nova persona
   const now = EMI_CV_toIso_(new Date());
   const idPersona = 'P-' + EMI_CV_uuid_();
   const cogs = EMI_CV_splitCognoms_(extracted.cognoms);

   const rowObj = {
      id_persona: idPersona,
      emi_id_persona: '',
      data_alta: now,
      origen_alta: 'CV',
      estat_fitxa: 'NOVA',

      nom: extracted.nom || '',
      cognom1: cogs.c1 || '',
      cognom2: cogs.c2 || '',

      telefon_mobil: extracted.telefon || '',
      email: extracted.email || '',
      doc_numero: extracted.doc || '',
      municipi: extracted.municipi || '',
      comarca: extracted.comarca || '',

      // Omplir camps bàsics (light) sense trepitjar res futur:
      disponibilitat_immediata: extracted.incorporacio || '',
      jornada_preferida: extracted.horari || '',
      carnet_conduir: extracted.carnet || '',
      vehicle_propi: extracted.vehicle || '',

      // IMPORTANT: NO toquem ambits_interes des del CV
      ambits_interes: '',

      // cv_* consultables
      cv_source_ts: EMI_CV_parseAnyDateToIso_(extracted.sourceTs) || '',
      cv_sector_objectiu: extracted.sectorObjectiu || '',
      cv_competencies: extracted.competencies || '',
      cv_idiomes: extracted.idiomesText || '',
      cv_experiencia: extracted.experienciaText || '',
      cv_formacio: extracted.formacioText || '',
      cv_voluntariat: extracted.voluntariatText || '',
      cv_horari: extracted.horari || '',
      cv_incorporacio: extracted.incorporacio || '',
      cv_text_index: extracted.cvTextIndex || '',

      tags_persona: 'CV',
      observacions: 'Creat automàticament des de CV',

      created_by_email: 'SYSTEM_CV',
      updated_by_email: 'SYSTEM_CV',
      data_ultima_actualitzacio: now
   };

   const out = new Array(header.length).fill('');
   for (let c = 0; c < header.length; c++) {
      const k = header[c];
      if (rowObj.hasOwnProperty(k)) out[c] = rowObj[k];
   }

   sh.appendRow(out);
   EMI_CV_log_('INFO', 'CV_PERSONA_CREATED', 'Persona creada des de CV', { id_persona: idPersona, email: extracted.email });
   return idPersona;
}

/**
  * Actualitza persona existent:
  * - sempre actualitza cv_* + updated_by_email + data_ultima_actualitzacio
  * - omple només si buit: tel/doc/municipi/comarca/carnet/vehicle/jornada/incorporació
  * - IMPORTANT: NO toca ambits_interes
  */
function EMI_CV_updatePersonaFromCv_(idPersona, extracted) {
   if (!idPersona) return;

   const sh = (typeof SHEETS !== 'undefined' && SHEETS.PERSONES)
      ? EMI_CV_getSheet_(SHEETS.PERSONES)
      : EMI_CV_getSheet_('PERSONES');

   const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v || '').trim());
   const hm = EMI_CV_headerMap_(header);

   const colId = hm['id_persona'];
   if (!colId) throw new Error('PERSONES necessita id_persona');

   const lastRow = sh.getLastRow();
   if (lastRow <= 1) return;

   const ids = sh.getRange(2, colId, lastRow - 1, 1).getValues();
   let rowIndex = -1;
   for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim() === String(idPersona).trim()) {
         rowIndex = i + 2;
         break;
      }
   }
   if (rowIndex < 0) return;

   const row = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];

   function getCell_(colName) {
      const c = hm[colName];
      if (!c) return '';
      return String(row[c - 1] || '').trim();
   }
   function setCell_(colName, value) {
      const c = hm[colName];
      if (!c) return;
      row[c - 1] = value;
   }
   function setIfEmpty_(colName, value) {
      const v = String(value || '').trim();
      if (!v) return;
      const cur = getCell_(colName);
      if (!cur) setCell_(colName, v);
   }

   const now = EMI_CV_toIso_(new Date());

   // Omple si buit (no trepitjar)
   setIfEmpty_('telefon_mobil', extracted.telefon || '');
   setIfEmpty_('doc_numero', extracted.doc || '');
   setIfEmpty_('municipi', extracted.municipi || '');
   setIfEmpty_('comarca', extracted.comarca || '');
   setIfEmpty_('carnet_conduir', extracted.carnet || '');
   setIfEmpty_('vehicle_propi', extracted.vehicle || '');
   setIfEmpty_('jornada_preferida', extracted.horari || '');
   setIfEmpty_('disponibilitat_immediata', extracted.incorporacio || '');

   // IMPORTANT: NO setIfEmpty_('ambits_interes', ...)

   // Sempre actualitza cv_*
   setCell_('cv_source_ts', EMI_CV_parseAnyDateToIso_(extracted.sourceTs) || '');
   setCell_('cv_sector_objectiu', extracted.sectorObjectiu || '');
   setCell_('cv_competencies', extracted.competencies || '');
   setCell_('cv_idiomes', extracted.idiomesText || '');
   setCell_('cv_experiencia', extracted.experienciaText || '');
   setCell_('cv_formacio', extracted.formacioText || '');
   setCell_('cv_voluntariat', extracted.voluntariatText || '');
   setCell_('cv_horari', extracted.horari || '');
   setCell_('cv_incorporacio', extracted.incorporacio || '');
   setCell_('cv_text_index', extracted.cvTextIndex || '');

   // Audit
   setCell_('updated_by_email', 'SYSTEM_CV');
   setCell_('data_ultima_actualitzacio', now);

   sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
   EMI_CV_log_('INFO', 'CV_PERSONA_UPDATED', 'Persona actualitzada des de CV', { id_persona: idPersona });
}

/**
  * Enllaça adjunts detectats a DOCUMENTS (si la taula existeix i té headers compatibles).
  * Si la teva taula DOCUMENTS té altres noms de columnes, no peta: simplement omple el que troba.
  */
function EMI_CV_linkDocumentsFromCv_(idPersona, fileUrls) {
   if (!idPersona || !fileUrls || !fileUrls.length) return;

   // Si no existeix SHEETS.DOCUMENTS, provem 'DOCUMENTS'
   let sh;
   try {
      sh = (typeof SHEETS !== 'undefined' && SHEETS.DOCUMENTS)
         ? EMI_CV_getSheet_(SHEETS.DOCUMENTS)
         : EMI_CV_getSheet_('DOCUMENTS');
   } catch (e) {
      // si no hi ha taula, no fem res
      EMI_CV_log_('WARN', 'CV_DOCS_SKIP', 'No hi ha taula DOCUMENTS al MASTER', { id_persona: idPersona });
      return;
   }

   const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v || '').trim());
   const now = EMI_CV_toIso_(new Date());

   const rows = [];

   fileUrls.forEach(u => {
      const url = String(u || '').trim();
      if (!url) return;

      const rowObj = {
         id_document: 'DOC-' + EMI_CV_uuid_(),
         entity_type: 'PERSONA',
         entity_id: idPersona,
         ts: now,
         doc_type: 'CV',
         file_url: url,
         notes: 'Importat automàticament del Form CV',
         created_by_email: 'SYSTEM_CV'
      };

      const out = new Array(header.length).fill('');
      for (let c = 0; c < header.length; c++) {
         const k = header[c];
         if (rowObj.hasOwnProperty(k)) out[c] = rowObj[k];
      }
      rows.push(out);
   });

   if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, header.length).setValues(rows);
      EMI_CV_log_('INFO', 'CV_DOCS_LINKED', 'Documents CV vinculats', { id_persona: idPersona, docs: rows.length });
   }
}

// ================================================
//   SMALL UTIL WRAPPERS (per no dependre del projecte)
// ================================================

function EMI_CV_getMasterSs_() {
   if (typeof getMasterSs_ === 'function') return getMasterSs_();
   return SpreadsheetApp.getActiveSpreadsheet();
}

function EMI_CV_getSheet_(name) {
   if (typeof getSheet_ === 'function') return getSheet_(name);
   const ss = EMI_CV_getMasterSs_();
   const sh = ss.getSheetByName(name);
   if (!sh) throw new Error('No trobo la pestanya: ' + name);
   return sh;
}

function EMI_CV_uuid_() {
   if (typeof uuid_ === 'function') return uuid_();
   return Utilities.getUuid();
}

function EMI_CV_toIso_(d) {
   if (typeof toIso_ === 'function') return toIso_(d);
   return (d instanceof Date ? d : new Date(d)).toISOString();
}

// Parse robust per evitar confusions dd/mm vs mm/dd i formats d'Excel
function EMI_CV_parseAnyDateToIso_(v) {
   if (v === null || v === undefined || v === '') return '';
   try {
      if (v instanceof Date) return EMI_CV_toIso_(v);
   } catch (e) {}

   // Numeric (serial) -> intenta Date()
   if (typeof v === 'number') {
      var dNum = new Date(v);
      if (!isNaN(dNum.getTime())) return EMI_CV_toIso_(dNum);
   }

   var s = String(v || '').trim();
   if (!s) return '';

   // ISO o similar
   var dIso = new Date(s);
   if (!isNaN(dIso.getTime()) && /\d{4}/.test(s)) return EMI_CV_toIso_(dIso);

   // dd/mm/yyyy [hh:mm[:ss]]
   var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
   if (m) {
      var dd = parseInt(m[1], 10);
      var mm = parseInt(m[2], 10);
      var yy = parseInt(m[3], 10);
      if (yy < 100) yy = 2000 + yy;

      var hh = m[4] ? parseInt(m[4], 10) : 0;
      var mi = m[5] ? parseInt(m[5], 10) : 0;
      var ss = m[6] ? parseInt(m[6], 10) : 0;

      var d = new Date(yy, mm - 1, dd, hh, mi, ss);
      if (!isNaN(d.getTime())) return EMI_CV_toIso_(d);
   }

   // fallback
   var d2 = new Date(s);
   if (!isNaN(d2.getTime())) return EMI_CV_toIso_(d2);

   return '';
}

function EMI_CV_withLock_(name, fn) {
   if (typeof withLock_ === 'function') return withLock_(name, fn);
   return fn();
}

function EMI_CV_log_(level, code, msg, obj) {
   if (typeof log_ === 'function') return log_(level, code, msg, obj || {});
   Logger.log(level + ' ' + code + ' ' + msg + ' ' + JSON.stringify(obj || {}));
}

function EMI_CV_getProp_(k) {
   if (typeof getProp_ === 'function') return getProp_(k);
   return PropertiesService.getScriptProperties().getProperty(k);
}

function EMI_CV_setProp_(k, v) {
   if (typeof setProp_ === 'function') return setProp_(k, v);
   PropertiesService.getScriptProperties().setProperty(k, v);
}


/**
 * REPAIR: reescriu CV_RAW.source_ts a partir de raw_json["Marca temporal"/"Marca de temps"/"timestamp"].
 * Útil si tens source_ts desquadrats (Excel/format).
 */
function EMI_CV_repairCvRawSourceTsFromRawJsonNow(maxRows) {
   maxRows = maxRows || 2000;

   return EMI_CV_withLock_('CV_REPAIR_SOURCE_TS', function() {
      var ss = getMasterSs_();
      var sh = ss.getSheetByName(EMI_CV_RAW_SHEET);
      if (!sh || sh.getLastRow() < 2) {
         EMI_CV_log_('INFO','CV_REPAIR_EMPTY','No hi ha CV_RAW per reparar', {});
         return { repaired: 0, scanned: 0 };
      }

      var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(x){ return String(x||'').trim(); });
      var idx = {};
      headers.forEach(function(h,i){ idx[h]=i; });

      var cSource = idx['source_ts'];
      var cRaw = idx['raw_json'];
      if (cSource === undefined || cRaw === undefined) {
         EMI_CV_log_('ERROR','CV_REPAIR_HEADERS','Falten columnes source_ts/raw_json a CV_RAW', { headers: headers });
         return { repaired: 0, scanned: 0, error: 'missing_headers' };
      }

      var lr = sh.getLastRow();
      var n = Math.min(maxRows, lr - 1);
      var rng = sh.getRange(2,1,n,sh.getLastColumn());
      var vals = rng.getValues();

      var repaired = 0;

      function findTsInObj_(obj) {
         if (!obj) return '';
         var keys = Object.keys(obj);
         for (var i=0;i<keys.length;i++){
            var k = String(keys[i]||'').toLowerCase();
            if (/marca\s*(temporal|de\s*temps)|timestamp/.test(k)) return obj[keys[i]];
         }
         return '';
      }

      for (var r=0;r<vals.length;r++){
         var rawStr = String(vals[r][cRaw] || '').trim();
         if (!rawStr) continue;
         var obj;
         try { obj = JSON.parse(rawStr); } catch(e) { continue; }

         var ts = findTsInObj_(obj);
         var iso = EMI_CV_parseAnyDateToIso_(ts);
         if (!iso) continue;

         var prev = String(vals[r][cSource] || '').trim();
         if (prev !== iso) {
            vals[r][cSource] = iso;
            repaired++;
         }
      }

      if (repaired) {
         rng.setValues(vals);
      }

      EMI_CV_log_('INFO','CV_REPAIR_DONE','Reparació CV_RAW.source_ts completada', { scanned: n, repaired: repaired });
      return { scanned: n, repaired: repaired };
   });
}
