/*******************************************************
  * 10_BRAIN_INDEX_MATCH.gs — CERVELL
  * - text_index (per buscar millor a AppSheet)
  * - tags_auto (per fer matching i filtres)
  * - matching VACANTS ↔ PERSONES (suggeriments)
  *
  * IMPORTANT:
  *   - No depèn d'AppSheet Bots.
  *   - El CRON del 06_TRIGGERS_IDS fa de “bot”.
  *******************************************************/

function EMI_brainBootstrap() {
   withLock_('BRAIN_BOOTSTRAP', () => {
      const ss = getMasterSs_();

      // 1) assegura columnes (no trenca res, només afegeix si falten)
      ensureBrainColumns_(ss);

      // 2) assegura sheet de suggeriments
      ensureVacantSuggerimentsSheet_(ss);

      // 3) reindex complet + matching complet (1 cop)
      EMI_reindexAllNow();
      EMI_recomputeVacantMatchesNow(25);

      log_('INFO', 'BRAIN_BOOTSTRAP_OK', 'Cervell inicialitzat', {});
      uiAlert_(
         'Cervell inicialitzat ✅\n\n' +
         '- text_index + tags_auto creats\n' +
         '- matching vacants↔persones generat\n\n' +
         "Ja pots buscar a AppSheet amb molta més potència."
      );
   });
}

function ensureBrainColumns_(ss) {
   // PERSONES
   ensureColumns_(ss, SHEETS.PERSONES, ['tags_auto','text_index','tags_persona_auto']);
   // EMPRESES
   ensureColumns_(ss, SHEETS.EMPRESES, ['tags_auto','text_index']); // (si algun dia vols, hi podem afegir tags_empresa_auto)
   // VACANTS
   ensureColumns_(ss, SHEETS.VACANTS, ['tags_auto','text_index','tags_vacant_auto']);

   SpreadsheetApp.flush();
}

function EMI_lastHeaderCol_(_headers) {
   // Retorna l'últim index (1-based) amb header no buit.
   if (!_headers || !_headers.length) return 0;
   for (let i = _headers.length; i >= 1; i--) {
      const v = _headers[i - 1];
      if (v !== null && v !== undefined && String(v).trim() !== '') return i;
   }
   return 0;
}

function ensureColumns_(ss, sheetName, cols) {
   const sh = ss.getSheetByName(sheetName);
   if (!sh) return;

   const lastCol = Math.min(Math.max(1, sh.getLastColumn()), EMI_BRAIN_MAX_HEADER_COLS);
   const headersRaw = sh.getRange(1, 1, 1, lastCol).getValues()[0];
   const headers = headersRaw.map(h => String(h || '').trim());

   // IMPORTANT: no fem servir getLastColumn() per afegir headers, perquè hi pot haver "columnes fantasma".
   let lastHeaderCol = EMI_lastHeaderCol_(headers);
   if (!lastHeaderCol) lastHeaderCol = 1;

   const set = new Set(headers.map(h => String(h || '').trim()).filter(Boolean));
   let added = 0;

   cols.forEach(c => {
      if (set.has(c)) return;
      sh.insertColumnAfter(lastHeaderCol);
      lastHeaderCol = lastHeaderCol + 1;
      sh.getRange(1, lastHeaderCol).setValue(c);
      set.add(c);
      added++;
   });

   if (added) log_('INFO','BRAIN_COLUMNS_ADDED','Afegides columnes cervell', { sheetName, added, cols });
}

function ensureVacantSuggerimentsSheet_(ss) {
   const name = 'VACANT_SUGGERIMENTS';
   let sh = ss.getSheetByName(name);
   if (!sh) sh = ss.insertSheet(name);

   const headers = ['id_sug','id_vacant','id_persona','score','reasons','kw_hits','tag_hits','created_at'];
   if (sh.getLastRow() === 0) {
      sh.appendRow(headers);
      return;
   }

   const curCols = Math.min(Math.max(1, sh.getLastColumn()), 50);
   const current = sh.getRange(1,1,1,curCols).getValues()[0].map(String);
   if (current.join('').trim() === '') {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
   } else {
      const set = new Set(current.map(s => String(s||'').trim()));
      const toAdd = headers.filter(h => !set.has(h));
      if (toAdd.length) {
         sh.insertColumnsAfter(sh.getLastColumn(), toAdd.length);
         sh.getRange(1, sh.getLastColumn() - toAdd.length + 1, 1, toAdd.length).setValues([toAdd]);
      }
   }
}

function EMI_reindexAllNow() {
   withLock_('REINDEX_ALL', () => {
      const ss = getMasterSs_();
      ensureBrainColumns_(ss);

      const dict = buildBrainDictionaries_();

      reindexSheet_(ss, SHEETS.PERSONES, dict, buildPersonaBaseText_);
      reindexSheet_(ss, SHEETS.EMPRESES, dict, buildEmpresaBaseText_);
      reindexSheet_(ss, SHEETS.VACANTS,   dict, buildVacantBaseText_);

      SpreadsheetApp.flush();
      log_('INFO','REINDEX_ALL_OK','Reindex complet OK', {});
   });

   // Omplim també els camps visibles (tags_persona_auto / tags_vacant_auto) si tens el script 17 instal·lat
   if (typeof EMI_refreshAutoTagsLabelsNow === 'function') {
      try { EMI_refreshAutoTagsLabelsNow(); } catch (err) { /* no trenquem res */ }
   }
}

function EMI_reindexRecent() {
   EMI_reindexAllNow();
}

function reindexSheet_(ss, sheetName, dict, baseTextFn) {
   const sh = ss.getSheetByName(sheetName);
   if (!sh) return;

   const lr = sh.getLastRow();
   if (lr < 2) return;

   const lastCol = Math.min(Math.max(1, sh.getLastColumn()), EMI_BRAIN_MAX_HEADER_COLS);
   const headersRaw = sh.getRange(1, 1, 1, lastCol).getValues()[0];
   const headers = headersRaw.map(h => String(h || '').trim());
   const idx = indexMap2_(headers);

   const tagsAutoC = idx['tags_auto'];
   const textIndexC = idx['text_index'];

   if (!tagsAutoC || !textIndexC) {
      log_('WARN','REINDEX_MISSING_COLS','Falten columnes tags_auto/text_index', { sheetName });
      return;
   }

   // Limitem el rang a les columnes amb header (evita llegir columnes fantasma i excedir el temps màxim)
   const lastHeaderCol = Math.max(
      EMI_lastHeaderCol_(headers),
      Number(tagsAutoC || 0),
      Number(textIndexC || 0)
   );

   const rng = sh.getRange(2, 1, lr - 1, lastHeaderCol);
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

   log_('INFO','REINDEX_SHEET_DONE','Reindex sheet', { sheetName, changed, lastHeaderCol });
}

function buildPersonaBaseText_(row, idx) {
   // Agafa camps “importants” per indexació (ajusta quan vulguis, no trenca res)
   const parts = [
      v_(row, idx, 'nom'),
      v_(row, idx, 'cognom1'),
      v_(row, idx, 'cognom2'),
      v_(row, idx, 'municipi'),
      v_(row, idx, 'comarca'),
      v_(row, idx, 'situacio_administrativa'),
      v_(row, idx, 'situacio_laboral_actual'),
      v_(row, idx, 'nivell_formacio'),
      v_(row, idx, 'ambits_interes'),
      v_(row, idx, 'programes_interes'),
      v_(row, idx, 'tags_persona'),
      v_(row, idx, 'tags_persona_auto'),

      // CV (consultable) — així el matching aprofita el text del CV
      v_(row, idx, 'cv_sector_objectiu'),
      v_(row, idx, 'cv_competencies'),
      v_(row, idx, 'cv_idiomes'),
      v_(row, idx, 'cv_experiencia'),
      v_(row, idx, 'cv_formacio'),
      v_(row, idx, 'cv_voluntariat'),
      v_(row, idx, 'cv_horari'),
      v_(row, idx, 'cv_incorporacio'),
      v_(row, idx, 'cv_text_index'),

      v_(row, idx, 'observacions')
   ];
   return parts.filter(Boolean).join(' ');
}

function buildEmpresaBaseText_(row, idx) {
   const parts = [
      v_(row, idx, 'nom_empresa'),
      v_(row, idx, 'cif'),
      v_(row, idx, 'sector_principal'),
      v_(row, idx, 'altres_sectors'),
      v_(row, idx, 'municipi'),
      v_(row, idx, 'comarca'),
      v_(row, idx, 'busques_a_emi'),
      v_(row, idx, 'tags_empresa'),
      v_(row, idx, 'observacions')
   ];
   return parts.filter(Boolean).join(' ');
}

function buildVacantBaseText_(row, idx) {
   const parts = [
      v_(row, idx, 'titol_vacant'),
      v_(row, idx, 'descripcio'),
      v_(row, idx, 'requisits'),
      v_(row, idx, 'sector'),
      v_(row, idx, 'municipi'),
      v_(row, idx, 'comarca'),
      v_(row, idx, 'tags_vacant'),
      v_(row, idx, 'tags_vacant_auto')
   ];
   return parts.filter(Boolean).join(' ');
}

function v_(row, idx, col) {
   const c = idx[col];
   if (!c) return '';
   return String(row[c - 1] || '').trim();
}

// ====== DICTIONARIES ======

function buildBrainDictionaries_() {
   const ss = getMasterSs_();
   const cat = readCatalogs_(ss);
   const syn = readSinonims_(ss);
   const forms = readFormacions_(ss);

   // reverse token -> [{tipus,codi,valor,w}]
   const reverse = {};

   function addTok_(tok, entry) {
      const t = normalizeToken_(tok);
      if (!t) return;
      if (!reverse[t]) reverse[t] = [];
      reverse[t].push(entry);
   }

   // Només TAG detallat (per scoring); evitem inflar amb altres tipus
   Object.keys(cat).forEach(k => {
      const entry = cat[k];
      if (!entry || entry.tipus !== 'TAG') return;

      // Tokens del nom oficial
      tokenize_(entry.valor).forEach(t => addTok_(t, { tipus:'TAG', codi: entry.codi, valor: entry.valor, w: 1 }));
   });

   // SINÒNIMS (amb pes)
   syn.forEach(s => {
      if (!s || String(s.tipus || '').trim() !== 'TAG') return;
      const c = String(s.codi_ref || '').trim();
      const key = 'TAG||' + c;
      const catEntry = cat[key];
      const label = catEntry ? String(catEntry.valor || '').trim() : '';
      const w = Math.max(0.1, Math.min(5, Number(s.pes || 1)));

      const base = { tipus:'TAG', codi: c, valor: label, w: w };

      addTok_(s.sinonim, base);
      tokenize_(s.sinonim).forEach(t => addTok_(t, base));
   });

   // FORMATIONS_CATALOG: NO entra al scoring (es pot usar a cerca, però no com TAGS)
   // Per evitar soroll, NO l’afegim al reverse del matching.

   return { reverse: reverse, catalogs: cat };
}

// ====== COMPUTE INDEX + TAGS ======

function EMI_brainExtractSpecialTokens_(rawText) {
   const out = new Set();
   const s0 = String(rawText || '');
   if (!s0) return [];

   // Normalize: majúscules + sense accents
   const s = s0.toUpperCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '');

   // Detecta permisos/carnets Espanya (UE): AM, A1, A2, A, B1, B, BE, C1, C1E, C, CE, D1, D1E, D, DE
   // IMPORTANT: no acceptem 'B' sol com a token; generem 'carnet_b', 'carnet_ce', etc.
   const CODE_RX = /(AM|A1|A2|A|B1|B|C1|C|D1|D)\s*(?:\+?\s*E)?/g;

   function canon_(base, hasE) {
      base = String(base || '').toUpperCase();
      const allowE = (base === 'B' || base === 'C' || base === 'C1' || base === 'D' || base === 'D1');
      if (hasE && allowE) return base + 'E';
      return base;
   }

   // Casos típics: "carnet B", "permiso C+E", "permis de conduir D1+E", "carnets: B, C, C+E"
   const HEAD_RX = /(CARNET(?:S)?|PERMIS(?:OS)?|PERMISO(?:S)?|LICENCIA(?:S)?)(?:\s+DE\s+(?:CONDUIR|CONDUCIR))?[\s:;-]*/g;

   // 1) Captura després d'un "head"
   let m;
   while ((m = HEAD_RX.exec(s)) !== null) {
      const start = m.index + m[0].length;
      const chunk = s.slice(start, Math.min(s.length, start + 80)); // finestra curta
      let c;
      while ((c = CODE_RX.exec(chunk)) !== null) {
         const raw = String(c[0] || '').replace(/\s+/g, '');
         const base = raw.replace(/\+E$/,'').replace(/E$/,'');
         const hasE = /E$/.test(raw) || /\+E$/.test(raw);
         const canon = canon_(base, hasE);
         if (canon) out.add('carnet_' + canon.toLowerCase());
      }
   }

   // 2) Captura directa de patrons "CARNET B", "PERMISO C+E" sense llista
   const DIRECT_RX = /(CARNET|PERMIS|PERMISO|LICENCIA)(?:\s+DE\s+(?:CONDUIR|CONDUCIR))?[\s:,-]*((AM|A1|A2|A|B1|B|C1|C|D1|D)(?:\s*\+\s*E)?)/g;
   while ((m = DIRECT_RX.exec(s)) !== null) {
      const raw = String(m[2] || '').replace(/\s+/g,'').toUpperCase();
      const hasE = /\+E$/.test(raw) || /E$/.test(raw);
      const base = raw.replace(/\+E$/,'').replace(/E$/,'');
      const canon = canon_(base, hasE);
      if (canon) out.add('carnet_' + canon.toLowerCase());
   }

   return Array.from(out);
}

function computeIndexAndTags_(baseText, dict) {
   const toks = tokenize_(baseText);
   const specials = EMI_brainExtractSpecialTokens_(baseText);
   const expanded = new Set([].concat(toks || [], specials || []));

   const tags = new Set();

   expanded.forEach(t => {
      const hits = dict.reverse[t];
      if (!hits || !hits.length) return;
      if (hits.length > EMI_BRAIN_MAX_HITS_PER_TOKEN) return;

      hits.forEach(h => {
         if (!h || h.tipus !== 'TAG') return; // per scoring: només TAG detallat
         tags.add('TAG:' + h.codi);
      });
   });

   const finalToks = Array.from(expanded).filter(Boolean).slice(0, 2500);

   return {
      tags_auto: Array.from(tags).sort().join(';'),
      text_index: finalToks.join(' ')
   };
}

function normalizeFreeText_(s) {
   return String(s || '')
      .toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      // permet '_' perquè fem servir tokens especials (p.ex. carnet_b)
      .replace(/[^a-z0-9_]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
}

function normalizeToken_(s) {
   return normalizeFreeText_(s).split(' ')[0] || '';
}

// Stopwords bàsiques (CAT/ES/EN) per evitar soroll a tags_auto + matching
const EMI_BRAIN_MAX_HEADER_COLS = 600; // cap per evitar columnes fantasma i timeouts

const EMI_BRAIN_STOPWORDS = new Set([
   // CAT
   'a','al','als','amb','aquest','aquesta','aquests','aquestes','aqui','avui','cap','com','d','de','del','dels','des','doncs',
   'el','els','en','entre','era','es','est','esta','estem','estan','estar','fins','ha','han','he','hi','i','ja','la','les','li',
   'lo','m','mes','meu','meva','meus','meves','mi','molt','no','nosaltres','o','per','pero','poc','que','qui','s','se','sense',
   'ser','si','sota','també','te','tenim','tens','tinc','tot','tota','tots','totes','un','una','uns','unes','va','vam','van',
   // ES
   'a','al','algo','algunos','ante','antes','con','contra','como','de','del','desde','donde','dos','el','ella','ellas','ellos','en',
   'entre','era','es','esa','ese','eso','esta','este','estos','estas','hay','la','las','le','les','lo','los','mas','me','mi','mis',
   'muy','no','nos','o','para','pero','por','porque','que','quien','se','sin','sobre','su','sus','tambien','tengo','tiene','tienen',
   'todo','toda','todos','todas','tu','una','uno','unos','unas','ya',
   // EN
   'a','an','and','are','as','at','be','by','for','from','has','have','he','her','his','i','in','is','it','its','me','my','no',
   'of','on','or','our','she','so','that','the','their','them','they','this','to','was','we','were','with','you','your'
]);

// Tokens curts (2 caràcters) que sí que volem conservar
const EMI_BRAIN_ALLOW2 = new Set([
   // idiomes (nivells)
   'a1','a2','b1','b2','c1','c2',
   // sigles freqüents i útils
   'it','qa','fp','eso','ui','ux','hr','rr'
]);

// Si un token coincideix amb massa tags, és massa genèric i es descarta
const EMI_BRAIN_MAX_HITS_PER_TOKEN = 25;

function tokenize_(s) {
   const norm = normalizeFreeText_(s);
   if (!norm) return [];
   const parts = norm.split(' ').filter(Boolean);
   const out = [];

   for (let i = 0; i < parts.length; i++) {
      const t = String(parts[i] || '').trim();
      if (!t) continue;

      if (EMI_BRAIN_STOPWORDS.has(t)) continue;

      if (t.length >= 3) {
         out.push(t);
         continue;
      }

      if (t.length === 2) {
         if (EMI_BRAIN_ALLOW2.has(t)) out.push(t);
         continue;
      }

      // 1 lletra = soroll (els carnets es detecten com a token especial via EMI_brainExtractSpecialTokens_)
   }

   return out;
}

function indexMap2_(headers) {
   const m = {};
   headers.forEach((h, i) => m[String(h || '').trim()] = i + 1);
   return m;
}

// ====== MATCHING VACANTS ↔ PERSONES ======

function EMI_recomputeVacantMatchesNow(topN) {
   withLock_('MATCH_ALL', () => {
      const ss = getMasterSs_();
      ensureVacantSuggerimentsSheet_(ss);
      ensureBrainColumns_(ss);

      const dict = buildBrainDictionaries_();
      const cfg = readMatchConfig_(dict.catalogs);

      // 1) carrega persones
      const persones = loadEntitiesForMatch_(ss, SHEETS.PERSONES, 'id_persona', dict, buildPersonaBaseText_);

      // 2) carrega vacants
      const vacants = loadEntitiesForMatch_(ss, SHEETS.VACANTS, 'id_vacant', dict, buildVacantBaseText_);

      // 3) calcula suggeriments
      const outRows = [];
      const now = toIso_(new Date());
      const N = Number(topN || 25);

      vacants.forEach(v => {
         const scored = [];

         persones.forEach(p => {
            const scoreObj = scoreMatch_(v, p, cfg, dict.catalogs);
            if (scoreObj.score <= 0) return;
            scored.push({ p, ...scoreObj });
         });

         scored.sort((a,b) => b.score - a.score);
         scored.slice(0, N).forEach(s => {
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
      });

      // 4) escriu a sheet
      const sh = ss.getSheetByName('VACANT_SUGGERIMENTS');
      if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,8).clearContent();
      if (outRows.length) sh.getRange(2,1,outRows.length,8).setValues(outRows);

      SpreadsheetApp.flush();
      log_('INFO','MATCH_ALL_OK','Matching complet OK', { vacants: vacants.length, persones: persones.length, rows: outRows.length, cfg });
   });
}

function loadEntitiesForMatch_(ss, sheetName, idColName, dict, baseFn) {
   const sh = ss.getSheetByName(sheetName);
   if (!sh || sh.getLastRow() < 2) return [];

   const lastCol = Math.min(Math.max(1, sh.getLastColumn()), EMI_BRAIN_MAX_HEADER_COLS);
   const headersRaw = sh.getRange(1,1,1,lastCol).getValues()[0];
   const headers = headersRaw.map(h => String(h || '').trim());
   const idx = indexMap2_(headers);

   const idC = idx[idColName];
   const tagsAutoC = idx['tags_auto'];
   const textIndexC = idx['text_index'];

   if (!idC) return [];

   // Limitem el rang a les columnes amb header (evita llegir columnes fantasma i excedir el temps màxim)
   const lastHeaderCol = Math.max(
      EMI_lastHeaderCol_(headers),
      Number(idC || 0),
      Number(tagsAutoC || 0),
      Number(textIndexC || 0)
   );

   const rng = sh.getRange(2,1,sh.getLastRow()-1,lastHeaderCol);
   const vals = rng.getValues();

   function computeTagWeights_(tokens) {
      const w = {};
      let sum = 0;
      (tokens || []).forEach(t => {
         const hits = dict.reverse[t];
         if (!hits || !hits.length) return;
         if (hits.length > EMI_BRAIN_MAX_HITS_PER_TOKEN) return;
         hits.forEach(h => {
            if (!h || h.tipus !== 'TAG') return;
            const key = 'TAG:' + h.codi;
            const ww = Math.max(0.1, Math.min(5, Number(h.w || 1)));
            if (!w[key] || ww > w[key]) w[key] = ww;
         });
      });
      Object.keys(w).forEach(k => { sum += Number(w[k] || 0); });
      return { w, sum };
   }

   const out = [];
   vals.forEach(row => {
      const id = String(row[idC] || '').trim();
      if (!id) return;

      let text_index = textIndexC ? String(row[textIndexC] || '').trim() : '';
      let tags_auto = tagsAutoC ? String(row[tagsAutoC] || '').trim() : '';

      // si falta index/tags, ho recomputem (no destructiu)
      if (!text_index || !tags_auto) {
         const base = baseFn(row, idx);
         const comp = computeIndexAndTags_(base, dict);
         if (!text_index) text_index = comp.text_index;
         if (!tags_auto) tags_auto = comp.tags_auto;
      }

      const tokens = tokenize_(text_index);
      const tagsSet = new Set(parseTags_(tags_auto));
      const tw = computeTagWeights_(tokens);

      out.push({
         id,
         kw: new Set(tokens),
         tags: tagsSet,
         tagW: tw.w,
         tagWSum: tw.sum
      });
   });

   return out;
}

function parseTags_(s) {
   return String(s || '')
      .split(/[;,\n]+/)
      .map(x => x.trim())
      .filter(Boolean);
}

function scoreMatch_(vac, per, cfg, catalogs) {
   const cfgSafe = cfg || { kwWeight: 0.6, tagWeight: 0.4, bonusCarnet: 0, topKw: 8, topTags: 8, bonusTokens: {} };
   const kwHits = intersectCount_(vac.kw, per.kw);
   const denomKw = Math.sqrt(Math.max(1, vac.kw.size) * Math.max(1, per.kw.size));
   const kwScore = denomKw ? (kwHits / denomKw) : 0;

   let tagScore = 0;

   if (vac.tagW && per.tagW) {
      let num = 0;
      let interCount = 0;

      vac.tags.forEach(t => {
         if (!per.tags.has(t)) return;
         interCount++;
         const wv = Number(vac.tagW[t] || 1);
         const wp = Number(per.tagW[t] || 1);
         num += Math.min(wv, wp);
      });

      const denomTag = Math.sqrt(Math.max(1, vac.tagWSum || 0.0001) * Math.max(1, per.tagWSum || 0.0001));
      tagScore = denomTag ? (num / denomTag) : 0;

      // fallback suau si no hi ha pes real
      if (!isFinite(tagScore) || tagScore < 0) tagScore = 0;
      if (tagScore > 1) tagScore = 1;
   } else {
      const tagHits = intersectCount_(vac.tags, per.tags);
      tagScore = Math.min(1, tagHits / 5);
   }

   const kwWeight = Number(cfgSafe.kwWeight || 0);
   const tagWeight = Number(cfgSafe.tagWeight || 0);
   const denom = Math.max(0.0001, kwWeight + tagWeight);
   let score = ((kwWeight * kwScore) + (tagWeight * tagScore)) / denom;

   const bonus = computeSpecialBonus_(vac, per, cfgSafe);
   score = Math.min(1, score + bonus);

   const reasons = buildReasons_(vac, per, catalogs, cfgSafe);

   return {
      score: Number(score.toFixed(4)),
      reasons: reasons.text,
      kw_hits: reasons.kw_hits,
      tag_hits: reasons.tag_hits
   };
}

function intersectCount_(a, b) {
   let c = 0;
   a.forEach(x => { if (b.has(x)) c++; });
   return c;
}

function buildReasons_(vac, per, catalogs, cfg) {
   const cfgSafe = cfg || { topKw: 8, topTags: 8 };
   const maxKw = Math.max(1, Number(cfgSafe.topKw || 8));
   const maxTags = Math.max(1, Number(cfgSafe.topTags || 8));
   const kwCommon = [];
   vac.kw.forEach(x => { if (per.kw.has(x) && x.length >= 3) kwCommon.push(x); });

   const tagCommon = [];
   vac.tags.forEach(x => { if (per.tags.has(x)) tagCommon.push(x); });

   kwCommon.sort();
   tagCommon.sort();

   const kwTop = kwCommon.slice(0, maxKw).join(', ');
   const tagTop = tagCommonToLabels_(tagCommon, catalogs, maxTags).join(', ');
   const specialTop = intersectSpecialTokens_(vac.kw, per.kw, cfgSafe).join(', ');

   return {
      text: `KW: ${kwTop} | TAGS: ${tagTop}` + (specialTop ? ` | SPECIAL: ${specialTop}` : ''),
      kw_hits: kwCommon.length,
      tag_hits: tagCommon.length
   };
}

function tagCommonToLabels_(tagCommon, catalogs, maxItems) {
   const out = [];
   const seen = new Set();
   (tagCommon || []).forEach(t => {
      const m = String(t || '').match(/^([^:]+):(.+)$/);
      if (!m) return;
      const tipus = m[1].trim();
      const codi = m[2].trim();
      const key = tipus + '||' + codi;
      const entry = catalogs && catalogs[key];
      const label = entry && entry.valor ? String(entry.valor).trim() : codi;
      if (!label) return;
      const norm = normalizeToken_(label);
      if (seen.has(norm)) return;
      seen.add(norm);
      out.push(label);
   });
   out.sort();
   return out.slice(0, maxItems || 8);
}

function readMatchConfig_(catalogs) {
   const cfg = {
      kwWeight: 0.6,
      tagWeight: 0.4,
      bonusCarnet: 0.15,
      topKw: 8,
      topTags: 8,
      bonusTokens: {}
   };
   if (!catalogs) return cfg;

   Object.keys(catalogs).forEach(k => {
      const entry = catalogs[k];
      if (!entry || entry.tipus !== 'MATCH_CFG') return;
      const code = String(entry.codi || '').trim().toUpperCase();
      const raw = String(entry.valor || '').trim();
      if (!code) return;

      const num = parseFloat(raw.replace(',', '.'));
      if (code === 'KW_WEIGHT' && isFinite(num)) cfg.kwWeight = Math.max(0, num);
      if (code === 'TAG_WEIGHT' && isFinite(num)) cfg.tagWeight = Math.max(0, num);
      if (code === 'BONUS_CARNET' && isFinite(num)) cfg.bonusCarnet = Math.max(0, num);
      if (code === 'TOP_KW' && isFinite(num)) cfg.topKw = Math.max(1, Math.round(num));
      if (code === 'TOP_TAGS' && isFinite(num)) cfg.topTags = Math.max(1, Math.round(num));

      if (code === 'BONUS_TOKENS' && raw) {
         raw.split(/[;,\n]+/).forEach(part => {
            const [tokRaw, wRaw] = part.split('=').map(s => String(s || '').trim());
            const tok = normalizeToken_(tokRaw);
            const w = parseFloat(String(wRaw || '').replace(',', '.'));
            if (!tok || !isFinite(w)) return;
            cfg.bonusTokens[tok] = Math.max(0, w);
         });
      }
   });

   return cfg;
}

function intersectSpecialTokens_(kwA, kwB, cfg) {
   const cfgSafe = cfg || {};
   const out = [];
   kwA.forEach(t => {
      if (!kwB.has(t)) return;
      if (t.indexOf('carnet_') === 0) out.push(t);
      if (cfgSafe.bonusTokens && cfgSafe.bonusTokens[t]) out.push(t);
   });
   return Array.from(new Set(out)).sort();
}

function computeSpecialBonus_(vac, per, cfg) {
   const cfgSafe = cfg || {};
   let bonus = 0;
   const specials = intersectSpecialTokens_(vac.kw, per.kw, cfgSafe);
   if (!specials.length) return 0;

   const hasCarnet = specials.some(t => t.indexOf('carnet_') === 0);
   if (hasCarnet) bonus += Number(cfgSafe.bonusCarnet || 0);

   if (cfgSafe.bonusTokens) {
      specials.forEach(t => {
         const w = Number(cfgSafe.bonusTokens[t] || 0);
         if (w > 0) bonus += w;
      });
   }

   return Math.max(0, bonus);
}
