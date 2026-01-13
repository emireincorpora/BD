/*******************************************************
  * 16_DICT_HARVEST.gs — Diccionari “màxims” + ESCO + Pendents · V1.4 (SAFE + ESCO FIX)
  *******************************************************/

const EMI16 = {
   CATALEGS: 'CATALEGS',
   SINONIMS: 'SINONIMS',
   DICT_PENDENTS: 'DICT_PENDENTS',
   PERSONES: 'PERSONES',
   VACANTS: 'VACANTS',

   ESCO_BASE: 'https://ec.europa.eu/esco/api',

   CFG: {
      PERSONES_BATCH: 12,
      VACANTS_BATCH: 10,
      MAX_PENDENTS_PER_RUN: 180,

      // ESCO queda PAUSAT per defecte (segons requeriment)
      ESCO_ENABLED: false,
      ESCO_LIMIT: 12,
      ESCO_TYPES: ['occupation','skill']
   },

   PROP: {
      ESCO_FAM_CURSOR: 'EMI16_ESCO_FAM_CURSOR',
      PERSONES_CURSOR: 'EMI16_PERSONES_CURSOR',
      VACANTS_CURSOR: 'EMI16_VACANTS_CURSOR'
   }
};

function EMI_DICT_seedMaxims() {
   return EMI16_withLock_('DICT16_SEED', () => {
      const idx = EMI16_buildIndexes_();
      EMI16_seedCompanySectors_(idx);
      EMI16_seedCoreTags_(idx);
      EMI16_flush_(idx);

      EMI16_log_('INFO','DICT16_SEED_OK','Seed màxims aplicat', { cat_rows: idx.catData.length, syn_rows: idx.synData.length });
      EMI16_toast_('Seed diccionari OK ✅');
      return true;
   });
}

/** =========================
  *   ESCO toggle (Script Properties)
  *   ========================= */

function EMI_DICT_enableEscoNow() {
   setProp_('EMI16_ESCO_ENABLED', 'true');
   EMI16_log_('INFO','ESCO_ENABLED','ESCO activat (prop EMI16_ESCO_ENABLED=true)', {});
   return true;
}

function EMI_DICT_disableEscoNow() {
   setProp_('EMI16_ESCO_ENABLED', 'false');
   EMI16_log_('INFO','ESCO_DISABLED','ESCO pausat (prop EMI16_ESCO_ENABLED=false)', {});
   return true;
}

function EMI_DICT_escoStatusNow() {
   const v = EMI16_isEscoEnabled_();
   EMI16_log_('INFO','ESCO_STATUS','ESCO status', { enabled: v, defaultCfg: EMI16.CFG.ESCO_ENABLED, prop: getProp_('EMI16_ESCO_ENABLED') });
   return v;
}

function EMI_DICT_cleanupBrokenEscoRowsNow() {
   return EMI16_withLock_('ESCO_CLEAN', () => {
      const ss = EMI16_getMaster_();
      const sh = ss.getSheetByName(EMI16.CATALEGS);
      if (!sh || sh.getLastRow() < 2) return { scanned: 0, deactivated: 0 };

      const hm = headerMap_(sh);
      const cTip = hm['tipus'], cCodi = hm['codi'], cVal = hm['valor'], cAct = hm['actiu'];
      if (!cTip || !cCodi || !cVal || !cAct) return { scanned: 0, deactivated: 0, error: 'missing headers in CATALEGS' };

      const n = sh.getLastRow() - 1;
      const vals = sh.getRange(2, 1, n, sh.getLastColumn()).getValues();
      let deactivated = 0;
      for (let i = 0; i < n; i++) {
         const tipus = String(vals[i][cTip-1] || '').trim();
         const codi = String(vals[i][cCodi-1] || '').trim();
         const valor = String(vals[i][cVal-1] || '').trim();
         const actiu = String(vals[i][cAct-1] || '').trim();
         if (actiu === 'NO') continue;
         if (tipus === 'TAG' && (codi.indexOf('ESCO_') === 0) && (valor === '[object Object]' || valor.indexOf('[object Object]') >= 0)) {
            vals[i][cAct-1] = 'NO';
            deactivated++;
         }
      }
      if (deactivated) {
         sh.getRange(2, 1, n, sh.getLastColumn()).setValues(vals);
      }
      EMI16_log_('INFO','ESCO_CLEAN_DONE','Neteja ESCO trencat (valor=[object Object])', { scanned: n, deactivated });
      return { scanned: n, deactivated };
   });
}

function EMI16_isEscoEnabled_() {
   const p = getProp_('EMI16_ESCO_ENABLED');
   if (p === null || p === undefined || String(p).trim() === '') return !!EMI16.CFG.ESCO_ENABLED;
   return String(p).toLowerCase() === 'true';
}

function EMI_DICT_runHarvestBatch() {
   return EMI16_withLock_('DICT16_HARVEST', () => {
      const idx = EMI16_buildIndexes_();

      const escoEnabled = EMI16_isEscoEnabled_();

      const esco = escoEnabled
         ? EMI16_harvestEscoOneFamilySafe_(idx)
         : { addedTags: 0, addedSyn: 0, skipped: true };

      const per = EMI16_harvestPendentsFromPersonesSafe_(idx);
      const vac = EMI16_harvestPendentsFromVacantsSafe_(idx);

      EMI16_flush_(idx);

      EMI16_log_('INFO','DICT16_HARVEST_OK','Harvest complet (PENDENTS PERSONES + VACANTS)', { esco, per, vac });
      EMI16_toast_(`Harvest OK ✅ Pendents PERSONES:+${per.pendents} · VACANTS:+${vac.pendents} · ESCO:${esco.skipped ? 'PAUSAT' : ('+' + esco.addedTags + '/' + esco.addedSyn)}`);
      return { esco, per, vac };
   });
}

function EMI_DICT_runHarvestBatchMany(n) {
   n = Number(n || 5);
   const start = Date.now();
   const maxMs = 5.5 * 60 * 1000;

   const out = [];
   for (let i = 0; i < n; i++) {
      if ((Date.now() - start) > maxMs) {
         EMI16_log_('WARN','DICT16_MANY_STOP','Aturat per no petar de temps', { requested: n, done: i });
         break;
      }
      out.push(EMI_DICT_runHarvestBatch());
   }
   EMI16_toast_('Passades fetes ✅');
   return out;
}

/** =========================
  *   ESCO (SAFE + FIX parser)
  *   ========================= */

function EMI16_harvestEscoOneFamilySafe_(idx) {
   const fams = EMI16_listFamilies_(idx);
   if (!fams.length) return { addedTags: 0, addedSyn: 0, family: '', results: 0 };

   let cursor = Number(PropertiesService.getScriptProperties().getProperty(EMI16.PROP.ESCO_FAM_CURSOR) || 0);
   if (cursor >= fams.length) cursor = 0;

   const fam = fams[cursor];
   const kwList = EMI16_familyKeywords_(fam.code, fam.name);

   const search = EMI16_escoSearchMulti_(kwList, EMI16.CFG.ESCO_TYPES, EMI16.CFG.ESCO_LIMIT);

   let addedTags = 0, addedSyn = 0;

   search.results.forEach(item => {
      const uri = item.uri;
      const type = item.type;
      if (!uri || !type) return;

      const ca = EMI16_escoGetResource_(type, uri, 'ca');
      if (!ca || !ca.prefLabel) return;

      const es = EMI16_escoGetResource_(type, uri, 'es');

      const code = (type === 'occupation')
         ? ('ESCO_OCC_' + EMI16_hash10_(uri))
         : ('ESCO_SKL_' + EMI16_hash10_(uri));

      // Guardem ESCO com a TAG perquè el motor actual ja ho entén
      EMI16_upsertCat_(idx, 'TAG', code, ca.prefLabel, fam.code, '');

      addedTags++;

      (ca.altLabels || []).slice(0, 8).forEach(s => { EMI16_upsertSyn_(idx, 'TAG', code, s, 0.7, 'ESCO;lang=ca'); addedSyn++; });

      if (es && es.prefLabel) { EMI16_upsertSyn_(idx, 'TAG', code, es.prefLabel, 0.9, 'ESCO;lang=es'); addedSyn++; }
      (es && es.altLabels ? es.altLabels : []).slice(0, 8).forEach(s => { EMI16_upsertSyn_(idx, 'TAG', code, s, 0.6, 'ESCO;lang=es'); addedSyn++; });
   });

   PropertiesService.getScriptProperties().setProperty(EMI16.PROP.ESCO_FAM_CURSOR, String(cursor + 1));

   EMI16_log_('INFO','ESCO_HARVEST','ESCO 1 família (SAFE)', {
      family: fam.code,
      family_name: fam.name,
      keyword_used: search.keyword_used,
      lang_used: search.lang_used,
      http_code: search.http_code,
      total: search.total,
      results: search.results.length,
      addedTags,
      addedSyn
   });

   return { addedTags, addedSyn, family: fam.code, results: search.results.length, keyword_used: search.keyword_used, lang_used: search.lang_used };
}

function EMI16_escoSearchMulti_(keywords, types, limit) {
   const langs = ['es','en'];
   const list = (keywords || []).filter(Boolean);

   for (let li = 0; li < langs.length; li++) {
      const lang = langs[li];
      for (let ki = 0; ki < list.length; ki++) {
         const kw = list[ki];
         const r = EMI16_escoSearch_(kw, lang, types, limit);
         if (r.results.length) return { ...r, keyword_used: kw, lang_used: lang };
      }
   }
   // cap resultat: retornem igualment info útil
   return { results: [], keyword_used: list[0] || '', lang_used: langs[0], http_code: 200, total: 0 };
}

function EMI16_escoSearch_(text, lang, types, limit) {
   try {
      const params = [];
      params.push('text=' + encodeURIComponent(text || ''));
      params.push('language=' + encodeURIComponent(lang || 'es'));
      (types || []).forEach(t => params.push('type=' + encodeURIComponent(t)));
      params.push('offset=0');
      params.push('limit=' + encodeURIComponent(String(limit || 10)));
      params.push('full=false');
      params.push('selectedVersion=latest');
      params.push('viewObsolete=false');

      const url = EMI16.ESCO_BASE + '/search?' + params.join('&');
      const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, headers: { 'Accept': 'application/json' } });

      const code = res.getResponseCode();
      if (code !== 200) return { results: [], http_code: code, total: 0 };

      const json = JSON.parse(res.getContentText());

      const total =
         Number(json.total) ||
         Number(json.totalResults) ||
         Number(json.page && (json.page.totalElements || json.page.total)) ||
         0;

      const results = EMI16_extractEscoResults_(json);

      // Si ESCO diu que hi ha resultats però nosaltres no en traiem, ho deixem registrat
      if (total > 0 && results.length === 0) {
         EMI16_log_('WARN','ESCO_PARSE_EMPTY','ESCO total>0 però no he pogut extreure elements', {
            text, lang, total, sample: res.getContentText().slice(0, 350)
         });
      }

      return { results, http_code: code, total };
   } catch (e) {
      return { results: [], http_code: 0, total: 0 };
   }
}

function EMI16_extractEscoResults_(json) {
   const out = [];

   const add = (it) => {
      if (!it || typeof it !== 'object') return;

      const uri =
         it.uri ||
         (it._links && it._links.self && it._links.self.href) ||
         '';

      let typeRaw =
         it.type ||
         it.className ||
         it['@type'] ||
         '';

      typeRaw = String(typeRaw || '').toLowerCase();
      let type = '';
      if (typeRaw.indexOf('occupation') !== -1) type = 'occupation';
      if (typeRaw.indexOf('skill') !== -1) type = 'skill';

      if (!uri || !type) return;
      out.push({ uri, type });
   };

   const emb = json && json._embedded ? json._embedded : null;
   if (emb) {
      if (Array.isArray(emb.results)) emb.results.forEach(add);
      Object.keys(emb).forEach(k => {
         const v = emb[k];
         if (Array.isArray(v)) v.forEach(add);
      });
   }

   if (out.length === 0 && Array.isArray(json && json.results)) json.results.forEach(add);

   // dedupe per uri
   const seen = {};
   const uniq = [];
   for (let i = 0; i < out.length; i++) {
      const u = out[i].uri;
      if (seen[u]) continue;
      seen[u] = true;
      uniq.push(out[i]);
   }
   return uniq;
}

function EMI16_pickLangString_(v, lang) {
   if (v === null || v === undefined) return '';
   if (typeof v === 'string') return v;
   if (Array.isArray(v)) {
      for (let i = 0; i < v.length; i++) {
         const s = EMI16_pickLangString_(v[i], lang);
         if (s) return s;
      }
      return '';
   }
   if (typeof v === 'object') {
      const keys = [lang, 'ca', 'es', 'en', '@value', 'value', 'literal', 'label', 'text'];
      for (let i = 0; i < keys.length; i++) {
         const k = keys[i];
         if (typeof v[k] === 'string' && String(v[k]).trim()) return v[k];
      }
      // fallback: first string property
      for (const k in v) {
         if (typeof v[k] === 'string' && String(v[k]).trim()) return v[k];
      }
   }
   return '';
}

function EMI16_pickLangStringList_(v, lang) {
   if (v === null || v === undefined) return [];
   if (typeof v === 'string') return [v];
   if (Array.isArray(v)) {
      const out = [];
      v.forEach(it => {
         if (typeof it === 'string') out.push(it);
         else {
            const s = EMI16_pickLangString_(it, lang);
            if (s) out.push(s);
         }
      });
      return out;
   }
   if (typeof v === 'object') {
      // language map: {ca:[...]} or {ca:'...'}
      const keys = [lang, 'ca', 'es', 'en'];
      for (let i = 0; i < keys.length; i++) {
         const k = keys[i];
         const vv = v[k];
         if (typeof vv === 'string') return [vv];
         if (Array.isArray(vv)) return vv.map(x => String(x || '')).filter(Boolean);
      }
      const out = [];
      for (const k in v) {
         const vv = v[k];
         if (typeof vv === 'string') out.push(vv);
         else if (Array.isArray(vv)) out.push.apply(out, vv.map(x => String(x || '')).filter(Boolean));
      }
      return out;
   }
   return [];
}

function EMI16_escoGetResource_(type, uri, lang) {
   try {
      const endpoint = (type === 'occupation') ? '/resource/occupation' : '/resource/skill';
      const url = EMI16.ESCO_BASE + endpoint
         + '?uri=' + encodeURIComponent(uri)
         + '&language=' + encodeURIComponent(lang || 'ca')
         + '&selectedVersion=latest&viewObsolete=false';

      const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, headers: { 'Accept': 'application/json' } });
      if (res.getResponseCode() !== 200) return null;

      const json = JSON.parse(res.getContentText());
      const prefRaw = json.preferredLabel || json.prefLabel || json.title || '';
      const pref = EMI16_pickLangString_(prefRaw, lang || 'ca');

      const altsRaw = json.alternativeLabel || json.altLabels || json.alternativeLabels || [];
      let alts = EMI16_pickLangStringList_(altsRaw, lang || 'ca');
      if (!Array.isArray(alts)) alts = [];

      return { prefLabel: String(pref || '').trim(), altLabels: alts.map(x => String(x || '').trim()).filter(Boolean) };
   } catch (e) {
      return null;
   }
}

/** =========================
  *   PENDENTS des de PERSONES (SAFE)
  *   ========================= */

function EMI16_harvestPendentsFromPersonesSafe_(idx) {
   const ss = EMI16_getMaster_();
   const sh = ss.getSheetByName(EMI16.PERSONES);
   if (!sh || sh.getLastRow() < 2) return { pendents: 0, scanned: 0 };

   const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x => String(x||'').trim());
   const m = EMI16_indexMap_(headers);

   const idCol = m['id_persona'] || m['emi_id_persona'] || 0;
   const cols = [m['cv_sector_objectiu'], m['cv_competencies'], m['cv_experiencia'], m['cv_formacio']].filter(Boolean);
   if (!cols.length) return { pendents: 0, scanned: 0 };

   let cursor = Number(PropertiesService.getScriptProperties().getProperty(EMI16.PROP.PERSONES_CURSOR) || 2);
   if (cursor < 2) cursor = 2;
   if (cursor > sh.getLastRow()) cursor = 2;

   const batch = Math.min(EMI16.CFG.PERSONES_BATCH, sh.getLastRow() - cursor + 1);
   const data = sh.getRange(cursor, 1, batch, sh.getLastColumn()).getValues();

   const variantMap = EMI16_buildVariantMap_(idx);
   const seenPending = EMI16_buildPendingSetFast_(ss);

   const outRows = [];
   const now = new Date();
   const maxPend = EMI16.CFG.MAX_PENDENTS_PER_RUN;

   data.forEach(row => {
      if (outRows.length >= maxPend) return;

      const pid = idCol ? String(row[idCol - 1] || '').trim() : '';

      cols.forEach(c => {
         if (outRows.length >= maxPend) return;

         const raw = String(row[c - 1] || '').trim();
         if (!raw) return;

         const terms = EMI16_extractTerms_(raw);
         terms.forEach(t => {
            if (outRows.length >= maxPend) return;

            const norm = EMI16_norm_(t);
            if (!norm || norm.length < 3) return;
            if (variantMap[norm]) return;
            if (seenPending.has(norm)) return;

            outRows.push(['TERMES_PERSONA_CV', t, norm, '', '', 0, 'PERSONES', now, pid ? ('id_persona=' + pid) : '']);
            seenPending.add(norm);
         });
      });
   });

   if (outRows.length) {
      const p = EMI16_ensureSheet_(ss, EMI16.DICT_PENDENTS);
      EMI16_ensureHeader_(p, ['tipus','text_original','text_normalitzat','proposta_codi_tag','proposta_tag','score','font','created_at','notes']);
      p.getRange(p.getLastRow() + 1, 1, outRows.length, 9).setValues(outRows);
   }

   PropertiesService.getScriptProperties().setProperty(EMI16.PROP.PERSONES_CURSOR, String(cursor + batch));

   EMI16_log_('INFO','PENDENTS_FROM_PERSONES','Pendents des de PERSONES (SAFE)', { cursor, batch, pendents: outRows.length });
   return { pendents: outRows.length, scanned: batch };
}

/** =========================
  *   PENDENTS des de VACANTS (SAFE)
  *   ========================= */

function EMI16_harvestPendentsFromVacantsSafe_(idx) {
   const ss = EMI16_getMaster_();
   const sh = ss.getSheetByName(EMI16.VACANTS);
   if (!sh || sh.getLastRow() < 2) return { pendents: 0, scanned: 0 };

   const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x => String(x||'').trim());
   const m = EMI16_indexMap_(headers);

   const idCol = m['id_vacant'] || m['emi_id_vacant'] || 0;
   const idEmpresaCol = m['id_empresa'] || m['empresa_id'] || 0;
   const cols = [m['titol_vacant'], m['descripcio'], m['requisits'], m['sector']].filter(Boolean);
   if (!cols.length) return { pendents: 0, scanned: 0 };

   let cursor = Number(PropertiesService.getScriptProperties().getProperty(EMI16.PROP.VACANTS_CURSOR) || 2);
   if (cursor < 2) cursor = 2;
   if (cursor > sh.getLastRow()) cursor = 2;

   const batch = Math.min(EMI16.CFG.VACANTS_BATCH, sh.getLastRow() - cursor + 1);
   const data = sh.getRange(cursor, 1, batch, sh.getLastColumn()).getValues();

   const variantMap = EMI16_buildVariantMap_(idx);
   const seenPending = EMI16_buildPendingSetFast_(ss);

   const outRows = [];
   const now = new Date();
   const maxPend = EMI16.CFG.MAX_PENDENTS_PER_RUN;

   data.forEach(row => {
      if (outRows.length >= maxPend) return;

      const vid = idCol ? String(row[idCol - 1] || '').trim() : '';
      const eid = idEmpresaCol ? String(row[idEmpresaCol - 1] || '').trim() : '';

      cols.forEach(c => {
         if (outRows.length >= maxPend) return;
         const raw = String(row[c - 1] || '').trim();
         if (!raw) return;

         const terms = EMI16_extractTerms_(raw);
         terms.forEach(t => {
            if (outRows.length >= maxPend) return;

            const norm = EMI16_norm_(t);
            if (!norm) return;

            // Permetem tokens curts NOMÉS si són útils (idiomes/abreviatures). La resta, evitem soroll.
            if (norm.length < 3 && !EMI16_isAllowedShortPendingToken_(norm)) return;

            if (variantMap[norm]) return;
            if (seenPending.has(norm)) return;

            const notesParts = [];
            if (vid) notesParts.push('id_vacant=' + vid);
            if (eid) notesParts.push('id_empresa=' + eid);

            outRows.push(['TERMES_VACANT', t, norm, '', '', 0, 'VACANTS', now, notesParts.join(' ; ')]);
            seenPending.add(norm);
         });
      });
   });

   if (outRows.length) {
      const p = EMI16_ensureSheet_(ss, EMI16.DICT_PENDENTS);
      EMI16_ensureHeader_(p, ['tipus','text_original','text_normalitzat','proposta_codi_tag','proposta_tag','score','font','created_at','notes']);
      p.getRange(p.getLastRow() + 1, 1, outRows.length, 9).setValues(outRows);
   }

   PropertiesService.getScriptProperties().setProperty(EMI16.PROP.VACANTS_CURSOR, String(cursor + batch));

   EMI16_log_('INFO','PENDENTS_FROM_VACANTS','Pendents des de VACANTS (SAFE)', { cursor, batch, pendents: outRows.length });
   return { pendents: outRows.length, scanned: batch };
}

function EMI16_isAllowedShortPendingToken_(norm) {
   norm = String(norm || '').toLowerCase().trim();
   // Idiomes nivells (A1..C2)
   if (/^[abc][12]$/.test(norm)) return true;
   // Sigles freqüents (evitem omplir DICT_PENDENTS amb codis aleatoris)
   const ok = {
      'it':1,'qa':1,'fp':1,'eso':1,'ux':1,'ui':1,'rr':1,'hr':1,'am':1
   };
   return !!ok[norm];
}

function EMI16_buildPendingSetFast_(ss) {
   const sh = ss.getSheetByName(EMI16.DICT_PENDENTS);
   const set = new Set();
   if (!sh || sh.getLastRow() < 2) return set;

   const lastCol = sh.getLastColumn();
   const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(x => String(x||'').trim());
   const m = EMI16_indexMap_(headers);
   const colNorm = m['text_normalitzat'];
   if (!colNorm) return set;

   const vals = sh.getRange(2, colNorm, sh.getLastRow()-1, 1).getValues();
   vals.forEach(v => { const s = String(v[0] || '').trim(); if (s) set.add(s); });
   return set;
}

/** =========================
  *   Keywords per família (ESCO)
  *   ========================= */

function EMI16_familyKeywords_(famCode, famNameCa) {
  // Paraules clau (ES/EN) per buscar a ESCO.
  // IMPORTANT: "famNameCa" és català; ESCO cerca sobretot en ES/EN.
  const K = {
    'AA': ['artesanía','cerámica','joyería','talla madera','handicraft'],
    'AE': ['deporte','monitor deportivo','entrenador','socorrismo','sports'],
    'AF': ['diseño gráfico','preimpresión','impresión','maquetación','graphic'],
    'AG': ['administración','auxiliar administrativo','contabilidad','gestión','office'],
    'AR': ['agricultura','jardinería','ganadería','forestal','agriculture'],
    'CM': ['comercio','ventas','dependiente','marketing','retail'],
    'EA': ['energías renovables','fotovoltaica','eólica','agua','energy'],
    'EE': ['electricidad','electrónica','automatización','plc','electrical'],
    'EO': ['construcción','albañil','obra','reformas','construction'],
    'FM': ['soldador','mecanizado','cnc','calderería','welding'],
    'HS': ['hostelería','cocinero','camarero','turismo','hospitality'],
    'IA': ['industria alimentaria','alimentación','panadería','carnicería','food'],
    'IM': ['climatización','frigorista','mantenimiento industrial','instalador','maintenance'],
    'IP': ['peluquería','estética','imagen personal','cosmética','beauty'],
    'IS': ['audiovisual','sonido','vídeo','producción audiovisual','media'],
    'MP': ['pesca','marítimo','acuicultura','navegación','fishing'],
    'QU': ['química','laboratorio','procesos químicos','operador químico','chemical'],
    'SA': ['sanidad','auxiliar de enfermería','farmacia','clínica','health'],
    'SC': ['servicios sociales','educación infantil','integración social','animación sociocultural','community'],
    'SM': ['seguridad','prevención riesgos','medio ambiente','prl','safety'],
    'TM': ['automoción','mecánica','mantenimiento vehículos','conductor','automotive'],
    'TX': ['textil','confección','costura','patronaje','fashion'],
    'IF': ['informática','programación','soporte técnico','redes','it'],
    'AD': ['arte','música','danza','teatro','arts']
  };

  const list = (K[famCode] || []).slice();

  // Fallback: fem servir 2-3 paraules del nom en català (algunes coincideixen amb ES/EN)
  if (!list.length) {
    const toks = String(famNameCa || '')
      .toLowerCase()
      .replace(/[^a-z0-9àèéíïòóúüçñ\s]/g, ' ')
      .split(/\s+/)
      .filter(Boolean);
    toks.slice(0, 3).forEach(t => list.push(t));
    // i una paraula genèrica que sovint retorna ocupacions
    list.push('ocupación');
  }
  return list;
}

/** =========================
  *   Seed base (mínim)
  *   ========================= */

function EMI16_seedCompanySectors_(idx) {
   const sectors = [
      ['METALL', 'Metall', ['metall','soldadura','caldereria','serralleria','mecanitzat','cnc'], ['metal','soldadura','calderería','cerrajería','mecanizado','cnc']]
   ];
   sectors.forEach((s, i) => {
      EMI16_upsertCat_(idx, 'SECTOR_EMPRESA', s[0], s[1], '', i + 10);
      s[2].forEach(v => EMI16_upsertSyn_(idx, 'SECTOR_EMPRESA', s[0], v, 1.0, 'seed;lang=ca'));
      s[3].forEach(v => EMI16_upsertSyn_(idx, 'SECTOR_EMPRESA', s[0], v, 1.0, 'seed;lang=es'));
   });
}

function EMI16_seedCoreTags_(idx) {
   // CORE (mínim) — pots ampliar manualment
   const CORE = [
      ['FM','T_SOLDADURA_TIG','Soldadura TIG', ['tig','soldadura tig'], ['tig','soldadura tig']],
      ['FM','T_CNC','CNC', ['cnc','control numèric'], ['cnc','control numérico']],
      ['CM','T_CARRETONER','Carretoner', ['carretó elevador'], ['carretilla elevadora']]
   ];

   CORE.forEach((x, n) => {
      const fam = String(x[0] || '').trim();
      const code = String(x[1] || '').trim();
      const label = String(x[2] || '').trim();
      (x[3]||[]).forEach(s => EMI16_upsertSyn_(idx, 'TAG', code, s, 1.0, 'seed;lang=ca'));
      (x[4]||[]).forEach(s => EMI16_upsertSyn_(idx, 'TAG', code, s, 1.0, 'seed;lang=es'));
      EMI16_upsertCat_(idx, 'TAG', code, label, fam, 1000 + n);
   });

   // Permisos/carnets de conduir (Espanya / UE) — DETECCIÓ COMPLETA
   const DL = ['AM','A1','A2','A','B1','B','BE','C1','C1E','C','CE','D1','D1E','D','DE'];

   function dlDisplay_(code) {
      // Mostra amb "+E" quan toca
      if (code === 'BE') return 'B+E';
      if (code === 'C1E') return 'C1+E';
      if (code === 'CE') return 'C+E';
      if (code === 'D1E') return 'D1+E';
      if (code === 'DE') return 'D+E';
      return code;
   }

   DL.forEach((c, i) => {
      const code = 'DL_' + c;
      const disp = dlDisplay_(c);
      const label = 'Carnet ' + disp;

      // token especial (coincideix amb el cervell): carnet_be, carnet_c1e, etc.
      const specialTok = 'carnet_' + String(c).toLowerCase();

      // sinònims CA
      EMI16_upsertSyn_(idx, 'TAG', code, 'carnet ' + disp.toLowerCase(), 1.0, 'seed;lang=ca');
      EMI16_upsertSyn_(idx, 'TAG', code, 'permis ' + disp.toLowerCase(), 1.0, 'seed;lang=ca');
      EMI16_upsertSyn_(idx, 'TAG', code, 'permis de conduir ' + disp.toLowerCase(), 1.0, 'seed;lang=ca');
      EMI16_upsertSyn_(idx, 'TAG', code, specialTok, 1.2, 'seed;special');

      // sinònims ES
      EMI16_upsertSyn_(idx, 'TAG', code, 'carnet ' + disp.toLowerCase(), 1.0, 'seed;lang=es');
      EMI16_upsertSyn_(idx, 'TAG', code, 'permiso ' + disp.toLowerCase(), 1.0, 'seed;lang=es');
      EMI16_upsertSyn_(idx, 'TAG', code, 'permiso de conducir ' + disp.toLowerCase(), 1.0, 'seed;lang=es');

      EMI16_upsertCat_(idx, 'TAG', code, label, '', 2000 + i);
   });
}

/** =========================
  *   Indexos + IO
  *   ========================= */

function EMI16_buildIndexes_() {
   const ss = EMI16_getMaster_();
   const catSh = EMI16_ensureSheet_(ss, EMI16.CATALEGS);
   const synSh = EMI16_ensureSheet_(ss, EMI16.SINONIMS);
   EMI16_ensureSheet_(ss, EMI16.DICT_PENDENTS);

   EMI16_ensureHeader_(catSh, ['tipus','codi','valor','etiqueta','actiu','ordre']);
   EMI16_ensureHeader_(synSh, ['tipus','codi_ref','sinonim','pes','actiu','notes','id_sinonim']);

   const catData = (catSh.getLastRow() >= 2)
      ? catSh.getRange(2,1,catSh.getLastRow()-1,6).getValues().map(r => ({ tipus:r[0], codi:r[1], valor:r[2], etiqueta:r[3], actiu:r[4], ordre:r[5] }))
      : [];

   const synData = (synSh.getLastRow() >= 2)
      ? synSh.getRange(2,1,synSh.getLastRow()-1,7).getValues().map(r => ({ tipus:r[0], codi_ref:r[1], sinonim:r[2], pes:r[3], actiu:r[4], notes:r[5], id_sinonim:r[6] }))
      : [];

   const catIndex = {};
   catData.forEach((x,i) => { const t = String(x.tipus||'').trim(); const c = String(x.codi||'').trim(); if (t && c) catIndex[t + '||' + c] = i; });

   const synIndex = {};
   synData.forEach((x,i) => {
      const t = String(x.tipus||'').trim();
      const c = String(x.codi_ref||'').trim();
      const s = EMI16_norm_(x.sinonim);
      if (t && c && s) synIndex[t + '||' + c + '||' + s] = i;
   });

   return { ss, catSh, synSh, catData, synData, catIndex, synIndex, catDirty:false, synDirty:false };
}

function EMI16_upsertCat_(idx, tipus, codi, valor, etiqueta, ordre) {
   const key = String(tipus||'').trim() + '||' + String(codi||'').trim();
   if (!tipus || !codi || !valor) return;

   const rec = { tipus:String(tipus).trim(), codi:String(codi).trim(), valor:String(valor).trim(), etiqueta:String(etiqueta||'').trim(), actiu:'SI', ordre: ordre === undefined ? '' : ordre };

   if (idx.catIndex[key] !== undefined) idx.catData[idx.catIndex[key]] = rec;
   else { idx.catIndex[key] = idx.catData.length; idx.catData.push(rec); }

   idx.catDirty = true;
}

function EMI16_upsertSyn_(idx, tipus, codiRef, sinonim, pes, notes) {
   if (!tipus || !codiRef || !sinonim) return;
   const norm = EMI16_norm_(sinonim);
   if (!norm) return;

   const key = String(tipus).trim() + '||' + String(codiRef).trim() + '||' + norm;

   const rec = {
      tipus:String(tipus).trim(), codi_ref:String(codiRef).trim(), sinonim:String(sinonim).trim(),
      pes:Number(pes||1), actiu:'SI', notes:String(notes||''), id_sinonim:'SYN-' + EMI16_hash10_(key)
   };

   if (idx.synIndex[key] !== undefined) {
      const i = idx.synIndex[key];
      const oldId = idx.synData[i].id_sinonim;
      if (oldId && String(oldId).trim()) rec.id_sinonim = oldId;
      idx.synData[i] = rec;
   } else {
      idx.synIndex[key] = idx.synData.length;
      idx.synData.push(rec);
   }
   idx.synDirty = true;
}

function EMI16_flush_(idx) {
   if (idx.catDirty) {
      const out = idx.catData.map(x => [x.tipus,x.codi,x.valor,x.etiqueta,x.actiu,x.ordre]);
      if (idx.catSh.getLastRow() > 1) idx.catSh.getRange(2,1,idx.catSh.getLastRow()-1,6).clearContent();
      if (out.length) idx.catSh.getRange(2,1,out.length,6).setValues(out);
      idx.catDirty = false;
   }
   if (idx.synDirty) {
      const out = idx.synData.map(x => [x.tipus,x.codi_ref,x.sinonim,x.pes,x.actiu,x.notes,x.id_sinonim]);
      if (idx.synSh.getLastRow() > 1) idx.synSh.getRange(2,1,idx.synSh.getLastRow()-1,7).clearContent();
      if (out.length) idx.synSh.getRange(2,1,out.length,7).setValues(out);
      idx.synDirty = false;
   }
   SpreadsheetApp.flush();
}

/** =========================
  *   Helpers
  *   ========================= */

function EMI16_listFamilies_(idx) {
   const fams = [];
   idx.catData.forEach(r => {
      if (String(r.tipus||'').trim() !== 'FPCAT_FAMILIA') return;
      const c = String(r.codi||'').trim();
      const v = String(r.valor||'').trim();
      const a = String(r.actiu||'').trim().toUpperCase();
      if (!c || !v) return;
      if (!(a === 'SI' || a === 'SÍ' || a === 'YES' || a === 'TRUE' || a === '1')) return;
      fams.push({ code: c, name: v });
   });
   return fams;
}

function EMI16_buildVariantMap_(idx) {
   const map = {};
   idx.catData.forEach(r => {
      const c = String(r.codi||'').trim();
      const v = String(r.valor||'').trim();
      const a = String(r.actiu||'').trim().toUpperCase();
      if (!c || !v) return;
      if (!(a === 'SI' || a === 'SÍ' || a === 'YES' || a === 'TRUE' || a === '1')) return;
      map[EMI16_norm_(v)] = c;
   });
   idx.synData.forEach(s => {
      const c = String(s.codi_ref||'').trim();
      const v = String(s.sinonim||'').trim();
      const a = String(s.actiu||'').trim().toUpperCase();
      if (!c || !v) return;
      if (!(a === 'SI' || a === 'SÍ' || a === 'YES' || a === 'TRUE' || a === '1')) return;
      map[EMI16_norm_(v)] = c;
   });
   return map;
}

function EMI16_extractTerms_(text) {
   const s = String(text||'').replace(/\r/g,'\n').replace(/[•·|]/g,';').replace(/\s{2,}/g,' ').trim();
   const rawParts = s.split(/[\n;,/]+/).map(x => String(x||'').trim()).filter(Boolean);
   const out = [];
   rawParts.forEach(p => {
      let x = p.trim();
      if (x.length < 3) return;
      if (x.length > 80) x = x.slice(0,80).trim();
      out.push(x);
   });
   const seen = {};
   return out.filter(x => {
      const n = EMI16_norm_(x);
      if (!n) return false;
      if (seen[n]) return false;
      seen[n] = true;
      return true;
   }).slice(0, 30);
}

function EMI16_getMaster_() {
   try { if (typeof getMasterSs_ === 'function') return getMasterSs_(); } catch(e) {}
   return SpreadsheetApp.getActiveSpreadsheet();
}

function EMI16_withLock_(name, timeoutOrFn, maybeFn) {
   // IMPORTANT: no llancem error si el lock està ocupat; fem "skip" i ho provarem al següent tick.
   // Suporta dues signatures:
   //  - EMI16_withLock_('NAME', () => {...})
   //  - EMI16_withLock_('NAME', 12000, () => {...})

   let timeoutMs = 5000;
   let fn = null;
   if (typeof timeoutOrFn === 'function') {
      fn = timeoutOrFn;
   } else {
      timeoutMs = Number(timeoutOrFn || 5000);
      fn = maybeFn;
   }
   if (typeof fn !== 'function') throw new Error('EMI16_withLock_: falta callback fn()');

   const lock = LockService.getScriptLock();
   const wait = Math.max(100, Math.min(timeoutMs, 30000));
   const ok = lock.tryLock(wait);
   if (!ok) {
      try { EMI16_log_('WARN','DICT16_LOCK_BUSY','Skip run: lock ocupat', { name, wait }); } catch (e1) {
         try { if (typeof log_ === 'function') log_('WARN','DICT16_LOCK_BUSY','Skip run: lock ocupat', { name, wait }); } catch (e2) {}
      }
      return { skipped: true, reason: 'lock_busy', name, wait };
   }

   try { return fn(); }
   finally { try { lock.releaseLock(); } catch (e3) {} }
}

function EMI16_log_(level, action, message, obj) {
   try { if (typeof log_ === 'function') return log_(level, action, message, obj || {}); } catch(e) {}
   try { console.log(level, action, message, obj || {}); } catch(e2) {}
}

function EMI16_toast_(msg) {
   try { const ss = EMI16_getMaster_(); if (ss && ss.toast) ss.toast(String(msg||''), 'EMI', 5); } catch(e) {}
}

function EMI16_norm_(s) {
   s = String(s || '').toLowerCase().trim();
   try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch(e) {}
   s = s.replace(/[^a-z0-9]+/g,' ').replace(/\s+/g,' ').trim();
   return s;
}

function EMI16_hash10_(s) {
   const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(s||''), Utilities.Charset.UTF_8);
   const hex = raw.map(b => {
      const v = (b < 0) ? b + 256 : b;
      return (v < 16 ? '0' : '') + v.toString(16);
   }).join('');
   return hex.substring(0,10);
}

function EMI16_indexMap_(headers) {
   const m = {};
   headers.forEach((h,i)=>{ const k = String(h||'').trim(); if(k) m[k]=i+1; });
   return m;
}

function EMI16_ensureSheet_(ss, name) {
   let sh = ss.getSheetByName(name);
   if (!sh) sh = ss.insertSheet(name);
   return sh;
}

function EMI16_ensureHeader_(sh, headers) {
   const lc = Math.max(headers.length, sh.getLastColumn() || headers.length);
   const cur = sh.getRange(1,1,1,lc).getValues()[0];
   let ok = true;
   for (let i = 0; i < headers.length; i++) {
      if (String(cur[i]||'').trim() !== headers[i]) { ok = false; break; }
   }
   if (!ok) sh.getRange(1,1,1,headers.length).setValues([headers]);
}

/*******************************************************
 * TAULA INVERSA (debug del diccionari)
 * - Construeix una taula "DICT_INVERSA" amb:
 *   token -> (tipus, codi, valor, etiqueta)
 * - És només per inspecció / depuració (no cal per al matching).
 *******************************************************/
function EMI_DICT_buildInverseTableNow(maxRows) {
  maxRows = Number(maxRows || 50000);

  return EMI16_withLock_('DICT_INVERSE', 120, () => {
    const ss = EMI16_getMaster_();
    const sh = EMI16_ensureSheet_(ss, 'DICT_INVERSA');
    const headers = ['token','tipus','codi','valor','etiqueta'];

    // Construïm el diccionari "com el veu" el motor de matching
    if (typeof buildBrainDictionaries_ !== 'function') {
      throw new Error('Falta buildBrainDictionaries_(). Revisa 10_BRAIN_INDEX_MATCH.gs');
    }
    const dict = buildBrainDictionaries_();
    const reverse = dict && dict.reverse ? dict.reverse : {};
    const tokens = Object.keys(reverse).sort();

    const rows = [];
    for (let i = 0; i < tokens.length; i++) {
      const tok = tokens[i];
      const arr = reverse[tok] || [];
      for (let j = 0; j < arr.length; j++) {
        const e = arr[j] || {};
        rows.push([tok, e.tipus || '', e.codi || '', e.valor || '', e.etiqueta || '']);
        if (rows.length >= maxRows) break;
      }
      if (rows.length >= maxRows) break;
    }

    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) {
      sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    EMI16_log_('INFO','DICT_INVERSE_OK','Taula inversa generada', { tokens: tokens.length, rows: rows.length, maxRows });
    EMI16_toast_('Taula inversa creada ✅');
    return { tokens: tokens.length, rows: rows.length };
  });
}
