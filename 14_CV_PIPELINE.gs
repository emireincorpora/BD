/*******************************************************
  * 14_CV_PIPELINE.gs
  * - Handler estable: EMI_CV_ingestNewResponses()
  * - Manté CV source 100% read-only
  * - Merge de CV_RAW -> PERSONES amb política segura
  *******************************************************/

// Política de camps (PERSONES)
// A) Sempre (només si hi ha canvis): updated_by_email, data_ultima_actualitzacio
// B) Omplir si buit: email, telefon_mobil, doc_numero, municipi, comarca, nom, cognom1
// C) Mai sobreescriure: tecnic_referent, tags_persona, ambits_interes, programes_interes, observacions, situacions, etc.

function EMI_CV_ingestNewResponses() {
   // 1) Importa noves respostes a CV_RAW (funció existent del teu 11_CV_INGEST.gs)
   if (typeof EMI_CV_pullNewRows !== 'function') {
      throw new Error('No existeix EMI_CV_pullNewRows(). Revisa 11_CV_INGEST.gs');
   }

   const result = EMI_CV_pullNewRows(); // NO escriu al CV source
   if (!result || !result.imported || !result.fromRow || !result.toRow) return result;

   // 2) Merge CV_RAW -> PERSONES (només files importades ara)
   EMI_CV_mergeCvRawToPersones_(result.fromRow, result.toRow);
   return result;
}

function EMI_CV_mergeCvRawToPersones_(fromSourceRow, toSourceRow) {
   const runner = () => {
      const ss = (typeof getMasterSs_ === 'function') ? getMasterSs_() : SpreadsheetApp.getActiveSpreadsheet();

      const shCv = ss.getSheetByName('CV_RAW');
      if (!shCv) throw new Error('No existeix CV_RAW (hauria de crear-se amb EMI_CV_pullNewRows)');

      const shP = (typeof getSheet_ === 'function' && typeof SHEETS !== 'undefined')
         ? getSheet_(SHEETS.PERSONES)
         : ss.getSheetByName('PERSONES');

      if (!shP) throw new Error('No trobo la pestanya PERSONES');

      const cvHeader = shCv.getRange(1,1,1,shCv.getLastColumn()).getValues()[0].map(String);
      const pHeader   = shP.getRange(1,1,1,shP.getLastColumn()).getValues()[0].map(String);

      const idxCv = indexMap_(cvHeader);
      const idxP   = indexMap_(pHeader);

      // Columnes CV_RAW necessàries
      const c_source = idxCv['source_row'];
      const c_email   = idxCv['email'];
      const c_nom      = idxCv['nom'];
      const c_cog      = idxCv['cognoms'];
      const c_tel      = idxCv['telefon'];
      const c_doc      = idxCv['doc'];
      const c_mun      = idxCv['municipi'];
      const c_com      = idxCv['comarca'];
      const c_link    = idxCv['linked_id_persona'];
      const c_status = idxCv['status'];
      const c_notes   = idxCv['notes'];

      if (c_source < 0) throw new Error('CV_RAW necessita la columna source_row');

      // Llegeix totes les files de CV_RAW i filtra per source_row (rang importat)
      const lastCvRow = shCv.getLastRow();
      if (lastCvRow <= 1) return;

      const cvValues = shCv.getRange(2, 1, lastCvRow - 1, shCv.getLastColumn()).getValues();

      // Crea índex de PERSONES 1 cop (email / tel / name) per dedupe
      const lookup = buildPersonLookup_(shP, pHeader, idxP);

      let touchedCv = 0;

      for (let i = 0; i < cvValues.length; i++) {
         const sheetRow = i + 2;
         const sourceRow = Number(cvValues[i][c_source] || 0);
         if (sourceRow < fromSourceRow || sourceRow > toSourceRow) continue;

         const extracted = {
            email:    safeStr_(cvValues[i], c_email).toLowerCase(),
            nom:       safeStr_(cvValues[i], c_nom),
            cognoms: safeStr_(cvValues[i], c_cog),
            telefon: safeStr_(cvValues[i], c_tel),
            doc:       safeStr_(cvValues[i], c_doc),
            municipi:safeStr_(cvValues[i], c_mun),
            comarca: safeStr_(cvValues[i], c_com),
         };

         let linked = safeStr_(cvValues[i], c_link);
         let status = 'LINKED';
         let notes   = '';

         // 1) Si ja ve linkat per la ingesta, fem update si cal
         if (linked) {
            const upd = updatePersonFromCv_(shP, pHeader, idxP, lookup, linked, extracted, 'linked');
            status = upd.updated ? 'UPDATED' : 'LINKED';
            notes = upd.notes || '';
         } else {
            // 2) Dedupe: email > tel > nom+cognoms (només si únic)
            const match = findMatch_(lookup, extracted);

            if (match && match.ambiguous) {
               status = 'SKIPPED';
               notes = 'Ambiguous match (' + match.reason + ')';
               linked = '';
            } else if (match && match.id_persona) {
               linked = match.id_persona;
               const upd = updatePersonFromCv_(shP, pHeader, idxP, lookup, linked, extracted, match.by);
               status = upd.updated ? 'UPDATED' : 'LINKED';
               notes = upd.notes || '';
            } else {
               // 3) Crear persona (només si tenim email o tel)
               const email = extracted.email;
               const telNorm = normalizePhone_(extracted.telefon);
               if (!email && !telNorm) {
                  status = 'SKIPPED';
                  notes = 'No email ni telèfon: no creo persona (evitem duplicats)';
                  linked = '';
               } else {
                  const created = createPersonFromCv_(shP, pHeader, idxP, lookup, extracted);
                  linked = created.id_persona;
                  status = 'CREATED';
                  notes = 'Created from CV';
               }
            }
         }

         // Escriu de tornada a CV_RAW (ok, és MASTER)
         if (c_link >= 0) cvValues[i][c_link] = linked;
         if (c_status >= 0) cvValues[i][c_status] = status;
         if (c_notes >= 0) cvValues[i][c_notes] = notes;

         touchedCv++;
      }

      if (touchedCv) {
         shCv.getRange(2, 1, cvValues.length, shCv.getLastColumn()).setValues(cvValues);
      }

      try { if (typeof log_ === 'function') log_('INFO','CV_MERGE_DONE','Merge CV_RAW -> PERSONES completat', { fromSourceRow, toSourceRow, touchedCv }); } catch (_) {}
      return true;
   };

   if (typeof withLock_ === 'function') return withLock_('CV_MERGE', runner);
   return runner();
}

/******** Helpers ********/

function indexMap_(header) {
   const m = {};
   for (let i = 0; i < header.length; i++) m[String(header[i] || '').trim()] = i;
   return m;
}

function safeStr_(row, idx) {
   if (idx === undefined || idx === null || idx < 0) return '';
   const v = row[idx];
   return (v === null || v === undefined) ? '' : String(v).trim();
}

function normalizePhone_(s) {
   s = String(s || '');
   // només dígits i +
   const digits = s.replace(/[^\d+]/g, '');
   // normalització simple per comparar (pots ajustar si vols)
   return digits.replace(/\s+/g,'').replace(/^\+34/, '34');
}

function normalizeName_(nom, cognoms) {
   const full = (String(nom || '') + ' ' + String(cognoms || '')).trim();
   if (!full) return '';
   if (typeof normalizeText_ === 'function') return normalizeText_(full);
   return full.toLowerCase().replace(/\s+/g,' ').trim();
}

function buildPersonLookup_(shP, pHeader, idxP) {
   const colId = idxP['id_persona'];
   if (colId === undefined) throw new Error('PERSONES necessita id_persona');

   const colEmail = idxP['email'];
   const colTel    = idxP['telefon_mobil'];
   const colNom    = idxP['nom'];
   const colC1      = idxP['cognom1'];
   const colC2      = idxP['cognom2'];

   const lastRow = shP.getLastRow();
   const emailMap = {};
   const phoneMap = {};
   const nameMap   = {};

   if (lastRow <= 1) return { emailMap, phoneMap, nameMap, rowById: {} };

   const values = shP.getRange(2,1,lastRow-1,shP.getLastColumn()).getValues();
   const rowById = {};

   for (let i = 0; i < values.length; i++) {
      const sheetRow = i + 2;
      const id = String(values[i][colId] || '').trim();
      if (!id) continue;

      rowById[id] = sheetRow;

      const email = (colEmail !== undefined) ? String(values[i][colEmail] || '').trim().toLowerCase() : '';
      const tel    = (colTel !== undefined) ? normalizePhone_(String(values[i][colTel] || '')) : '';
      const nom    = (colNom !== undefined) ? String(values[i][colNom] || '') : '';
      const c1      = (colC1 !== undefined) ? String(values[i][colC1] || '') : '';
      const c2      = (colC2 !== undefined) ? String(values[i][colC2] || '') : '';
      const name   = normalizeName_(nom, (c1 + ' ' + c2).trim());

      if (email) emailMap[email] = { id_persona: id, by: 'email' };

      if (tel) {
         if (phoneMap[tel] && phoneMap[tel].id_persona !== id) phoneMap[tel] = { ambiguous:true, reason:'phone' };
         else phoneMap[tel] = { id_persona: id, by: 'phone' };
      }

      if (name) {
         if (nameMap[name] && nameMap[name].id_persona !== id) nameMap[name] = { ambiguous:true, reason:'name' };
         else nameMap[name] = { id_persona: id, by: 'name' };
      }
   }

   return { emailMap, phoneMap, nameMap, rowById };
}

function findMatch_(lookup, extracted) {
   const email = String(extracted.email || '').trim().toLowerCase();
   const tel    = normalizePhone_(extracted.telefon || '');
   const name   = normalizeName_(extracted.nom || '', extracted.cognoms || '');

   if (email && lookup.emailMap[email]) return lookup.emailMap[email];
   if (tel && lookup.phoneMap[tel]) return lookup.phoneMap[tel];
   if (name && lookup.nameMap[name]) return lookup.nameMap[name];
   return null;
}

function updatePersonFromCv_(shP, pHeader, idxP, lookup, idPersona, extracted, matchBy) {
   const row = lookup.rowById[idPersona];
   if (!row) return { updated:false, notes:'Persona not found by id' };

   const values = shP.getRange(row,1,1,pHeader.length).getValues()[0];
   let changed = false;

   function setIfEmpty(key, val) {
      val = String(val || '').trim();
      if (!val) return false;
      const i = idxP[key];
      if (i === undefined) return false;
      const cur = String(values[i] || '').trim();
      if (cur) return false;
      values[i] = val;
      return true;
   }

   // B) omplir si buit
   changed = setIfEmpty('telefon_mobil', extracted.telefon) || changed;
   changed = setIfEmpty('doc_numero', extracted.doc) || changed;
   changed = setIfEmpty('municipi', extracted.municipi) || changed;
   changed = setIfEmpty('comarca', extracted.comarca) || changed;
   changed = setIfEmpty('nom', extracted.nom) || changed;
   changed = setIfEmpty('cognom1', extracted.cognoms) || changed;

   // email: només si NO hem fet match per email i està buit
   if (matchBy !== 'email') {
      changed = setIfEmpty('email', extracted.email) || changed;
   }

   // A) sempre (només si hi ha canvis)
   if (changed) {
      const now = (typeof toIso_ === 'function') ? toIso_(new Date()) : new Date().toISOString();
      if (idxP['updated_by_email'] !== undefined) values[idxP['updated_by_email']] = 'SYSTEM_CV';
      if (idxP['data_ultima_actualitzacio'] !== undefined) values[idxP['data_ultima_actualitzacio']] = now;
      shP.getRange(row,1,1,pHeader.length).setValues([values]);
      return { updated:true, notes:'Updated empty fields (match=' + matchBy + ')' };
   }

   return { updated:false, notes:'' };
}

function createPersonFromCv_(shP, pHeader, idxP, lookup, extracted) {
   const idPersona = (typeof uuid_ === 'function') ? ('P-' + uuid_()) : ('P-' + Utilities.getUuid());
   const now = (typeof toIso_ === 'function') ? toIso_(new Date()) : new Date().toISOString();

   const rowObj = {};
   rowObj['id_persona'] = idPersona;
   rowObj['emi_id_persona'] = '';
   rowObj['data_alta'] = now;
   rowObj['origen_alta'] = 'ALTRES';
   rowObj['estat_fitxa'] = 'ACTIU';

   rowObj['nom'] = extracted.nom || '';
   rowObj['cognom1'] = extracted.cognoms || '';
   rowObj['cognom2'] = '';

   rowObj['telefon_mobil'] = extracted.telefon || '';
   rowObj['email'] = (extracted.email || '').toLowerCase();

   rowObj['doc_tipus'] = '';
   rowObj['doc_numero'] = extracted.doc || '';

   rowObj['municipi'] = extracted.municipi || '';
   rowObj['comarca'] = extracted.comarca || '';

   rowObj['observacions'] = 'Creat automàticament des de CV';
   rowObj['created_by_email'] = 'SYSTEM_CV';
   rowObj['updated_by_email'] = 'SYSTEM_CV';
   rowObj['data_ultima_actualitzacio'] = now;

   const out = new Array(pHeader.length).fill('');
   for (let c = 0; c < pHeader.length; c++) {
      const k = pHeader[c];
      if (rowObj.hasOwnProperty(k)) out[c] = rowObj[k];
   }

   shP.appendRow(out);

   // refresca lookup mínim
   lookup.rowById[idPersona] = shP.getLastRow();
   const email = String(rowObj['email'] || '').trim();
   const tel = normalizePhone_(rowObj['telefon_mobil'] || '');
   const name = normalizeName_(rowObj['nom'] || '', rowObj['cognom1'] || '');

   if (email) lookup.emailMap[email] = { id_persona:idPersona, by:'email' };
   if (tel) lookup.phoneMap[tel] = { id_persona:idPersona, by:'phone' };
   if (name) lookup.nameMap[name] = { id_persona:idPersona, by:'name' };

   return { id_persona:idPersona };
}