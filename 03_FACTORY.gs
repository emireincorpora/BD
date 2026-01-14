/*******************************************************
  * 03_FACTORY.gs — factory principal (crear/upgrade master)
  *******************************************************/

function EMI_FACTORY_V2() {
   withLock_('FACTORY_V2', () => {
      const ss = getMasterSs_();
      setProp_(PROP_KEYS.VERSION, EMI_BD_VERSION);

      toastSafe_(ss, 'Creant/actualitzant MASTER…', 'EMI Factory V2', 8);

      ensureSheetsAndHeaders_(ss);
      ensureCriticalHeadersAndAliases_(ss);

      // Drive root + carpetes
      const root = getOrCreateRootFolder_();
      try { moveFileToFolder_(ss.getId(), root); } catch (_) {}

      const attach = getOrCreateChildFolder_(root, CFG.SUBFOLDER_ATTACHMENTS, PROP_KEYS.ATTACH_FOLDER_ID);
      const exp = getOrCreateChildFolder_(root, CFG.SUBFOLDER_EXPORTS, PROP_KEYS.EXPORTS_FOLDER_ID);

      // Seeds
      ensureDefaultTecnics_();
      ensureCatalogDefaults_();
      ensureFormacionsDefaults_();

      // Sinònims (auto des de CATALEGS.etiqueta)
      syncSinonimsFromCatalogs_();

      // Visibilitat
      if (CFG.HIDE_SUPPORT_SHEETS_BY_DEFAULT) EMI_hideSupportSheets();

      SpreadsheetApp.flush();

      log_('INFO','FACTORY_V2_OK','Factory V2 completada', {
         masterId: ss.getId(),
         rootFolderId: root.getId(),
         attachmentsFolderId: attach.getId(),
         exportsFolderId: exp.getId()
      });

      toastSafe_(ss, 'Factory V2 complet ✅', 'EMI Factory V2', 5);
      uiAlert_('Factory V2 complet ✅\n\nRevisa LOG i HEALTHCHECK.');
   });
}

/******** Drive helpers ********/

function getOrCreateRootFolder_() {
   const stored = getProp_(PROP_KEYS.ROOT_FOLDER_ID);
   if (stored) {
      try { return DriveApp.getFolderById(stored); } catch (_) {}
   }
   const root = DriveApp.getRootFolder();
   const it = root.getFoldersByName(CFG.ROOT_FOLDER_NAME);
   const folder = it.hasNext() ? it.next() : root.createFolder(CFG.ROOT_FOLDER_NAME);
   setProp_(PROP_KEYS.ROOT_FOLDER_ID, folder.getId());
   return folder;
}

function getOrCreateChildFolder_(parent, name, propKey) {
   const stored = getProp_(propKey);
   if (stored) {
      try { return DriveApp.getFolderById(stored); } catch (_) {}
   }
   const it = parent.getFoldersByName(name);
   const folder = it.hasNext() ? it.next() : parent.createFolder(name);
   setProp_(propKey, folder.getId());
   return folder;
}

function moveFileToFolder_(fileId, folder) {
   const file = DriveApp.getFileById(fileId);
   try { file.moveTo(folder); } catch (_) {
      // en alguns casos moveTo falla: fem add/remove
      try { folder.addFile(file); } catch (_) {}
      try {
         const parents = file.getParents();
         while (parents.hasNext()) {
            const p = parents.next();
            if (p.getId() !== folder.getId()) {
               try { p.removeFile(file); } catch (_) {}
            }
         }
      } catch (_) {}
   }
}

/**
 * Ensure critical columns exist (via HEADERS) and backfill alias columns so AppSheet and scripts stay consistent.
 * - PERSONES: backfill 'cv_sector_objectiu' from legacy 'sector_objectiu' if needed
 * - VACANTS: keeps 'id_empresa' and 'empresa_id' synchronized (backfill missing values in either direction)
 */
function ensureCriticalHeadersAndAliases_(ss) {
   try {
      // PERSONES legacy alias: sector_objectiu -> cv_sector_objectiu
      const shP = ss.getSheetByName(SHEETS.PERSONES);
      if (shP && shP.getLastRow() >= 2) {
         const hmP = headerMap_(shP);
         const colLegacy = hmP['sector_objectiu'];
         const colCv = hmP['cv_sector_objectiu'];
         if (colLegacy && colCv) {
            const n = shP.getLastRow() - 1;
            const legacy = shP.getRange(2, colLegacy, n, 1).getValues();
            const cv = shP.getRange(2, colCv, n, 1).getValues();
            let changed = 0;
            for (let i = 0; i < n; i++) {
               const legacyVal = String(legacy[i][0] || '').trim();
               const cvVal = String(cv[i][0] || '').trim();
               if (!cvVal && legacyVal) {
                  cv[i][0] = legacyVal;
                  changed++;
               }
            }
            if (changed) {
               shP.getRange(2, colCv, n, 1).setValues(cv);
               log_('INFO', 'ALIAS_BACKFILL_OK', 'Backfill alias columns (PERSONES)', { changed, from: 'sector_objectiu', to: 'cv_sector_objectiu' });
            }
         }
      }

      // VACANTS alias sync: id_empresa <-> empresa_id
      const sh = ss.getSheetByName(SHEETS.VACANTS);
      if (sh && sh.getLastRow() >= 2) {
         const hm = headerMap_(sh);
         const colA = hm['id_empresa'];
         const colB = hm['empresa_id'];
         if (colA && colB) {
            const n = sh.getLastRow() - 1;
            const a = sh.getRange(2, colA, n, 1).getValues();
            const b = sh.getRange(2, colB, n, 1).getValues();
            let changed = 0;
            for (let i = 0; i < n; i++) {
               const av = String(a[i][0] || '').trim();
               const bv = String(b[i][0] || '').trim();
               if (!bv && av) { b[i][0] = av; changed++; }
               else if (!av && bv) { a[i][0] = bv; changed++; }
            }
            if (changed) {
               sh.getRange(2, colA, n, 1).setValues(a);
               sh.getRange(2, colB, n, 1).setValues(b);
               log_('INFO', 'ALIAS_BACKFILL_OK', 'Backfill alias columns (VACANTS)', { changed, colA: 'id_empresa', colB: 'empresa_id' });
            }
         }
      }
   } catch (e) {
      log_('WARN', 'ALIAS_BACKFILL_FAIL', 'No s’ha pogut sincronitzar alias columns', { err: String(e) });
   }
}
