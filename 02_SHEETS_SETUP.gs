/*******************************************************
  * 02_SHEETS_SETUP.gs — create/upgrade sheets + headers + visibility
  *******************************************************/

function ensureSheetsAndHeaders_(ss) {
   const wantedNames = Object.keys(HEADERS);

   wantedNames.forEach(name => {
      const headers = HEADERS[name];
      let sh = ss.getSheetByName(name);
      if (!sh) sh = ss.insertSheet(name);

      if (sh.getLastRow() === 0) {
         sh.appendRow(headers);
         return;
      }

      const current = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      const emptyish = current.join('').trim() === '';
      if (emptyish) {
         sh.getRange(1, 1, 1, headers.length).setValues([headers]);
         return;
      }

      const set = {};
      current.forEach(h => { if (h) set[h] = true; });
      const toAdd = headers.filter(h => !set[h]);

      if (toAdd.length) {
         const startCol = sh.getLastColumn() + 1;
         sh.insertColumnsAfter(sh.getLastColumn(), toAdd.length);
         sh.getRange(1, startCol, 1, toAdd.length).setValues([toAdd]);
      }
   });

   // No esborrem pestanyes desconegudes per defecte (CFG ho controla)
   if (CFG.CLEAN_DELETE_UNKNOWN_SHEETS) {
      deleteUnknownSheets_(ss, wantedNames);
   }

   // Ordena una mica per llegibilitat
   reorderSheets_(ss);
}

function deleteUnknownSheets_(ss, whitelist) {
   const keep = {};
   whitelist.forEach(n => keep[n] = true);

   const all = ss.getSheets();
   const active = ss.getActiveSheet();
   const activeName = active ? active.getName() : '';

   all.forEach(sh => {
      const name = sh.getName();
      if (keep[name]) return;
      if (name === activeName) return;
      try {
         ss.deleteSheet(sh);
         log_('INFO','DELETE_SHEET','Esborrada pestanya no oficial', { name });
      } catch (e) {
         log_('WARN','DELETE_SHEET_FAIL','No s’ha pogut esborrar pestanya', { name, err: String(e) });
      }
   });
}

function reorderSheets_(ss) {
   const order = []
      .concat(UI_VISIBLE_SHEETS)
      .concat([SHEETS.TECNICS, SHEETS.CATALEGS, SHEETS.SINONIMS, SHEETS.FORMACIONS_CATALOG])
      .concat([SHEETS.DOCUMENTS, SHEETS.SEGUIMENTS_PERSONA, SHEETS.SEGUIMENTS_EMPRESA])
      .concat([SHEETS.LOG, SHEETS.HEALTH]);

   let pos = 1;
   order.forEach(name => {
      const sh = ss.getSheetByName(name);
      if (!sh) return;
      try { sh.setIndex(pos++); } catch (_) {}
   });
}

function EMI_showAllSheets() {
   const ss = getMasterSs_();
   ss.getSheets().forEach(sh => { try { sh.showSheet(); } catch (_) {} });
   SpreadsheetApp.flush();
   log_('INFO','SHOW_ALL_SHEETS','Mostrades totes les pestanyes', {});
}

function EMI_hideSupportSheets() {
   const ss = getMasterSs_();

   // assegura activa una de visible abans d’amagar
   const firstVisibleName = UI_VISIBLE_SHEETS.find(n => ss.getSheetByName(n));
   if (firstVisibleName) ss.setActiveSheet(ss.getSheetByName(firstVisibleName));

   // mostra visibles
   UI_VISIBLE_SHEETS.forEach(name => {
      const sh = ss.getSheetByName(name);
      if (!sh) return;
      try { sh.showSheet(); } catch (_) {}
   });

   if (!CFG.HIDE_SUPPORT_SHEETS_BY_DEFAULT) return;

   // amaga suport
   SUPPORT_SHEETS.forEach(name => {
      const sh = ss.getSheetByName(name);
      if (!sh) return;
      try { sh.hideSheet(); } catch (e) {
         log_('WARN','HIDE_FAIL','No s’ha pogut amagar pestanya', { name, err: String(e) });
      }
   });

   SpreadsheetApp.flush();
   log_('INFO','HIDE_SUPPORT_SHEETS','Amagades pestanyes de suport', {});
}