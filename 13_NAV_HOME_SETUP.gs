/*******************************************************
  * 13_NAV_HOME_SETUP.gs   (V2)
  * - Crea/actualitza la pestanya NAV_HOME al MASTER
  * - Seed inicial (si està buit)
  * - Fix targets perquè coincideixin amb els noms reals de vistes AppSheet
  *******************************************************/

const EMI_NAV_HOME_SHEET = 'NAV_HOME';
const EMI_NAV_HOME_HEADERS = [
   'id_nav','ordre','titol','subtitol','emoji','icon_url','target_type','target_name','actiu'
];

// Seed base (si està buit). Els target_name es fixen després amb EMI_navHomeFixTargetsForAppSheet()
const EMI_NAV_HOME_SEED = [
   ['HOME_01', 10, 'Persones', 'Consulta i edita perfils', '👤', '', 'VIEW', '', true],
   ['HOME_02', 20, 'Empreses', 'Fitxes i seguiment', '🏢', '', 'VIEW', '', true],
   ['HOME_03', 30, 'Vacants', 'Gestió de vacants', '📌', '', 'VIEW', '', true],
   ['HOME_04', 40, 'Nova persona', 'Alta ràpida', '➕', '', 'FORM', '', true],
   ['HOME_05', 50, 'Nova empresa', 'Alta ràpida', '➕', '', 'FORM', '', true],
   ['HOME_06', 60, 'Nova vacant', 'Alta ràpida', '➕', '', 'FORM', '', true],
];

function EMI_setupNavHome() {
   const runner = () => {
      const ss = (typeof getMasterSs_ === 'function') ? getMasterSs_() : SpreadsheetApp.getActiveSpreadsheet();
      let sh = ss.getSheetByName(EMI_NAV_HOME_SHEET);
      if (!sh) sh = ss.insertSheet(EMI_NAV_HOME_SHEET);

      // Headers
      const lastCol = Math.max(sh.getLastColumn(), EMI_NAV_HOME_HEADERS.length);
      const current = (sh.getLastRow() >= 1)
         ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim())
         : [];
      const emptyish = current.join('').trim() === '';

      if (sh.getLastRow() === 0 || emptyish) {
         sh.getRange(1, 1, 1, EMI_NAV_HOME_HEADERS.length).setValues([EMI_NAV_HOME_HEADERS]);
      } else {
         const set = {};
         current.forEach(h => { if (h) set[h] = true; });
         const toAdd = EMI_NAV_HOME_HEADERS.filter(h => !set[h]);
         if (toAdd.length) {
            const start = sh.getLastColumn() + 1;
            sh.insertColumnsAfter(sh.getLastColumn(), toAdd.length);
            sh.getRange(1, start, 1, toAdd.length).setValues([toAdd]);
         }
      }

      // Seed si buit
      if (sh.getLastRow() <= 1) {
         sh.getRange(2, 1, EMI_NAV_HOME_SEED.length, EMI_NAV_HOME_HEADERS.length).setValues(EMI_NAV_HOME_SEED);
      }

      try { sh.setFrozenRows(1); } catch (_) {}

      // Fix targets perquè coincideixin amb els noms de vistes que tu tens (MAJÚSCULES)
      EMI_navHomeFixTargetsForAppSheet_();

      try { if (typeof log_ === 'function') log_('INFO','NAV_HOME_OK','NAV_HOME creat/actualitzat + targets fixats', {}); } catch (_) {}
      try { if (typeof uiAlert_ === 'function') uiAlert_('NAV_HOME creat/actualitzat ✅ (targets AppSheet fixats)'); } catch (_) {}
      return true;
   };

   if (typeof withLock_ === 'function') return withLock_('NAV_HOME_SETUP', runner);
   return runner();
}

/**
  * Ajusta target_name (i target_type) segons els noms reals de vistes AppSheet:
  *   - VIEW: PERSONES / EMPRESES / VACANTS
  *   - FORM: PERSONES FORM / EMPRESES FORM / VACANTS FORM
  */
function EMI_navHomeFixTargetsForAppSheet_() {
   const ss = (typeof getMasterSs_ === 'function') ? getMasterSs_() : SpreadsheetApp.getActiveSpreadsheet();
   const sh = ss.getSheetByName(EMI_NAV_HOME_SHEET);
   if (!sh) throw new Error('No existeix NAV_HOME');

   const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(v => String(v || '').trim());
   const idx = {};
   header.forEach((h,i) => { if (h) idx[h] = i + 1; });

   const colId = idx['id_nav'];
   const colType = idx['target_type'];
   const colName = idx['target_name'];
   if (!colId || !colType || !colName) throw new Error('NAV_HOME necessita id_nav, target_type, target_name');

   const fixes = {
      'HOME_01': { type:'VIEW', name:'PERSONES' },
      'HOME_02': { type:'VIEW', name:'EMPRESES' },
      'HOME_03': { type:'VIEW', name:'VACANTS' },
      'HOME_04': { type:'VIEW', name:'PERSONES FORM' },
      'HOME_05': { type:'VIEW', name:'EMPRESES FORM' },
      'HOME_06': { type:'VIEW', name:'VACANTS FORM' },
   };

   const lastRow = sh.getLastRow();
   if (lastRow <= 1) return;

   const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
   let changed = 0;

   for (let r = 0; r < data.length; r++) {
      const id = String(data[r][colId - 1] || '').trim();
      if (!id || !fixes[id]) continue;

      const want = fixes[id];
      const curType = String(data[r][colType - 1] || '').trim();
      const curName = String(data[r][colName - 1] || '').trim();

      if (curType !== want.type) { data[r][colType - 1] = want.type; changed++; }
      if (curName !== want.name) { data[r][colName - 1] = want.name; changed++; }
   }

   if (changed) {
      sh.getRange(2, 1, data.length, sh.getLastColumn()).setValues(data);
   }
}