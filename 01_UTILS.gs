/*******************************************************
  * 01_UTILS.gs — EMI BD MASTER (AppSheet) · Factory V2
  * - helpers, logging, ui-safe, locks
  *******************************************************/

function props_() {
   return PropertiesService.getScriptProperties();
}
function getProp_(k) {
   return props_().getProperty(k);
}
function setProp_(k, v) {
   props_().setProperty(k, String(v));
}
function delProp_(k) {
   props_().deleteProperty(k);
}

function toIso_(d) {
   return Utilities.formatDate(d instanceof Date ? d : new Date(d), EMI_TZ, "yyyy-MM-dd HH:mm:ss");
}
function safeJson_(obj) {
   try { return obj ? JSON.stringify(obj) : ''; } catch (e) { return ''; }
}
function uuid_() {
   return Utilities.getUuid();
}

function normalizeText_(s) {
   return String(s || '')
      .trim()
      .toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/\s+/g, ' ');
}

function withLock_(name, fn) {
   const lock = LockService.getScriptLock();
   if (!lock.tryLock(30000)) throw new Error('No s’ha pogut adquirir lock: ' + name);
   try { return fn(); }
   finally { try { lock.releaseLock(); } catch (_) {} }
}

function getMasterSs_() {
   let ss = null;
   try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (_) { ss = null; }

   if (ss) {
      setProp_(PROP_KEYS.MASTER_SS_ID, ss.getId());
      return ss;
   }

   const id = getProp_(PROP_KEYS.MASTER_SS_ID) || (typeof EMI_MASTER_SS_ID_DEFAULT !== 'undefined' ? EMI_MASTER_SS_ID_DEFAULT : '');
   if (id) return SpreadsheetApp.openById(id);

   throw new Error("No hi ha Spreadsheet actiu. Obre el MASTER i executa des del Sheets.");
}

function getSheet_(name) {
   const ss = getMasterSs_();
   const sh = ss.getSheetByName(name);
   if (!sh) throw new Error("No s'ha trobat la pestanya: " + name);
   return sh;
}

function log_(level, action, message, obj) {
   try {
      const sh = getSheet_(SHEETS.LOG);
      sh.appendRow([toIso_(new Date()), level || 'INFO', action || '', message || '', safeJson_(obj)]);
   } catch (e) {
      // últim recurs si encara no existeix LOG
      console.log(level, action, message, obj, e);
   }
}

/** UI helpers (safe) **/
function ui_() {
   try { return SpreadsheetApp.getUi(); } catch (_) { return null; }
}

function uiAlert_(msg) {
   const uiDisabled = (typeof CFG !== 'undefined' && CFG && CFG.UI_ALERTS === false);
   if (uiDisabled) {
      try { log_('INFO', 'UI_ALERT_SKIPPED', String(msg || ''), {}); } catch (_) {}
      return;
   }
   const ui = ui_();
   if (ui) ui.alert(String(msg || ''));
}

function toastSafe_(ss, message, title, seconds) {
   const toastDisabled = (typeof CFG !== 'undefined' && CFG && CFG.UI_TOASTS === false);
   if (toastDisabled) return;
   try {
      if (ss && typeof ss.toast === 'function') {
         ss.toast(String(message || ''), String(title || ''), Number(seconds || 5));
      }
   } catch (e) {
      try { log_('WARN', 'TOAST_FAIL', 'No s’ha pogut fer toast', { err: String(e) }); } catch (_) {}
   }
}

/** Headers helpers **/
function headerMap_(sh) {
   const lastCol = sh.getLastColumn();
   const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
   const m = {};
   hdr.forEach((h, i) => { if (h) m[h] = i + 1; });
   return m;
}

function setCell_(sh, row, col, value) {
   sh.getRange(row, col, 1, 1).setValue(value);
}

/** Dates / seq helpers **/
function yearNow_() {
   return Utilities.formatDate(new Date(), EMI_TZ, 'yyyy');
}

function padLeft_(n, width) {
   const s = String(n);
   if (s.length >= width) return s;
   return new Array(width - s.length + 1).join('0') + s;
}

/** Debug **/
function EMI_debugUtils() {
   const ss = getMasterSs_();
   toastSafe_(ss, 'toastSafe_ OK', 'EMI DEBUG', 3);
   log_('INFO', 'DEBUG_UTILS', 'utils OK', { EMI_BD_VERSION: (typeof EMI_BD_VERSION !== 'undefined' ? EMI_BD_VERSION : 'MISSING') });
   uiAlert_('Utils OK ✅ (si tens UI_ALERTS=false, potser no veuràs aquest missatge i és normal).');
}
