/*******************************************************
 * 23_MASTER_CLEANUP.gs
 * - Neteja "columnes fantasma" (sense capçalera) al final dels fulls
 * - Ajuda AppSheet i fa el MASTER més net
 *******************************************************/

function EMI_MASTER_trimGhostColumnsNow() {
  return withLock_('MASTER_TRIM_GHOST', function () {
    var ss = getMasterSs_();
    var targets = [
      SHEETS.PERSONES,
      SHEETS.EMPRESES,
      SHEETS.VACANTS,
      'CV_RAW',
      'CATALEGS',
      'SINONIMS'
    ];

    var out = [];
    targets.forEach(function (sheetName) {
      var sh = ss.getSheetByName(sheetName);
      if (!sh) return;

      var maxCols = sh.getMaxColumns();
      var header = sh.getRange(1, 1, 1, maxCols).getValues()[0];

      var lastHeaderCol = 0;
      for (var c = header.length; c >= 1; c--) {
        var v = header[c - 1];
        if (v !== null && v !== undefined && String(v).trim() !== '') { lastHeaderCol = c; break; }
      }
      if (!lastHeaderCol) return;

      var toDelete = maxCols - lastHeaderCol;
      if (toDelete > 0) {
        sh.deleteColumns(lastHeaderCol + 1, toDelete);
      }

      out.push({ sheet: sheetName, maxCols: maxCols, lastHeaderCol: lastHeaderCol, deleted: toDelete });
    });

    log_('INFO', 'MASTER_TRIM_GHOST_OK', 'Columnes fantasma retallades', { sheets: out });
    toastSafe_('Neteja columnes fantasma ✅');
    return out;
  });
}
