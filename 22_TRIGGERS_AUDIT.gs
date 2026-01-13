/*******************************************************
  * 22_TRIGGERS_AUDIT.gs
  * - Llistar triggers instal·lats
  * - Esborrar triggers per nom de funció (handler)
  * - "Tallafocs" per parar bucles
  *******************************************************/

function EMI_TRIGGERS_listNow() {
   const tr = ScriptApp.getProjectTriggers();
   const rows = tr.map(t => {
      return {
         handler: t.getHandlerFunction(),
         type: String(t.getEventType()),
         source: String(t.getTriggerSource()),
         id: t.getUniqueId()
      };
   });

   try { if (typeof log_ === 'function') log_('INFO','TRIGGERS_LIST','Triggers instal·lats', { count: rows.length, triggers: rows }); } catch(e) {}
   Logger.log(JSON.stringify(rows, null, 2));

   // Toast
   try { SpreadsheetApp.getActiveSpreadsheet().toast('Triggers llistats. Mira el LOG/Executions ✅', 'EMI', 6); } catch(e2) {}

   return rows;
}

function EMI_TRIGGERS_removeByHandler(handlerName) {
   const tr = ScriptApp.getProjectTriggers();
   let removed = 0;

   tr.forEach(t => {
      if (t.getHandlerFunction() === handlerName) {
         ScriptApp.deleteTrigger(t);
         removed++;
      }
   });

   try { if (typeof log_ === 'function') log_('WARN','TRIGGERS_REMOVED','Triggers eliminats per handler', { handlerName, removed }); } catch(e) {}
   try { SpreadsheetApp.getActiveSpreadsheet().toast(`Eliminats ${removed} triggers de ${handlerName}`, 'EMI', 6); } catch(e2) {}

   return removed;
}

function EMI_TRIGGERS_removeCommonLoopersNow() {
   // Els típics que poden estar provocant el bucle
   const suspects = [
      'EMI_recomputeMatchesBothSidesNow',
      'EMI_PIPELINE_rebuildBrainNow',
      'EMI_reindexAllNow',
      'EMI_recomputeVacantMatchesNow',
      'EMI_recomputeVacantMatches',
      'EMI_recomputeMatchesNow',
      'EMI_cron_5m',
      'EMI_cron5m_V2',
      'EMI_cron_5m_V2'
   ];

   const out = {};
   suspects.forEach(name => out[name] = EMI_TRIGGERS_removeByHandler(name));

   try { if (typeof log_ === 'function') log_('WARN','TRIGGERS_PURGE_COMMON','Purge triggers comuns', out); } catch(e) {}
   return out;
}