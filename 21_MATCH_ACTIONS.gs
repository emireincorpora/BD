/*******************************************************
  * 21_MATCH_ACTIONS.gs
  * - Botons/accions visibles per recalcular matchs
  *******************************************************/

function EMI_recomputeMatchesBothSidesNow(topN) {
   topN = Number(topN || 25);

   // Fem servir el lock si existeix (millor) i si no, lock simple
   const runner = () => {
      const out = { topN, ran: [], missing: [] };

      // 1) Vacants -> persones
      if (typeof EMI_recomputeVacantMatchesNow === 'function') {
         EMI_recomputeVacantMatchesNow(topN);
         out.ran.push('EMI_recomputeVacantMatchesNow');
      } else {
         out.missing.push('EMI_recomputeVacantMatchesNow');
      }

      // 2) Persones -> vacants (provem diversos noms possibles)
      const candidates = [
         'EMI_recomputePersonaMatchesNow',
         'EMI_recomputePersonMatchesNow',
         'EMI_recomputePeopleMatchesNow'
      ];

      let ok = false;
      for (let i = 0; i < candidates.length; i++) {
         const fnName = candidates[i];
         if (typeof globalThis[fnName] === 'function') {
            globalThis[fnName](topN);
            out.ran.push(fnName);
            ok = true;
            break;
         }
      }
      if (!ok) out.missing.push('PersonaMatchesFn');

      // Si no hem pogut fer els 2, fem pipeline complet (si existeix)
      if (out.ran.length < 2 && typeof EMI_PIPELINE_rebuildBrainNow === 'function') {
         EMI_PIPELINE_rebuildBrainNow(topN);
         out.ran.push('EMI_PIPELINE_rebuildBrainNow (fallback)');
      }

      try { if (typeof log_ === 'function') log_('INFO','MATCH_BOTHSIDES_UI','Recompute matches both sides', out); } catch (e) {}
      try { SpreadsheetApp.getActiveSpreadsheet().toast('Match recalculat ✅', 'EMI', 5); } catch (e2) {}

      return out;
   };

   if (typeof withLock_ === 'function') return withLock_('MATCH_BOTHSIDES_UI', runner);

   const lock = LockService.getScriptLock();
   lock.waitLock(30000);
   try { return runner(); } finally { try { lock.releaseLock(); } catch (e3) {} }
}