/*******************************************************
  * 20_COMPAT_HELPERS.gs
  * - Funcions petites de compatibilitat que altres fitxers poden necessitar
  *******************************************************/

/**
  * Retorna TRUE si el valor significa “sí”.
  * Accepta: SI / SÍ / YES / TRUE / 1 / (booleà true)
  */
function isYes_(v) {
   if (v === true) return true;
   if (v === false) return false;

   const s = String(v === null || v === undefined ? '' : v).trim().toUpperCase();
   return (s === 'SI' || s === 'SÍ' || s === 'YES' || s === 'TRUE' || s === '1');
}