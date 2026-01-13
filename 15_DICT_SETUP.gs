/*******************************************************
  * 15_DICT_SETUP.gs — Diccionari (setup + FPCAT seed) · V1.1
  * - CORRECCIÓ: eliminat placeholder {{...}} que provocava SyntaxError
  * - Compatible amb el teu format actual de CATALEGS/SINONIMS (actiu = 'SI')
  *******************************************************/

function EMI_DICT_SETUP_V1() {
   return EMI_DICT_withLock_('DICT_SETUP', function () {
      var ss = EMI_DICT_getMaster_();

      // 1) Carpetes base (sense Drive API avançada)
      EMI_DICT_ensureBaseFolders_();

      // 2) Assegura fulls
      EMI_DICT_ensureSheet_(ss, EMI_DICT_sheetName_('CATALEGS', 'CATALEGS'));
      EMI_DICT_ensureSheet_(ss, EMI_DICT_sheetName_('SINONIMS', 'SINONIMS'));
      EMI_DICT_ensureCustomDictSheets_(ss);

      // 3) Columnes que vols (manual + auto separats)
      EMI_DICT_addColumnIfMissing_(EMI_DICT_sheetName_('PERSONES', 'PERSONES'), 'tags_persona_auto');
      EMI_DICT_addColumnIfMissing_(EMI_DICT_sheetName_('VACANTS', 'VACANTS'), 'tags_vacant_auto');

      // 4) Seed FPCAT (famílies + àmbits)
      EMI_DICT_seedFpcatFamiliesAmbits_();

      // 5) Informe cobertura (opcional, però útil)
      EMI_DICT_updateCoverage_();

      EMI_DICT_toast_('Diccionari: setup OK');
      return true;
   });
}

/** =========================
  *   Helpers bàsics
  *   ========================= */

function EMI_DICT_getMaster_() {
   if (typeof getMasterSs_ === 'function') return getMasterSs_();
   return SpreadsheetApp.getActiveSpreadsheet();
}

function EMI_DICT_sheetName_(key, fallback) {
   try {
      if (typeof SHEETS !== 'undefined' && SHEETS && SHEETS[key]) return SHEETS[key];
   } catch (e) {}
   return fallback;
}

function EMI_DICT_withLock_(name, fn) {
   if (typeof withLock_ === 'function') return withLock_(name, fn);
   var lock = LockService.getScriptLock();
   lock.waitLock(30000);
   try { return fn(); }
   finally { try { lock.releaseLock(); } catch (e) {} }
}

function EMI_DICT_toast_(msg) {
   try {
      var ss = EMI_DICT_getMaster_();
      if (ss && typeof ss.toast === 'function') ss.toast(String(msg || ''), 'EMI', 5);
   } catch (e) {}
}

function EMI_DICT_log_(level, action, message, obj) {
   try {
      if (typeof log_ === 'function') return log_(level, action, message, obj || {});
   } catch (e) {}
   try { console.log(level, action, message, obj || {}); } catch (e2) {}
}

function EMI_DICT_norm_(s) {
   s = String(s || '').toLowerCase().trim();
   try {
      s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
   } catch (e) {}
   s = s.replace(/\s+/g, ' ');
   return s;
}

function EMI_DICT_ensureSheet_(ss, name) {
   var sh = ss.getSheetByName(name);
   if (!sh) sh = ss.insertSheet(name);
   return sh;
}

function EMI_DICT_ensureHeaderRow_(sh, headers) {
   var lastCol = sh.getLastColumn();
   if (lastCol < headers.length) lastCol = headers.length;
   var existing = (sh.getLastRow() >= 1 && lastCol >= 1)
      ? sh.getRange(1, 1, 1, lastCol).getValues()[0]
      : [];
   var ok = true;
   for (var i = 0; i < headers.length; i++) {
      if (String(existing[i] || '').trim() !== headers[i]) { ok = false; break; }
   }
   if (!ok) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function EMI_DICT_ensureCustomDictSheets_(ss) {
   var p = EMI_DICT_ensureSheet_(ss, 'DICT_PENDENTS');
   EMI_DICT_ensureHeaderRow_(p, ['tipus','text_original','text_normalitzat','proposta_codi_tag','proposta_tag','score','font','created_at','notes']);

   var c = EMI_DICT_ensureSheet_(ss, 'DICT_COVERAGE');
   EMI_DICT_ensureHeaderRow_(c, ['context_tipus','context_codi','context_nom','num_tags','num_sinonims','updated_at']);
}

function EMI_DICT_addColumnIfMissing_(sheetName, colName) {
   var ss = EMI_DICT_getMaster_();
   var sh = ss.getSheetByName(sheetName);
   if (!sh) return;

   var lastCol = Math.max(1, sh.getLastColumn());
   var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
   var target = EMI_DICT_norm_(colName);

   for (var i = 0; i < headers.length; i++) {
      if (EMI_DICT_norm_(headers[i]) === target) return;
   }

   sh.insertColumnAfter(sh.getLastColumn());
   sh.getRange(1, sh.getLastColumn()).setValue(colName);
}

/** =========================
  *   Carpetes (sense Drive API avançada)
  *   ========================= */
function EMI_DICT_ensureBaseFolders_() {
   try {
      var rootName = (typeof CFG !== 'undefined' && CFG && CFG.ROOT_FOLDER_NAME) ? CFG.ROOT_FOLDER_NAME : 'EMI_BD_ROOT';
      var root = null;

      if (typeof getOrCreateRootFolder_ === 'function') {
         root = getOrCreateRootFolder_();
      } else {
         var it = DriveApp.getFoldersByName(rootName);
         root = it.hasNext() ? it.next() : DriveApp.createFolder(rootName);
      }

      var dictName = '_DICT';
      var dit = root.getFoldersByName(dictName);
      if (!dit.hasNext()) root.createFolder(dictName);

   } catch (err) {
      EMI_DICT_log_('WARN', 'DICT_FOLDERS', 'No he pogut assegurar carpetes (no és crític)', { err: String(err) });
   }
}

/** =========================
  *   CATALEGS CRUD (format del teu projecte)
  *   Columns: tipus | codi | valor | etiqueta | actiu | ordre
  *   actiu: 'SI' / 'NO'
  *   ========================= */

function EMI_DICT_getCatalegsSheet_() {
   var ss = EMI_DICT_getMaster_();
   return ss.getSheetByName(EMI_DICT_sheetName_('CATALEGS', 'CATALEGS'));
}

function EMI_DICT_catalegIndex_() {
   var sh = EMI_DICT_getCatalegsSheet_();
   if (!sh) return { sh: null, idx: {}, headers: [] };
   var lastRow = sh.getLastRow();
   var lastCol = Math.max(6, sh.getLastColumn());

   EMI_DICT_ensureHeaderRow_(sh, ['tipus','codi','valor','etiqueta','actiu','ordre']);

   var data = (lastRow >= 2) ? sh.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
   var idx = {};
   for (var i = 0; i < data.length; i++) {
      var tipus = String(data[i][0] || '').trim();
      var codi = String(data[i][1] || '').trim();
      if (!tipus || !codi) continue;
      idx[tipus + '||' + codi] = 2 + i;
   }
   return { sh: sh, idx: idx, lastCol: lastCol };
}

function EMI_DICT_upsertCataleg_(tipus, codi, valor, etiqueta, ordre) {
   var pack = EMI_DICT_catalegIndex_();
   if (!pack.sh) return;

   var key = tipus + '||' + codi;
   var row = pack.idx[key];

   var out = [tipus, codi, valor, etiqueta || '', 'SI', (ordre === undefined || ordre === null) ? '' : ordre];

   if (row) {
      pack.sh.getRange(row, 1, 1, 6).setValues([out]);
   } else {
      pack.sh.getRange(pack.sh.getLastRow() + 1, 1, 1, 6).setValues([out]);
   }
}

/** =========================
  *   Seed FPCAT
  *   ========================= */

function EMI_DICT_seedFpcatFamiliesAmbits_() {
   // Taula extreta del teu PDF (taula-ambits.pdf)
   var FPCAT = [["AA","Arts i artesanies","Reproduccions de motlles i peces ceràmiques artesanals"],["AA","Arts i artesanies","Terrisseria artesanal"],["AE","Activitats físiques i esportives","Accés i conservació en instal·lacions esportives"],["AE","Activitats físiques i esportives","Animació físic-esportiva i recreativa per a persones amb discapacitat"],["AE","Activitats físiques i esportives","Condicionament físic"],["AE","Activitats físiques i esportives","Ensenyament i animació socioesportiva"],["AE","Activitats físiques i esportives","Guia en el medi natural i de temps de lleure"],["AE","Activitats físiques i esportives","Instrucció en ioga"],["AE","Activitats físiques i esportives","Socorrisme en instal·lacions aquàtiques"],["AF","Arts gràfiques","Disseny i edició de publicacions impreses i multimèdia"],["AF","Arts gràfiques","Disseny i gestió de la producció gràfica"],["AF","Arts gràfiques","Impressió gràfica"],["AF","Arts gràfiques","Postimpressió i acabats gràfics"],["AF","Arts gràfiques","Preimpressió digital"],["AG","Administració i gestió","Administració i finances"],["AG","Administració i gestió","Assistència a la direcció"],["AG","Administració i gestió","Assistència en la gestió dels procediments tributaris"],["AG","Administració i gestió","Creació i gestió de microempreses"],["AG","Administració i gestió","Gestió administrativa"],["AG","Administració i gestió","Serveis administratius"],["AR","Agrària","Activitats auxiliars en agricultura"],["AR","Agrària","Activitats auxiliars en conservació i millora de forests"],["AR","Agrària","Activitats auxiliars en floristeria"],["AR","Agrària","Activitats auxiliars en vivers, jardins i centres de jardineria"],["AR","Agrària","Activitats eqüestres"],["AR","Agrària","Aprofitament i conservació del medi natural"],["AR","Agrària","Gestió forestal i del medi natural"],["AR","Agrària","Jardineria i floristeria"],["AR","Agrària","Producció agroecològica"],["AR","Agrària","Producció agropecuària"],["AR","Agrària","Ramaderia i assistència en sanitat animal"],["AR","Agrària","Paisatgisme i medi rural"],["CM","Comerç i màrqueting","Activitats de venda"],["CM","Comerç i màrqueting","Comercialització de productes i serveis"],["CM","Comerç i màrqueting","Gestió comercial i financera del transport per carretera"],["CM","Comerç i màrqueting","Gestió de comerç internacional"],["CM","Comerç i màrqueting","Gestió de màrqueting i comunicació"],["CM","Comerç i màrqueting","Gestió de punts de venda i espais comercials"],["CM","Comerç i màrqueting","Transport i logística"],["EA","Energia i aigua","Centrals elèctriques"],["EA","Energia i aigua","Energies renovables"],["EA","Energia i aigua","Eficiència energètica"],["EA","Energia i aigua","Gestió de l'aigua"],["EA","Energia i aigua","Instal·lacions de producció de calor"],["EA","Energia i aigua","Xarxes i estacions de tractament d'aigües"],["EE","Electricitat i electrònica","Instal·lacions elèctriques i automàtiques"],["EE","Electricitat i electrònica","Instal·lacions de telecomunicacions"],["EE","Electricitat i electrònica","Manteniment electromecànic"],["EE","Electricitat i electrònica","Manteniment electrònic"],["EE","Electricitat i electrònica","Sistemes electrotècnics i automatitzats"],["EE","Electricitat i electrònica","Sistemes microinformàtics i xarxes"],["EE","Electricitat i electrònica","Telecomunicacions i informàtica"],["EE","Electricitat i electrònica","Automatització i robòtica industrial"],["EE","Electricitat i electrònica","Energia solar fotovoltaica"],["EO","Edificació i obra civil","Control d'obres"],["EO","Edificació i obra civil","Obres de formigó"],["EO","Edificació i obra civil","Operacions auxiliars d'acabats rígids i urbanització"],["EO","Edificació i obra civil","Operacions auxiliars de paleta de fàbriques i cobertes"],["EO","Edificació i obra civil","Pintura decorativa en construcció"],["EO","Edificació i obra civil","Organització i control d'obres de construcció"],["FM","Fabricació mecànica","Construccions metàl·liques"],["FM","Fabricació mecànica","Conformat i mecanitzat"],["FM","Fabricació mecànica","Disseny en fabricació mecànica"],["FM","Fabricació mecànica","Manteniment i serveis a la producció"],["FM","Fabricació mecànica","Mecanitzat"],["FM","Fabricació mecànica","Soldadura i caldereria"],["FM","Fabricació mecànica","Programació de la producció en fabricació mecànica"],["HS","Hoteleria i turisme","Allotjaments turístics"],["HS","Hoteleria i turisme","Agències de viatges i gestió d'esdeveniments"],["HS","Hoteleria i turisme","Cuina i gastronomia"],["HS","Hoteleria i turisme","Guia, informació i assistències turístiques"],["HS","Hoteleria i turisme","Serveis en restauració"],["HS","Hoteleria i turisme","Gestió d'allotjaments turístics"],["HS","Hoteleria i turisme","Direcció de cuina"],["HS","Hoteleria i turisme","Direcció de serveis en restauració"],["IA","Indústries alimentàries","Cerveseria artesanal"],["IA","Indústries alimentàries","Enologia i viticultura"],["IA","Indústries alimentàries","Indústries làcties"],["IA","Indústries alimentàries","Indústries càrniques"],["IA","Indústries alimentàries","Indústries de conserves i transformats"],["IA","Indústries alimentàries","Fleca, pastisseria i confiteria"],["IA","Indústries alimentàries","Oli d'oliva i derivats"],["IA","Indústries alimentàries","Elaboració de productes alimentaris"],["IM","Instal·lació i manteniment","Instal·lacions frigorífiques i de climatització"],["IM","Instal·lació i manteniment","Instal·lacions de producció de calor"],["IM","Instal·lació i manteniment","Manteniment electromecànic"],["IM","Instal·lació i manteniment","Manteniment de vehicles"],["IM","Instal·lació i manteniment","Instal·lacions de telecomunicacions"],["IM","Instal·lació i manteniment","Manteniment i muntatge d'instal·lacions"],["IM","Instal·lació i manteniment","Muntatge i manteniment d'instal·lacions elèctriques"],["IP","Imatge personal","Perruqueria"],["IP","Imatge personal","Estètica i bellesa"],["IP","Imatge personal","Caracterització i maquillatge professional"],["IP","Imatge personal","Termalisme i benestar"],["IS","Imatge i so","Animacions 3D, jocs i entorns interactius"],["IS","Imatge i so","Producció d'audiovisuals i espectacles"],["IS","Imatge i so","Realització de projectes audiovisuals i espectacles"],["IS","Imatge i so","So per a audiovisuals i espectacles"],["IS","Imatge i so","Vídeo discjòquei i so"],["IS","Imatge i so","Il·luminació, captació i tractament d'imatge"],["IS","Imatge i so","Producció i gestió de projectes d'imatge i so"],["MP","Maritimopesquera","Aqüicultura"],["MP","Maritimopesquera","Cultius aqüícoles"],["QU","Química","Anàlisi i control"],["QU","Química","Laboratori"],["QU","Química","Operacions de laboratori"],["QU","Química","Planta química"],["QU","Química","Fabricació de productes farmacèutics"],["QU","Química","Indústria química"],["QU","Química","Laboratori d'anàlisi i control de qualitat"],["QU","Química","Química industrial"],["QU","Química","Transformació de polímers"],["SA","Sanitat","Cures auxiliars d'infermeria"],["SA","Sanitat","Farmàcia i parafarmàcia"],["SA","Sanitat","Higiene bucodental"],["SA","Sanitat","Laboratori clínic i biomèdic"],["SA","Sanitat","Ortopròtesis i productes de suport"],["SA","Sanitat","Pròtesis dentals"],["SA","Sanitat","Radioteràpia i dosimetria"],["SA","Sanitat","Imatge per al diagnòstic i medicina nuclear"],["SA","Sanitat","Anatomia patològica i citodiagnòstic"],["SA","Sanitat","Documentació i administració sanitàries"],["SC","Serveis socioculturals i a la comunitat","Atenció a persones en situació de dependència"],["SC","Serveis socioculturals i a la comunitat","Educació infantil"],["SC","Serveis socioculturals i a la comunitat","Integració social"],["SC","Serveis socioculturals i a la comunitat","Mediació comunicativa"],["SC","Serveis socioculturals i a la comunitat","Promoció d'igualtat de gènere"],["SC","Serveis socioculturals i a la comunitat","Animació sociocultural i turística"],["SC","Serveis socioculturals i a la comunitat","Direcció i coordinació d'activitats de lleure"],["SC","Serveis socioculturals i a la comunitat","Formació i orientació laboral"],["SC","Serveis socioculturals i a la comunitat","Atenció sociosanitària"],["SC","Serveis socioculturals i a la comunitat","Activitats funeràries i de manteniment en cementiris"],["SM","Seguretat i medi ambient","Emergències i protecció civil"],["SM","Seguretat i medi ambient","Gestió de residus"],["SM","Seguretat i medi ambient","Prevenció de riscos professionals"],["SM","Seguretat i medi ambient","Educació i control ambiental"],["SM","Seguretat i medi ambient","Control i protecció del medi natural"],["SM","Seguretat i medi ambient","Salut ambiental"],["SM","Seguretat i medi ambient","Seguretat"],["TM","Transport i manteniment de vehicles","Automoció"],["TM","Transport i manteniment de vehicles","Carrosseria"],["TM","Transport i manteniment de vehicles","Electromecànica de vehicles"],["TM","Transport i manteniment de vehicles","Manteniment aeromecànic"],["TM","Transport i manteniment de vehicles","Manteniment de vaixells i embarcacions d'esbarjo"],["TM","Transport i manteniment de vehicles","Operacions de manteniment d'elements no estructurals d'aeronau"],["TX","Tèxtil, confecció i pell","Confecció i moda"],["TX","Tèxtil, confecció i pell","Calçat i marroquineria"],["TX","Tèxtil, confecció i pell","Fabricació i ennobliment de productes tèxtils"],["TX","Tèxtil, confecció i pell","Pell i cuir"],["TX","Tèxtil, confecció i pell","Tèxtil i confecció"],["TX","Tèxtil, confecció i pell","Tall i confecció"],["TX","Tèxtil, confecció i pell","Artesania tèxtil"],["IF","Informàtica i comunicacions","Administració de sistemes informàtics"],["IF","Informàtica i comunicacions","Desenvolupament d'aplicacions multiplataforma"],["IF","Informàtica i comunicacions","Desenvolupament d'aplicacions web"],["IF","Informàtica i comunicacions","Sistemes microinformàtics i xarxes"],["IF","Informàtica i comunicacions","Ciberseguretat"],["IF","Informàtica i comunicacions","Intel·ligència artificial i dades"],["AD","Activitats artístiques","Dansa"],["AD","Activitats artístiques","Música"],["AD","Activitats artístiques","Teatre"],["AD","Activitats artístiques","Circ"]];

   // 1) Famílies (una per codi)
   var seen = {};
   var famOrder = 10;

   for (var i = 0; i < FPCAT.length; i++) {
      var famCode = String(FPCAT[i][0] || '').trim();
      var famName = String(FPCAT[i][1] || '').trim();
      if (!famCode || !famName) continue;
      if (seen[famCode]) continue;
      seen[famCode] = true;

      // tipus: FPCAT_FAMILIA, codi: AA, valor: nom en català
      EMI_DICT_upsertCataleg_('FPCAT_FAMILIA', famCode, famName, famCode, famOrder++);
   }

   // 2) Àmbits (un per fila)
   var ambCount = {};
   for (var j = 0; j < FPCAT.length; j++) {
      var c = String(FPCAT[j][0] || '').trim();
      var ambit = String(FPCAT[j][2] || '').trim();
      if (!c || !ambit) continue;

      ambCount[c] = (ambCount[c] || 0) + 1;
      var n = ambCount[c];
      var ambCodi = c + '_' + ('000' + n).slice(-3);   // ex: EE_004

      // tipus: FPCAT_AMBIT, etiqueta: codi família (AA, EE...)
      EMI_DICT_upsertCataleg_('FPCAT_AMBIT', ambCodi, ambit, c, n);
   }
}

/** =========================
  *   Coverage (simple)
  *   ========================= */
function EMI_DICT_updateCoverage_() {
   var ss = EMI_DICT_getMaster_();
   var cov = ss.getSheetByName('DICT_COVERAGE');
   if (!cov) return;

   var cat = EMI_DICT_getCatalegsSheet_();
   if (!cat || cat.getLastRow() < 2) return;

   // --- Llegeix CATALEGS ---
   var lastRow = cat.getLastRow();
   var data = cat.getRange(2, 1, lastRow - 1, Math.max(6, cat.getLastColumn())).getValues();

   var famName = {};           // FPCAT code -> nom humà (valor)
   var tagFam = {};            // TAG codi -> família (etiqueta)
   var tagCount = {};          // família -> #TAGS
   var ambCount = {};          // família -> #AMBITs

   for (var i = 0; i < data.length; i++) {
      var tipus = String(data[i][0] || '').trim();
      var codi = String(data[i][1] || '').trim();
      var valor = String(data[i][2] || '').trim();
      var etiqueta = String(data[i][3] || '').trim();
      var actiu = String(data[i][4] || '').trim().toUpperCase();

      if (actiu && actiu !== 'SI' && actiu !== 'SÍ' && actiu !== 'YES') continue;

      if (tipus === 'FPCAT_FAMILIA' && codi) famName[codi] = valor || codi;

      if (tipus === 'FPCAT_AMBIT' && etiqueta) ambCount[etiqueta] = (ambCount[etiqueta] || 0) + 1;

      if (tipus === 'TAG' && codi) {
         if (etiqueta) tagFam[codi] = etiqueta;
         if (etiqueta) tagCount[etiqueta] = (tagCount[etiqueta] || 0) + 1;
      }
   }

   // --- Llegeix SINONIMS per comptar cobertura real ---
   var synCount = {};
   var synSh = ss.getSheetByName('SINONIMS');
   if (synSh && synSh.getLastRow() >= 2) {
      var synData = synSh.getRange(2, 1, synSh.getLastRow() - 1, Math.max(7, synSh.getLastColumn())).getValues();
      for (var r = 0; r < synData.length; r++) {
         var t = String(synData[r][0] || '').trim();
         var cRef = String(synData[r][1] || '').trim();
         var act = String(synData[r][4] || '').trim().toUpperCase();
         if (act && act !== 'SI' && act !== 'SÍ' && act !== 'YES') continue;
         if (t !== 'TAG' || !cRef) continue;

         var fam = tagFam[cRef] || '';
         if (!fam) continue;

         synCount[fam] = (synCount[fam] || 0) + 1;
      }
   }

   // Neteja sortida
   if (cov.getLastRow() > 1) cov.getRange(2, 1, cov.getLastRow() - 1, 6).clearContent();

   var out = [];
   var now = new Date();
   var famCodes = Object.keys(famName).sort();

   for (var j = 0; j < famCodes.length; j++) {
      var fc = famCodes[j];
      out.push([
         'FPCAT',
         fc,
         famName[fc] || fc,
         (tagCount[fc] || 0),
         (synCount[fc] || 0),
         now
      ]);
   }

   if (out.length) cov.getRange(2, 1, out.length, 6).setValues(out);
}