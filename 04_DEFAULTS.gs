/*******************************************************
  * 04_DEFAULTS.gs — seeds robustos (TECNICS + CATALEGS + SINONIMS)
  *******************************************************/

function ensureDefaultTecnics_() {
   const sh = getSheet_(SHEETS.TECNICS);
   if (sh.getLastRow() > 1) return;

   const rows = [
      ['SISTEMA','SISTEMA (automat)','', 'ADMIN','SI', 1],
      ['RECEPCIO','Recepció','', 'RECEPCIO','SI', 2],
      ['SENSE_TECNIC','Sense tècnic','', 'TECNIC','SI', 3],

      ['ARI','Ari Olmeda','', 'TECNIC','SI', 10],
      ['DANI','Dani Rodríguez','', 'TECNIC','SI', 11],
      ['ANNA','Anna Martínez','', 'TECNIC','SI', 12],
      ['PERE','Pere Solera','', 'TECNIC','SI', 13],
      ['FF','Francesc Ferre','', 'TECNIC','SI', 14],
      ['FC','Francesc Camino','', 'TECNIC','SI', 15],
      ['EVA','Eva González','', 'TECNIC','SI', 16],
      ['JOSEP','Josep Estany','', 'ADMIN','SI', 17],
      ['ISAAC','Isaac Vázquez','', 'TECNIC','SI', 18],

      ['ALTRES','Altres','', 'TECNIC','SI', 99]
   ];

   sh.getRange(2,1,rows.length,HEADERS[SHEETS.TECNICS].length).setValues(rows);
   log_('INFO','DEFAULT_TECNICS_OK','TECNICS seed creat', { count: rows.length });
}

/**
  * CATALEGS robustos (extensibles):
  * - SECTOR_EMPRESA (amb sinònims)
  * - AMBIT_FORMACIO (amb sinònims)
  * - PROGRAMA_EMI (amb sinònims)
  * - TAG_PERSONA / TAG_EMPRESA (amb sinònims)
  */
function ensureCatalogDefaults_() {
   const sh = getSheet_(SHEETS.CATALEGS);

   // Upsert per (tipus+codi)
   const existing = {};
   if (sh.getLastRow() > 1) {
      const data = sh.getRange(2,1,sh.getLastRow()-1,6).getValues();
      data.forEach(r => {
         const k = String(r[0]||'') + '||' + String(r[1]||'');
         if (r[0] && r[1]) existing[k] = true;
      });
   }

   const rows = [];

   function add_(tipus, codi, valor, etiqueta, ordre) {
      const k = tipus + '||' + codi;
      if (existing[k]) return;
      rows.push([tipus, codi, valor, etiqueta || '', 'SI', ordre || 999]);
      existing[k] = true;
   }

   // SECTOR_EMPRESA (inclou sinònims pipe, basat en el que ja tenies treballat)
   add_('SECTOR_EMPRESA','METALL','Metall i soldadura',
      'metall|metal|soldadura|soldador|ferralla|caldereria|serralle...|cnc|torner|fresador|mecànica', 10);
   add_('SECTOR_EMPRESA','ELEC','Electricitat i instal·lacions',
      'electricitat|electricista|instal·lacions|instal·lador|manten...|domòtica|climatització|aerotèrmia|llum|il·luminació', 20);
   add_('SECTOR_EMPRESA','LOG','Logística i magatzem',
      'logística|magatzem|almacén|carretiller|carretó elevador|pick...|repartiment|paqueteria|transport|inventari|estoc', 30);
   add_('SECTOR_EMPRESA','ADM','Administració i gestió',
      'administració|gestió|oficina|backoffice|recepció|secretaria|...|arxiu|gestió documental|excel|erp|tràmits', 40);
   add_('SECTOR_EMPRESA','COM','Comerç i atenció al client',
      'comerç|vendes|retail|botiga|dependenta|caixa|atenció client|...|teleoperador|call center|comercial|crm|ecommerce', 50);
   add_('SECTOR_EMPRESA','HOS','Hostaleria i cuina',
      'hostaleria|hotel|restauració|bar|cafeteria|camarer|cambrer|...|cuina|cuiner|ajudant cuina|càtering|rentaplats', 60);
   add_('SECTOR_EMPRESA','CURES','Cures i atenció a les persones',
      'cures|atenció persones|sociosanitari|geriatria|auxiliar|taps...|residència|assistència personal|cangur|acompanyament', 70);
   add_('SECTOR_EMPRESA','RENOV','Renovables',
      'renovables|energia solar|fotovoltaica|plaques solars|autocon...|eficiència energètica|sostenibilitat|biomassa|aerotèrmia', 80);
   add_('SECTOR_EMPRESA','ALTRES','Altres','altres|varis|diversos|no ho sé', 999);

   // AMBIT_FORMACIO
   add_('AMBIT_FORMACIO','LOG','Logística i magatzem',
      'logística|magatzem|carretó elevador|picking|packing|estocs|inventari|preparació comandes', 10);
   add_('AMBIT_FORMACIO','METALL','Metall i soldadura',
      'metall|soldadura|mecànica|cnc|torner|fresador|caldereria|serralleria', 20);
   add_('AMBIT_FORMACIO','ELEC','Electricitat i instal·lacions',
      'electricitat|instal·lacions|manteniment|quadres|cablejat|automatismes|domòtica', 30);
   add_('AMBIT_FORMACIO','RENOV','Renovables',
      'renovables|fotovoltaica|plaques solars|autoconsum|eficiència energètica|sostenibilitat', 40);
   add_('AMBIT_FORMACIO','ADM','Administració i gestió',
      'administració|gestió|ofimàtica|excel|facturació|comptabilitat|arxiu|tràmits', 50);
   add_('AMBIT_FORMACIO','COM','Comerç i atenció al client',
      'comerç|vendes|atenció client|caixa|retail|teleoperador|crm', 60);
   add_('AMBIT_FORMACIO','CURES','Cures i atenció a les persones',
      'cures|sociosanitari|geriatria|dependència|residència|domicili|assistència personal', 70);
   add_('AMBIT_FORMACIO','HOS','Hostaleria i cuina',
      'hostaleria|cuina|sala|camarer|ajudant cuina|càtering|hotel', 80);
   add_('AMBIT_FORMACIO','ALTRES','Altres','altres|varis|diversos|no ho sé', 999);

   // PROGRAMA_EMI
   add_('PROGRAMA_EMI','INC','INCORPORA','incorpora|inserció laboral|intermediació|ofertes', 10);
   add_('PROGRAMA_EMI','REI','REINCORPORA','reincorpora|reincorporació|retorn al mercat laboral', 20);
   add_('PROGRAMA_EMI','AVA','AVANT','avant|orientació|itinerari|millora ocupabilitat', 30);
   add_('PROGRAMA_EMI','FOAP','FOAP','foap|formació ocupacional|certificat professionalitat', 40);
   add_('PROGRAMA_EMI','FPOD','FPO Dual','fpo dual|dual|aprenentatge|empresa|pràctiques', 50);
   add_('PROGRAMA_EMI','AUTO','Autoocupació','autoocupació|emprenedoria|autònom|negoci propi|plans empresa', 60);
   add_('PROGRAMA_EMI','CONS','Consolida’t','consolida’t|consolidar negoci|creixement|gestió|finances', 70);
   add_('PROGRAMA_EMI','NS','No ho sé','no ho sé|no coneixo|dubtes', 900);
   add_('PROGRAMA_EMI','ALTRES','Altres','altres|varis', 999);

   // TAG_PERSONA
   add_('TAG_PERSONA','DOCS','Documentació','dni|nie|passaport|empadronament|permís treball|permís residència|dardo|soc|sepe|cv', 10);
   add_('TAG_PERSONA','MOBIL','Mobilitat','carnet b|carnet|vehicle propi|desplaçament|disponibilitat immediata', 20);
   add_('TAG_PERSONA','LOGI','Perfil logístic','carretó elevador|picking|packing|magatzem|repartiment', 30);
   add_('TAG_PERSONA','ADM','Perfil administratiu','ofimàtica|excel|arxiu|facturació|recepció', 40);
   add_('TAG_PERSONA','IDI','Idiomes','català|castellà|anglès|francès|àrab|urdú|wòlof|amazic', 50);

   // TAG_EMPRESA
   add_('TAG_EMPRESA','BPF','Bones pràctiques','responsabilitat social|inclusió|diversitat|igualtat|rsc', 10);
   add_('TAG_EMPRESA','NEEDS','Necessitats','vacants|contractació|formació plantilla|pràctiques|duals', 20);

   if (!rows.length) {
      log_('INFO','DEFAULT_CATALEGS_SKIP','CATALEGS ja tenia dades (no seed)', {});
      return;
   }

   sh.getRange(sh.getLastRow()+1,1,rows.length,6).setValues(rows);
   log_('INFO','DEFAULT_CATALEGS_OK','CATALEGS seed creat (upsert)', { added: rows.length });
}

function ensureFormacionsDefaults_() {
   const sh = getSheet_(SHEETS.FORMACIONS_CATALOG);

   // Upsert per codi_formacio
   const existing = {};
   if (sh.getLastRow() > 1) {
      const data = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
      data.forEach(r => { if (r[0]) existing[String(r[0])] = true; });
   }

   const now = toIso_(new Date());
   const rows = [];

   function add_(codi, tipus, nom, familia, nivell, hores, sinonims, ordre, notes) {
      if (existing[codi]) return;
      rows.push([codi, tipus, nom||'', familia||'', nivell||'', hores||'', sinonims||'', 'SI', ordre||999, notes||'', now]);
      existing[codi] = true;
   }

   // CP codes que heu citat (nom pot quedar buit si encara no voleu omplir-lo)
   add_('COML0110','CP','', '','', '', 'coml0110', 10, 'Afegiu nom oficial quan el tingueu');
   add_('ELEQ0111','CP','', '','', '', 'eleq0111', 20, 'Afegiu nom oficial quan el tingueu');
   add_('ADGD0308','CP','', '','', '', 'adgd0308|agdg0308', 30, 'Nota: corregit el typo AGDG0308 → ADGD0308');
   add_('ADGG0408','CP','', '','', '', 'adgg0408', 40, 'Afegiu nom oficial quan el tingueu');
   add_('ENAE0108','CP','', '','', '', 'enae0108', 50, 'Afegiu nom oficial quan el tingueu');
   add_('ENAE0208','CP','', '','', '', 'enae0208', 60, 'Afegiu nom oficial quan el tingueu');
   add_('FMEC0219','CP','', '','', '', 'fmec0219', 70, 'Afegiu nom oficial quan el tingueu');

   if (!rows.length) {
      log_('INFO','DEFAULT_FORMACIONS_SKIP','FORMACIONS_CATALOG ja tenia dades', {});
      return;
   }

   sh.getRange(sh.getLastRow()+1,1,rows.length,HEADERS[SHEETS.FORMACIONS_CATALOG].length).setValues(rows);
   log_('INFO','DEFAULT_FORMACIONS_OK','FORMACIONS_CATALOG seed creat', { added: rows.length });
}

/**
  * Genera SINONIMS a partir de CATALEGS.etiqueta (pipe-separated).
  * És la clau per tenir un “diccionari” ampliable sense tocar codi:
  * - edites CATALEGS.etiqueta
  * - executes EMI_syncSinonimsFromCatalogs
  */
function syncSinonimsFromCatalogs_() {
   const shCat = getSheet_(SHEETS.CATALEGS);
   const shSyn = getSheet_(SHEETS.SINONIMS);

   const catLast = shCat.getLastRow();
   if (catLast <= 1) {
      log_('WARN','SYNC_SINONIMS_EMPTY','CATALEGS buit: no puc generar sinònims', {});
      return;
   }

   // index existent (tipus||codi_ref||sinonim_normalitzat)
   const synIndex = {};
   const synLast = shSyn.getLastRow();
   if (synLast > 1) {
      const data = shSyn.getRange(2,1,synLast-1,5).getValues();
      data.forEach(r => {
         const tipus = String(r[0]||'');
         const codi = String(r[1]||'');
         const sin = normalizeText_(String(r[2]||''));
         if (tipus && codi && sin) synIndex[tipus+'||'+codi+'||'+sin] = true;
      });
   }

   const catData = shCat.getRange(2,1,catLast-1,6).getValues();
   const out = [];

   function pushSyn_(tipus, codiRef, raw, pes, notes) {
      const sinN = normalizeText_(raw);
      if (!sinN) return;
      const key = tipus + '||' + codiRef + '||' + sinN;
      if (synIndex[key]) return;
      out.push([tipus, codiRef, raw, pes || 1, 'SI', notes || 'from CATALEGS']);
      synIndex[key] = true;
   }

   catData.forEach(r => {
      const tipus = String(r[0]||'').trim();
      const codi = String(r[1]||'').trim();
      const valor = String(r[2]||'').trim();
      const etiqueta = String(r[3]||'').trim();
      const actiu = String(r[4]||'').trim().toUpperCase();

      if (!tipus || !codi || actiu !== 'SI') return;

      // sempre incloem el valor com a “sinònim” base
      if (valor) pushSyn_(tipus, codi, valor, 2, 'canonical');

      if (!etiqueta) return;

      // split pipe
      etiqueta.split('|').forEach(tok => {
         const t = String(tok||'').trim();
         if (t) pushSyn_(tipus, codi, t, 1, 'tag');
      });
   });

   if (!out.length) {
      log_('INFO','SYNC_SINONIMS_SKIP','No hi havia sinònims nous per afegir', {});
      return;
   }

   shSyn.getRange(shSyn.getLastRow()+1,1,out.length,6).setValues(out);
   log_('INFO','SYNC_SINONIMS_OK','Sinònims actualitzats', { added: out.length });
}

/******** Commands exposed ********/

function EMI_reseedDefaults() {
   withLock_('RESEED_DEFAULTS', () => {
      ensureDefaultTecnics_();
      ensureCatalogDefaults_();
      ensureFormacionsDefaults_();
      syncSinonimsFromCatalogs_();
      log_('INFO','RESEED_OK','Reseed complet', {});
      uiAlert_('Reseed complet ✅');
   });
}

function EMI_syncSinonimsFromCatalogs() {
   withLock_('SYNC_SINONIMS', () => {
      syncSinonimsFromCatalogs_();
      uiAlert_('Sinònims actualitzats ✅ (revisa LOG)');
   });
}