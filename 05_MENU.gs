/*******************************************************
  * 05_MENU.gs — menú clar EMI · BD (UI-safe)
  * - Menú en català, per passos
  * - Ajuda/guia ràpida visible (modal)
  *******************************************************/

function onOpen(e) {
   // Simple trigger: intenta construir menú, però mai ha de trencar el Sheets
   try { EMI_buildMenu_(); } catch (err) { try { console.log(err); } catch (_) {} }
}

function EMI_buildMenu_() {
   // UI-safe. Si no hi ha UI, NO fem res.
   const ui = ui_();
   if (!ui) return;

   const menu = ui.createMenu('EMI · BD');

   // Ajuda
   menu.addItem('Guia ràpida (què fa cada opció)', 'EMI_menuHelp');
   menu.addSeparator();

   // 1) Preparar Sheets
   const m1 = ui.createMenu('1) Preparar Sheets');
   m1.addItem('1. Crear/actualitzar pestanyes i columnes (Factory)', 'EMI_FACTORY_V2');
   m1.addItem('2. Omplir valors inicials (tècnics/catàlegs/formacions/sinònims)', 'EMI_reseedDefaults');
   m1.addItem('3. Actualitzar desplegables (validacions de Sheets)', 'EMI_refreshDropdowns');
   m1.addSeparator();
   m1.addItem('Mostrar totes les pestanyes', 'EMI_showAllSheets');
   m1.addItem('Amagar pestanyes de suport', 'EMI_hideSupportSheets');
   menu.addSubMenu(m1);

   // 2) Diagnosi
   const m2 = ui.createMenu('2) Diagnosi');
   m2.addItem('Smoke test (comprovar que tot està OK)', 'EMI_smokeTest_V2');
   m2.addItem('Debug (escriu detalls al LOG)', 'EMI_debug');
   m2.addSeparator();
   m2.addItem('Obrir pestanya LOG', 'EMI_openLogSheet');
   m2.addItem('Obrir pestanya HEALTHCHECK', 'EMI_openHealthSheet');
   menu.addSubMenu(m2);

   // 3) Automatitzacions
   const m3 = ui.createMenu('3) Automatitzacions');
   m3.addItem('Instal·lar automatitzacions (V3 recomanat)', 'EMI_installTriggers_V3');
   m3.addItem('Instal·lar automatitzacions (V2)', 'EMI_installTriggers_V2');
   m3.addItem('Llistar automatitzacions (escriu al LOG)', 'EMI_listTriggers_V2');
   m3.addItem('Eliminar automatitzacions (triggers)', 'EMI_removeTriggers_V2');
   menu.addSubMenu(m3);


   // 4) Motor (matching)
   const m4 = ui.createMenu('4) Motor (matching)');
   m4.addItem('Reindex + Matching ara (top 25)', 'EMI_PIPELINE_rebuildBrainNow_25');
   m4.addItem('Inicialitzar cervell (1 cop)', 'EMI_brainBootstrap');
   menu.addSubMenu(m4);

   // Avançat (ho deixem aquí però no cal tocar-ho ara)
   const mx = ui.createMenu('Avançat');
   mx.addItem('Preparar claus per AppSheet (id_sinonim a SINONIMS)', 'EMI_appsSheetFixKeys');
   menu.addSubMenu(mx);

   menu.addToUi();
}

/**
  * Ajuda: explica en llenguatge planer què fa cada opció.
  * (No depèn de CFG.UI_ALERTS; és un diàleg modal.)
  */
function EMI_menuHelp() {
   const ui = ui_();
   if (!ui) return;

   const html = HtmlService.createHtmlOutput(
      [
         '<div style="font-family:Arial; line-height:1.35; padding:12px;">',
         '<h2 style="margin:0 0 8px 0;">EMI · BD — Guia ràpida</h2>',
         '<p style="margin:0 0 10px 0;">Objectiu: deixar el <b>Google Sheet MASTER</b> preparat (després ja passarem a AppSheet).</p>',

         '<h3 style="margin:12px 0 6px 0;">1) Preparar Sheets</h3>',
         '<ul>',
         '<li><b>Factory</b>: crea les pestanyes que falten i afegeix columnes noves si cal (no hauria d’esborrar dades).</li>',
         '<li><b>Omplir valors inicials</b>: posa dades base a <i>TECNICS/CATALEGS/FORMACIONS</i> i genera <i>SINONIMS</i> a partir de CATALEGS.</li>',
         '<li><b>Desplegables</b>: crea/actualitza <i>_DV_LISTS</i> i aplica validacions (dropdowns) a columnes clau.</li>',
         '<li><b>Mostrar/Amagar pestanyes</b>: només és visibilitat (no canvia dades).</li>',
         '</ul>',

         '<h3 style="margin:12px 0 6px 0;">2) Diagnosi</h3>',
         '<ul>',
         '<li><b>Smoke test</b>: comprova que existeixen pestanyes, carpetes i mínims; deixa rastre a <i>LOG</i> i <i>HEALTHCHECK</i>.</li>',
         '<li><b>Debug</b>: escriu informació de versió i estat al <i>LOG</i>.</li>',
         '<li><b>Obrir LOG/HEALTHCHECK</b>: et porta directament a aquestes pestanyes.</li>',
         '</ul>',

         '<h3 style="margin:12px 0 6px 0;">3) Automatitzacions</h3>',
         '<ul>',
         '<li><b>Instal·lar triggers</b>: crea onOpen/onEdit + CRON 5 minuts + CRON diari (important per entorns tipus AppSheet).</li>',
         '<li><b>Llistar triggers</b>: escriu al LOG quins triggers hi ha actius.</li>',
         '<li><b>Eliminar triggers</b>: esborra els triggers instal·lats d’aquest projecte.</li>',
         '</ul>',

         '<h3 style="margin:12px 0 6px 0;">Avançat</h3>',
         '<ul>',
         '<li><b>Claus AppSheet</b>: afegeix/omple <i>id_sinonim</i> a la taula SINONIMS perquè AppSheet tingui una Key estable.</li>',
         '</ul>',

         '<p style="margin:12px 0 0 0; color:#555;">Consell: si no veus missatges emergents, revisa sempre la pestanya <b>LOG</b>.</p>',
         '</div>'
      ].join('')
   ).setWidth(520).setHeight(640);

   ui.showModalDialog(html, 'EMI · BD — Guia ràpida');
}

/** Utilitats petites per anar directe a pestanyes clau */
function EMI_openLogSheet() {
   try {
      const ss = getMasterSs_();
      const sh = ss.getSheetByName(SHEETS.LOG);
      if (sh) ss.setActiveSheet(sh);
   } catch (e) { try { console.log(e); } catch (_) {} }
}

function EMI_openHealthSheet() {
   try {
      const ss = getMasterSs_();
      const sh = ss.getSheetByName(SHEETS.HEALTH);
      if (sh) ss.setActiveSheet(sh);
   } catch (e) { try { console.log(e); } catch (_) {} }
}

/** Debug simple (ja el tenies, el deixo aquí perquè sigui accessible des del menú) */
function EMI_debug() {
   const ss = getMasterSs_();
   log_('INFO', 'DEBUG', 'Debug OK', {
      version: (typeof EMI_BD_VERSION !== 'undefined' ? EMI_BD_VERSION : 'MISSING'),
      tz: (typeof EMI_TZ !== 'undefined' ? EMI_TZ : 'MISSING'),
      ssId: ss.getId(),
      sheetsCount: ss.getSheets().length
   });
   // Això pot quedar “silenciat” si CFG.UI_ALERTS=false, però el LOG sempre hi serà.
   uiAlert_('Debug OK ✅\nMira la pestanya LOG.');
}