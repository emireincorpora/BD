# EMI BD (CRM + ATS + Search + Matching) — Deliverables

> **Nota d’honestedat:** els artefactes indicats al prompt (Excel/Docx/Zip) **no són presents** en aquest repositori i no s’han pogut auditar directament. Si els comparteixes, puc fer una auditoria de desalineacions exacta. Mentrestant, he alineat el codi amb el model declarat al prompt i he deixat la infraestructura per migració idempotent, reindex batch i matching resumable.

## 1) Informe d’auditoria (estat real vs. requerit)

**BLOCKER**
- **No hi ha arxius MASTER ni documentació** (`EMI_BD_MASTER (2).xlsx`, `*.docx`, `*.zip`) al repositori. Sense aquests artefactes no es pot validar l’estructura real ni discrepàncies exactes entre headers i codi. Cal adjuntar-los per fer l’auditoria completa.
- **Pipeline i matching eren monolítics** (reindex complet + matching complet), amb risc de timeout en volum moderat. Això s’ha resolt amb pipeline batch/resumable i estat persistent.

**MAJOR**
- **Headers del model no estaven alineats amb el contracte functional** (faltaven `text_index`, `tags_*_auto`, `CV_RAW`, `VACANT_SUGGERIMENTS`, `NAV_HOME`, `_DV_LISTS`, `STATE`, etc.). S’ha corregit el contracte de headers i els fulls de suport.
- **Matching sense configuració editable**: ara es llegeixen pesos/bònnus via `CATALEGS (tipus=MATCH_CFG)`.

**MINOR**
- **Desalineació del camp legacy `sector_objectiu`**: es fa backfill a `cv_sector_objectiu` per no perdre dades.

## 2) Disseny final (contracte de dades + polítiques)

### Model de dades (contracte actualitzat)
- **Core**: `PERSONES`, `EMPRESES`, `VACANTS`, `CANDIDATURES`.
- **CRM/traçabilitat**: `SEGUIMENTS_*`, `DOCUMENTS`.
- **Staging**: `CV_RAW`.
- **Outputs**: `VACANT_SUGGERIMENTS`.
- **Admin**: `TECNICS`, `CATALEGS`, `SINONIMS`, `FORMACIONS_CATALOG`, `_DV_LISTS`, `NAV_HOME`.
- **Operació**: `STATE`, `LOG`, `HEALTHCHECK`.
- **Opcional avançat**: `SEARCH_REQUESTS`, `SEARCH_RESULTS`.

### Camps derivats i persistents
- `text_index` a `PERSONES/EMPRESES/VACANTS` (persistència per matching/cerca).
- `tags_auto` (interns) + `tags_persona_auto` / `tags_vacant_auto` (llegibles).

### Política CV → PERSONES
- **Read-only** sobre el source del CV.
- **Merge segur**: `setIfEmpty` per camps bàsics; no sobreescriu camps manuals.
- **Idempotència** amb cursor de fila (PropertiesService + `CV_RAW`).

### Motor de matching (configurable)
- **Score híbrid**: `kwScore` + `tagScore` amb pesos de `MATCH_CFG`.
- **Bonus** per tokens especials (ex. `carnet_*`).
- **Reasons** explicables (keywords + tags amb etiqueta llegible + tokens especials).

### Motor de cerca
- **Nivell 1**: AppSheet + `text_index` + tags + facets (_DV_LISTS).
- **Nivell 2 (opcional)**: `SEARCH_REQUESTS`/`SEARCH_RESULTS` (Apps Script).

## 3) Pla de migració idempotent

1. **Factory**: `EMI_FACTORY_V2()` assegura fulls i headers.
2. **Migració cervell**: `EMI_brainBootstrap()` o `EMI_PIPELINE_rebuildBrainNow()`.
3. **Backfill alias**: `sector_objectiu` → `cv_sector_objectiu` via `ensureCriticalHeadersAndAliases_`.
4. **Reindex batch**: Pipeline per lots (evita timeouts).
5. **Matching batch**: per vacants amb neteja local (no wipe global).
6. **Validació**: `EMI_DIAGNOSE_masterNow()` + `LOG/HEALTHCHECK`.

## 4) Codi final (mòduls actualitzats / nous)

- **Actualitzats**: `00_CONFIG.gs`, `02_SHEETS_SETUP.gs`, `03_FACTORY.gs`, `06_TRIGGERS_IDS.gs`, `07_DIAG.gs`, `10_BRAIN_INDEX_MATCH.gs`, `19_DIAGNOSE_AND_FIX.gs`, `22_TRIGGERS_AUDIT.gs`.
- **Nou**: `24_PIPELINE_BATCH.gs` (pipeline resumable + estat persistent).

## 5) Guia d’implantació pas a pas

1. **Config**: fixa `EMI_BD_MASTER_SS_ID` (Script Properties).
2. **Factory**: executa `EMI_FACTORY_V2()`.
3. **Migració cervell**: executa `EMI_brainBootstrap()` o `EMI_PIPELINE_rebuildBrainNow(25)`.
4. **Triggers**: `EMI_installTriggers_V3()`.
5. **CV**: si cal, defineix `EMI_CV_SOURCE_SPREADSHEET_ID` (Script Properties) i prova `EMI_CV_ingestNewResponses()`.
6. **Smoke test**: `EMI_DIAGNOSE_masterNow()`.

## 6) Checklist AppSheet (taules, vistes, accions, seguretat)

### Taules
- Core: `PERSONES`, `EMPRESES`, `VACANTS`, `CANDIDATURES`.
- Relacionades: `DOCUMENTS`, `SEGUIMENTS_PERSONA`, `SEGUIMENTS_EMPRESA`.
- Output: `VACANT_SUGGERIMENTS`.
- Admin: `TECNICS`, `CATALEGS`, `SINONIMS`, `FORMACIONS_CATALOG`, `NAV_HOME`, `_DV_LISTS`, `STATE`.

### Vistes
- **INICI**: NAV_HOME com a dashboard.
- **PERSONES**: llista + detall + edició + subviews (seguiments/documents).
- **EMPRESES**: llista + detall + vacants relacionades.
- **VACANTS**: llista + detall + suggeriments inline.
- **CANDIDATURES**: pipeline d’estats.
- **ADMIN**: taules admin + logs.

### Accions
- **Reindex (batch)**: `EMI_PIPELINE_rebuildBrainNow`.
- **Recalcular matching**: `EMI_recomputeMatchesBothSidesNow`.
- **Diagnosi**: `EMI_DIAGNOSE_masterNow`.

### Seguretat / rols
- Control per `TECNICS.rol` (Admin / Orientador / Lectura).
- Restringeix edició de taules Admin a rols d’administració.
