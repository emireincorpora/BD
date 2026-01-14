/*******************************************************
  * 00_CONFIG.gs — EMI BD MASTER (AppSheet) · Factory V2
  *******************************************************/

const EMI_TZ = 'Europe/Madrid';
const EMI_BD_VERSION = 'EMI_MASTER_APPSHEET_V7.1.0'; // <-- AQUESTA LÍNIA ÉS LA CLAU
const EMI_MASTER_SS_ID_DEFAULT = '18kxjEexv086vvIXwXOgEajyi8EExp7r-v8plKFNqBz0';

const CFG = {
   ROOT_FOLDER_NAME: 'EMI_BD_ROOT',
   SUBFOLDER_ATTACHMENTS: '_ATTACHMENTS',
   SUBFOLDER_EXPORTS: '_EXPORTS',

   CLEAN_DELETE_UNKNOWN_SHEETS: false,

   UI_ALERTS: false,
   UI_TOASTS: false,

   HIDE_SUPPORT_SHEETS_BY_DEFAULT: true
};

const PROP_KEYS = {
   VERSION: 'EMI_BD_VERSION',
   MASTER_SS_ID: 'EMI_BD_MASTER_SS_ID',

   ROOT_FOLDER_ID: 'EMI_BD_ROOT_FOLDER_ID',
   ATTACH_FOLDER_ID: 'EMI_BD_ATTACH_FOLDER_ID',
   EXPORTS_FOLDER_ID: 'EMI_BD_EXPORTS_FOLDER_ID',

   SEQ_PERSONA_PREFIX: 'EMI_SEQ_PERSONA_',
   SEQ_EMPRESA_PREFIX: 'EMI_SEQ_EMPRESA_',
   SEQ_VACANT_PREFIX: 'EMI_SEQ_VACANT_'
};

const SHEETS = {
   PERSONES: 'PERSONES',
   EMPRESES: 'EMPRESES',
   VACANTS: 'VACANTS',
   CANDIDATURES: 'CANDIDATURES',

   SEGUIMENTS_PERSONA: 'SEGUIMENTS_PERSONA',
   SEGUIMENTS_EMPRESA: 'SEGUIMENTS_EMPRESA',

   DOCUMENTS: 'DOCUMENTS',
   CV_RAW: 'CV_RAW',
   VACANT_SUGGERIMENTS: 'VACANT_SUGGERIMENTS',

   TECNICS: 'TECNICS',
   CATALEGS: 'CATALEGS',
   SINONIMS: 'SINONIMS',
   FORMACIONS_CATALOG: 'FORMACIONS_CATALOG',
   NAV_HOME: 'NAV_HOME',
   DV_LISTS: '_DV_LISTS',
   SEARCH_REQUESTS: 'SEARCH_REQUESTS',
   SEARCH_RESULTS: 'SEARCH_RESULTS',
   STATE: 'STATE',

   LOG: 'LOG',
   HEALTH: 'HEALTHCHECK'
};

const HEADERS = {
   [SHEETS.LOG]: ['ts','nivell','accio','missatge','json'],
   [SHEETS.HEALTH]: ['ts','check','status','message','json'],

   [SHEETS.TECNICS]: ['id_tecnic','nom','email','rol','actiu','ordre'],
   [SHEETS.CATALEGS]: ['tipus','codi','valor','etiqueta','actiu','ordre'],
   [SHEETS.SINONIMS]: ['tipus','codi_ref','sinonim','pes','actiu','notes','id_sinonim'],

   [SHEETS.FORMACIONS_CATALOG]: [
      'codi_formacio','tipus','nom','familia','nivell','durada_hores',
      'sinonims','actiu','ordre','notes','updated_at'
   ],

   [SHEETS.PERSONES]: [
      'id_persona', 'emi_id_persona', 'data_alta', 'origen_alta', 'estat_fitxa',
      'nom', 'cognom1', 'cognom2', 'data_naixement', 'genere',
      'telefon_mobil', 'email', 'doc_tipus', 'doc_numero', 'adreca',
      'cp', 'municipi', 'comarca', 'situacio_administrativa', 'situacio_laboral_actual',
      'disponibilitat_immediata', 'jornada_preferida', 'zona_treball', 'carnet_conduir', 'vehicle_propi',
      'nivell_formacio', 'ambits_interes', 'programes_interes', 'tags_persona',
      'observacions', 'tecnic_referent', 'created_by_email', 'updated_by_email', 'tags_auto', 'tags_persona_auto',
      'data_ultima_actualitzacio', 'text_index',
      'cv_source_ts', 'cv_sector_objectiu', 'cv_competencies', 'cv_idiomes', 'cv_experiencia', 'cv_formacio',
      'cv_voluntariat', 'cv_horari', 'cv_incorporacio', 'cv_text_index'
   ],

   [SHEETS.EMPRESES]: [
      'id_empresa','emi_id_empresa','data_alta','origen_alta','estat_fitxa',
      'nom_empresa','nom_empresa_normalitzat','cif',
      'sector_principal','altres_sectors',
      'adreca','cp','municipi','comarca',
      'telefon_general','email_general','web',
      'contacte_nom','contacte_carrec','contacte_tel','contacte_email',
      'busques_a_emi','tags_empresa','observacions',
      'tecnic_referent','created_by_email','updated_by_email','tags_auto','data_ultima_actualitzacio','text_index'
   ],

   [SHEETS.VACANTS]: [
      'id_vacant', 'emi_id_vacant', 'data_alta', 'estat', 'id_empresa',
      'empresa_id', 'nom_empresa', 'titol_vacant', 'descripcio', 'municipi',
      'comarca', 'sector', 'requisits', 'horari', 'jornada',
      'salari', 'tags_vacant', 'tecnic_referent', 'tags_auto', 'tags_vacant_auto', 'text_index', 'updated_at'
   ],

   [SHEETS.CANDIDATURES]: [
      'id_candidatura','id_vacant','id_persona','ts','estat','notes'
   ],

   [SHEETS.SEGUIMENTS_PERSONA]: [
      'id_event','id_persona','ts','tipus','resum','detall','created_by_email'
   ],
   [SHEETS.SEGUIMENTS_EMPRESA]: [
      'id_event','id_empresa','ts','tipus','resum','detall','created_by_email'
   ],

   [SHEETS.DOCUMENTS]: [
      'id_document','entity_type','entity_id','ts','doc_type','file_url','notes','created_by_email'
   ],

   [SHEETS.CV_RAW]: [
      'id_cv', 'source_row', 'source_ts', 'ingested_ts', 'email', 'nom', 'cognoms', 'telefon', 'doc',
      'municipi', 'comarca', 'file_urls', 'raw_json', 'linked_id_persona', 'status', 'notes',
      'sector_objectiu', 'cv_competencies', 'cv_idiomes', 'cv_experiencia', 'cv_formacio', 'cv_voluntariat',
      'cv_horari', 'cv_incorporacio', 'carnet_conduir', 'vehicle_propi', 'cv_text_index', 'applied_to_persona_ts'
   ],

   [SHEETS.VACANT_SUGGERIMENTS]: [
      'id_sug','id_vacant','id_persona','score','reasons','kw_hits','tag_hits','created_at'
   ],

   [SHEETS.NAV_HOME]: [
      'id_nav','ordre','titol','subtitol','emoji','icon_url','target_type','target_name','actiu'
   ],

   [SHEETS.DV_LISTS]: [
      'TECNICS_ACTIUS','SECTOR_EMPRESA','AMBIT_FORMACIO','PROGRAMA_EMI','TAG_PERSONA','TAG_EMPRESA'
   ],

   [SHEETS.SEARCH_REQUESTS]: [
      'id_search','ts','user_email','entity','query_text','filters_json','topN','status'
   ],
   [SHEETS.SEARCH_RESULTS]: [
      'id_search','rank','entity','entity_id','score','reasons','created_at'
   ],

   [SHEETS.STATE]: [
      'key','value','updated_at','notes','json'
   ]
};

const UI_VISIBLE_SHEETS = [
   SHEETS.PERSONES,
   SHEETS.EMPRESES,
   SHEETS.VACANTS,
   SHEETS.CANDIDATURES
];

const SUPPORT_SHEETS = Object.keys(SHEETS).map(k => SHEETS[k])
   .filter(n => UI_VISIBLE_SHEETS.indexOf(n) === -1);
