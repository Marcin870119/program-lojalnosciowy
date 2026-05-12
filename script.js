const jsonUrl =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Slodycze%20Ranking%20Rumunia.json';
if(window.__wfScriptLoaded){
  console.warn('script.js loaded more than once; duplicate bindings will be skipped.');
}
window.__wfScriptLoaded = true;
const quickOrderBootConfig = (() => {
  const params = new URLSearchParams(window.location.search);
  const token = String(params.get('quickOrder') || '').trim();
  return {
    isPublicMode: Boolean(token),
    token
  };
})();
window.__WF_QUICK_ORDER_BOOT__ = quickOrderBootConfig;
if(quickOrderBootConfig.isPublicMode && document.body){
  document.body.classList.add('quick-order-mode');
}
const jsonUrlMieso =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Mieso%20Wedliny%20-%20Rumunia.json';
const jsonUrlNabial =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Nabial%20-%20Rumunia%20Ranking.json';
const jsonUrlNapoje =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Napoje%20-%20Ranking%20Rumunia.json';
const jsonUrlPrzyprawyProszek =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/PRZYPRAWY%20I%20DODATKI%20W%20PROSZKU%20RUMUNIA%20-%20Ranking%20Rumunia.json';
const jsonUrlPuszkiSloiki =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/PUSZKI%20I%20S%C5%81OIKI%20%20-%20Ranking%20Rumunia.json';
const jsonUrlProduktyPodstawowe =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Produkty%20Podstawowe%20-%20Rumunia%20Ranking.json';
const jsonUrlKawyHerbaty =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Kawa%20i%20Herbata%20-%20Ranking%20Rumunia.json';
const jsonUrlTopRumunia =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Top%20Rumunia.json';
const jsonUrlCustomerDiscounts =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/kh%3Aph%20-%20world%20food%20rabaty.json';
const jsonUrlSlodyczeUkraina =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/SLODYCZE%20UKRAINA%20-%20RANKING.json';
const jsonUrlMiesoUkraina =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Mieso%20i%20wedliny%20-%20Ukraina%20Ranking.json';
const jsonUrlKawyUkraina =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/KAWA%20I%20HERBATA%20-%20UKRAINA%20RANKING.json';
const xlsxUrlPuszkiUkraina =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Puszki%20i%20s%C5%82oiki%20%20-%20UKRAINA%20RANKING.xlsx';
const jsonUrlNapojeUkraina =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/NAPOJE%20-%20UKRAINA%20RANKING.json';
const jsonUrlPrzyprawyUkraina =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Przyprawy%20i%20dodatki%20-%20ukraina%20ranking.json';
const jsonUrlTopUkraina =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/UKRAINA%20-%20TOP%20RANKING.json';
const ukrainaDodatkiDoPotrawSource = {
  name: 'Ukraina - Dodatki do potraw',
  type: 'json',
  // URL uzupelnimy po wrzuceniu pliku z baza danych dla tego kafelka.
  url: ''
};
const jsonUrlMarion =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Baza%20danych%20-%20Marion.json';
const pdfUrl =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/katalog%20(59)_compressed.pdf';
const producerPdfMap = {
  alka:
    'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Alka%20-%20Ranking%20Katalog.pdf',
  'best foods':
    'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Best%20Foods%20-%20Ranking%20Katalog.pdf',
  boromir:
    'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Boromir%20-%20Ranking%20Katalog.pdf'
};
const pdfPreviewUrl =
  'https://mozilla.github.io/pdf.js/web/viewer.html?file=';
const MEDIA_STORAGE_GS_PREFIX = 'gs://wf-media-uk-prod/';
const imageBaseUrlRumunia =
  `${MEDIA_STORAGE_GS_PREFIX}World Food/Rumunia/`;
const rumuniaImageManifestUrl =
  'media-rumunia-manifest.json?v=2026-04-17-06-rumunia-manifest';
const imageBaseUrlUkraina =
  `${MEDIA_STORAGE_GS_PREFIX}World Food/Ukraina/`;
const imageBaseUrlMarkaWlasna =
  `${MEDIA_STORAGE_GS_PREFIX}Marka Wlasna/`;
const maspoLogoUrl =
  'https://firebasestorage.googleapis.com/v0/b/pdf-creator-f7a8b.firebasestorage.app/o/szablony%20maspo%2Fmaspo%20logo.png?alt=media&token=a5786f0b-19a7-4c01-a5bb-4b8d0cc9f583';
let currentImageBaseUrl = imageBaseUrlRumunia;
window.imageBaseUrl = currentImageBaseUrl;
window.imageBaseUrlRumunia = imageBaseUrlRumunia;
window.imageBaseUrlUkraina = imageBaseUrlUkraina;
window.imageBaseUrlMarkaWlasna = imageBaseUrlMarkaWlasna;

const firebaseConfig = {
  apiKey: 'AIzaSyCs59YeEZLvlYUio1QLPIZ1vZs3susOqEM',
  authDomain: 'maspo-reports.firebaseapp.com',
  projectId: 'maspo-reports',
  storageBucket: 'maspo-reports.firebasestorage.app',
  messagingSenderId: '392953712563',
  appId: '1:392953712563:web:f9c39997fab1af0d5a05c6',
  measurementId: 'G-W7LKC5BECX'
};

const USERS_COLLECTION = 'app_users';
const LEGACY_SALES_REPORTS_ROOT = 'raporty - maspo';
const SALES_REPORTS_ROOT = 'Raporty Maspo';
const SALES_REPORTS_REP_ROOT = `${SALES_REPORTS_ROOT}/Przedstawiciel handlowy`;
const SALES_REPORTS_ADMIN_ROOT = `${SALES_REPORTS_ROOT}/Administrator`;
const SALES_REPORTS_WEEKLY_REP_ROOT = `${SALES_REPORTS_REP_ROOT}/Sprzedaz tygodniowa`;
const SALES_REPORTS_WEEKLY_ADMIN_ROOT = `${SALES_REPORTS_ADMIN_ROOT}/Sprzedaz tygodniowa`;
const SALES_REPORTS_INDEX_REP_ROOT = `${SALES_REPORTS_REP_ROOT}/Sprzedaz per indeks`;
const SALES_REPORTS_INDEX_ADMIN_ROOT = `${SALES_REPORTS_ADMIN_ROOT}/Sprzedaz per indeks`;
const ADMIN_PRICE_LIST_STORAGE_PATH = `${SALES_REPORTS_ADMIN_ROOT}/Konfiguracja/cennik-indeksow.xlsx`;
const REPORTS_STORAGE_BUCKET = 'gs://wf-reports-uk-prod';
const MEDIA_STORAGE_BUCKET = 'gs://wf-media-uk-prod';
const MEDIA_SHARED_DOWNLOAD_TOKEN_BY_PREFIX = {
  'World Food/': '870dac9b-4d34-4d10-8622-920574e5c8f2',
  'Marka Wlasna/': '870dac9b-4d34-4d10-8622-920574e5c8f2'
};
const FUNCTIONS_REGION = 'europe-west2';
const STORAGE_LIST_TIMEOUT_MS = 8000;
const PERSONAL_SALES_USERS = {
  '9MB4MF77sRayF6jjgvLCMpphBJM2': {
    displayName: 'Robert Bubula',
    storageFolder: 'Bubula Robert'
  }
};

const authSplash = document.getElementById('auth-splash');
const authSplashMessage = document.getElementById('auth-splash-message');
const loginView = document.getElementById('login-view');
const appView = document.getElementById('app-view');
const loginEmail = document.getElementById('login-email');
const loginPassword = document.getElementById('login-password');
const loginBtn = document.getElementById('login-btn');
const logoutBtn = document.getElementById('logout-btn');
const loginError = document.getElementById('login-error');
const personalSalesInsight = document.getElementById('personal-sales-insight');
const personalSalesInsightContent = document.getElementById('personal-sales-insight-content');
const securityModal = document.getElementById('security-modal');
const modalClose = document.getElementById('modal-close');
const catalogModal = document.getElementById('catalog-modal');
const catalogNameInput = document.getElementById('catalog-name');
const catalogCoverInput = document.getElementById('catalog-cover');
const catalogCurrencySelect = document.getElementById('catalog-currency');
const catalogPriceColorInput = document.getElementById('catalog-price-color');
const catalogAdminPriceTools = document.getElementById('catalog-admin-price-tools');
const catalogExcelInput = document.getElementById('catalog-excel');
const catalogIndexColSelect = document.getElementById('catalog-index-col');
const catalogPriceColSelect = document.getElementById('catalog-price-col');
const catalogSaveBtn = document.getElementById('catalog-save-btn');
const catalogCancelBtn = document.getElementById('catalog-cancel-btn');
const catalogError = document.getElementById('catalog-error');
const priceWarningModal = document.getElementById('price-warning-modal');
const priceWarningClose = document.getElementById('price-warning-close');
const passwordModal = document.getElementById('password-modal');
const passwordInput = document.getElementById('password-input');
const passwordConfirm = document.getElementById('password-confirm');
const passwordCancel = document.getElementById('password-cancel');
const passwordError = document.getElementById('password-error');
const adminPanelBtn = document.getElementById('admin-panel-btn');
const appBrandLink = document.querySelector('#app-view .brand-link');
const worldFoodBtn = document.getElementById('world-food-btn');
const markaWlasnaBtn = document.getElementById('marka-wlasna-btn');
const salesReportsBtn = document.getElementById('sales-reports-btn');
const tasksBtn = document.getElementById('tasks-btn');
const worldFoodSubnav = document.getElementById('world-food-subnav');
const reportsSubnav = document.getElementById('reports-subnav');

const slodyczeContainer = document.getElementById('slodycze-content');
const miesoContainer = document.getElementById('mieso-wedliny-content');
const nabialContainer = document.getElementById('nabial-content');
const napojeContainer = document.getElementById('napoje-content');
const przyprawyProszekContainer = document.getElementById('przyprawy-proszek-content');
const puszkiSloikiContainer = document.getElementById('puszki-sloiki-content');
const produktyPodstawoweContainer = document.getElementById('produkty-podstawowe-content');
const kawyHerbatyContainer = document.getElementById('kawy-herbaty-content');
const topRumuniaContainer = document.getElementById('top-rumunia-content');
const importDaneContainer = document.getElementById('import-dane-content');
const reportsContainer = document.getElementById('reports-content');
const ukrainaSlodyczeContainer = document.getElementById('ukraina-slodycze-content');
const ukrainaMiesoContainer = document.getElementById('ukraina-mieso-wedliny-content');
const ukrainaKawyContainer = document.getElementById('ukraina-kawy-herbaty-content');
const ukrainaPuszkiContainer = document.getElementById('ukraina-puszki-sloiki-content');
const ukrainaNapojeContainer = document.getElementById('ukraina-napoje-content');
const ukrainaPrzyprawyContainer = document.getElementById('ukraina-przyprawy-proszek-content');
const ukrainaDodatkiContainer = document.getElementById('ukraina-dodatki-do-potraw-content');
const ukrainaTopContainer = document.getElementById('ukraina-top-ukraina-content');
const importDaneCard = document.getElementById('import-dane');
const reportsCard = document.getElementById('reports-card');
const MARKA_WLASNA_CARDS_CONFIG_PATH = 'Marka Wlasna/cards.json';
const MARKA_WLASNA_CARDS_FALLBACK_URL = 'data/marka-wlasna/cards.json?v=2026-04-20-03';
const MARKA_WLASNA_PUBLIC_DOWNLOAD_TOKEN = 'b6f35d52-9c87-4f3d-a1a1-8f5df0b45011';

function getMarkaWlasnaWedlinyContainer(){
  return getMarkaWlasnaCardContainer('wedliny');
}

function getMarkaWlasnaCardContainer(slug){
  return window.MARKA_WLASNA_LAYOUT?.getCardContainer?.(slug) || null;
}

var fullData = [];
let fullDataColumns = null;
let expanded = false;
let viewMode = 'table';
let activeContainer = slodyczeContainer;
let currentCategoryName = 'Slodycze';
let currentCategorySlug = 'slodycze';
let selectedProducts = new Map();
let filters = {
  producer: '',
  group: '',
  name: '',
  ean: '',
  index: '',
  limit: ''
};
let catalogBlobUrl = null;
let catalogLoading = false;
let catalogCoverDataUrl = null;
let markaWlasnaCardsConfigCache = null;
let catalogPriceMap = null;
let catalogPriceRows = null;
let adminCatalogPriceMapCache = null;
let adminCatalogPriceRowsCache = null;
let adminCatalogPriceLoadPromise = null;
let catalogOptionsOverride = null;
const LIMIT = 25;
const barcodeCache = new Map();
let importedIndexSet = null;
let importedIndexCount = 0;
let importedIndexFile = '';
let listingResults = [];
let listingResultsMap = new Map();
let listingAllDataCache = null;
let listingScannedCodes = new Set();
let listingCooldown = false;
let listingMarionCache = null;
let importDaneCache = null;
let reportsSourceRows = [];
let reportsGeneratedRows = [];
let reportsImportFile = '';
let reportsSummary = [];
let reportsMode = 'top-sales';
let clientReportSourceRows = [];
let clientReportImportFile = '';
let clientReportSelectedRepresentative = '';
let clientReportSelectedCustomer = '';
let clientReportPurchasedRows = [];
let clientReportMissingRows = [];
let clientReportSummary = [];
let clientReportSelectedMissingIndexes = new Set();
let clientReportPurchasedFilters = {
  producer: '',
  ranking: '',
  group: ''
};
let clientReportMissingFilters = {
  producer: '',
  ranking: '',
  group: ''
};
let clientRecommendationIncludeCover = true;
let rumuniaImageManifestPromise = null;
let detailedSalesCustomersCache = null;
let detailedSalesTopPurchasedRows = [];
let detailedSalesTopMissingRows = [];
let detailedSalesTopSummary = null;
let detailedSalesTopComparisonLoading = false;
let detailedSalesTopComparisonError = '';
let detailedSalesTopComparisonRequestId = 0;
let detailedSalesTopPurchasedFilters = { producer:'', group:'' };
let detailedSalesTopMissingFilters = { producer:'', group:'' };
let detailedSalesTopGroupSummaryRows = [];
let detailedSalesTopSelectedGroupKey = '';
let detailedSalesTopProducerSummaryRows = [];
let detailedSalesTopSelectedProducerKey = '';
let detailedSalesTopSelectedMissingKeys = new Set();
let detailedSalesTopSection = 'summary';
let detailedSalesTopCustomerSummaryRows = [];
let detailedSalesTopCustomerSearch = '';
let detailedSalesTopCustomerSort = 'purchased-desc';
let detailedSalesTopCustomerToneFilter = '';
let detailedSalesTopCustomerPortfolioExpanded = false;
let detailedSalesComparisonCustomerSearch = '';
let topSuggestionSourceRows = [];
let topSuggestionGeneratedRows = [];
let topSuggestionPurchasedRows = [];
let topSuggestionImportFile = '';
let topSuggestionSelectedRepresentative = '';
let topSuggestionSelectedCustomer = '';
let topSuggestionLimit = '';
let topSuggestionSummary = null;
let topRumuniaOfferCache = null;
let reportOfferRowsCache = new Map();
let isLoading = false;
let reportsBackgroundPrefetchStarted = false;
const IMAGE_PLACEHOLDER_SRC = 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==';
const IMAGE_EXTENSION_CACHE_KEY = 'wf-image-extension-cache-v2';
const MAX_IMAGE_EXTENSION_CACHE_ENTRIES = 3000;
const SUPPORTED_IMAGE_EXTENSIONS = ['webp', 'png', 'jpg', 'jpeg'];
const imageExtensionCache = createImageExtensionCache();
const privateImageUrlCache = new Map();
let lazyImageObserver = null;

let auth = null;
let db = null;
let storage = null;
let mediaStorage = null;
let functionsService = null;
let authReady = false;
let lastLoginEmail = '';
let lastLoginPassword = '';
let authStateVersion = 0;
let personalSalesDetailsLoading = false;
let phReportsAutoLoading = false;
let currentUserIsAdmin = false;
let personalSalesConfigCache = new Map();
let personalSalesFolderEntriesCache = new Map();
let adminSalesFolderEntriesCache = new Map();
let adminAssignedSalesConfigsCache = null;
let weeklySalesParsedReportCache = new Map();
let weeklySalesParsedReportPromiseCache = new Map();

function clearDetailedSalesTopComparisonResultState(options = {}){
  const preserveCustomerSummaryRows = Boolean(options.preserveCustomerSummaryRows);
  detailedSalesTopPurchasedRows = [];
  detailedSalesTopMissingRows = [];
  detailedSalesTopSummary = null;
  detailedSalesTopComparisonError = '';
  detailedSalesTopPurchasedFilters = { producer:'', group:'' };
  detailedSalesTopMissingFilters = { producer:'', group:'' };
  detailedSalesTopGroupSummaryRows = [];
  detailedSalesTopSelectedGroupKey = '';
  detailedSalesTopProducerSummaryRows = [];
  detailedSalesTopSelectedProducerKey = '';
  detailedSalesTopSelectedMissingKeys = new Set();
  detailedSalesTopSection = 'summary';
  if(!preserveCustomerSummaryRows){
    detailedSalesTopCustomerSummaryRows = [];
  }
}

function resetDetailedSalesTopCustomerPortfolioFilters(){
  detailedSalesTopCustomerSearch = '';
  detailedSalesTopCustomerSort = 'purchased-desc';
  detailedSalesTopCustomerToneFilter = '';
  detailedSalesTopCustomerPortfolioExpanded = false;
}

function getFilteredDetailedSalesComparisonCustomers(customers, query = detailedSalesComparisonCustomerSearch){
  const normalizedQuery = String(query || '').trim().toLowerCase();
  if(!normalizedQuery) return Array.isArray(customers) ? customers : [];
  return (Array.isArray(customers) ? customers : []).filter(customer => {
    const label = [customer?.code, customer?.name].filter(Boolean).join(' ').toLowerCase();
    return label.includes(normalizedQuery);
  });
}

function buildDetailedSalesComparisonCustomerOptionsHtml(customers, query = detailedSalesComparisonCustomerSearch){
  const visibleCustomers = getFilteredDetailedSalesComparisonCustomers(customers, query);
  const placeholder = customers.length
    ? (visibleCustomers.length ? 'Wybierz klienta do porównania' : 'Brak klientów dla tego wyszukiwania')
    : 'Brak klientów w raporcie';
  return `
    <option value="">${escapeHtml(placeholder)}</option>
    ${visibleCustomers.map(customer => {
      const value = `${customer.code}|||${customer.name}`;
      const label = [customer.code, customer.name].filter(Boolean).join(' - ') || 'Bez nazwy klienta';
      return `<option value="${escapeAttr(value)}" ${detailedSalesComparisonCustomer === value ? 'selected' : ''}>${escapeHtml(label)}</option>`;
    }).join('')}
  `;
}

function getDetailedSalesComparisonCustomerValue(customer){
  return `${customer?.code || ''}|||${customer?.name || ''}`;
}

function resolveDetailedSalesComparisonCustomerAutoMatch(customers, query = detailedSalesComparisonCustomerSearch){
  const normalizedQuery = normalizeCustomerCodeValue(query);
  if(!normalizedQuery){
    return '';
  }

  const customerList = Array.isArray(customers) ? customers : [];
  const exactCodeMatch = customerList.find(customer => normalizeCustomerCodeValue(customer?.code) === normalizedQuery);
  if(exactCodeMatch){
    return getDetailedSalesComparisonCustomerValue(exactCodeMatch);
  }

  const codePrefixMatches = customerList.filter(customer => (
    normalizeCustomerCodeValue(customer?.code).startsWith(normalizedQuery)
  ));
  if(codePrefixMatches.length === 1){
    return getDetailedSalesComparisonCustomerValue(codePrefixMatches[0]);
  }

  const visibleCustomers = getFilteredDetailedSalesComparisonCustomers(customerList, query);
  if(visibleCustomers.length === 1){
    return getDetailedSalesComparisonCustomerValue(visibleCustomers[0]);
  }

  return '';
}

function syncDetailedSalesComparisonCustomerControls(){
  if(typeof document === 'undefined') return;
  const inputs = Array.from(document.querySelectorAll('[data-detailed-sales-comparison-search]'));
  const selects = Array.from(document.querySelectorAll('[data-detailed-sales-comparison-select]'));
  if(!inputs.length && !selects.length) return;
  const customers = typeof getDetailedSalesCustomers === 'function' ? getDetailedSalesCustomers() : [];
  const optionsHtml = buildDetailedSalesComparisonCustomerOptionsHtml(customers);
  const visibleCustomers = getFilteredDetailedSalesComparisonCustomers(customers);
  const previewValue = visibleCustomers[0]
    ? `${visibleCustomers[0].code}|||${visibleCustomers[0].name}`
    : '';

  inputs.forEach(input => {
    input.disabled = !customers.length;
    if(input.value !== detailedSalesComparisonCustomerSearch){
      input.value = detailedSalesComparisonCustomerSearch;
    }
  });

  selects.forEach(select => {
    const previousValue = select.value;
    select.disabled = !customers.length;
    select.innerHTML = optionsHtml;
    if(detailedSalesComparisonCustomer && (!detailedSalesComparisonCustomerSearch || visibleCustomers.some(customer => `${customer.code}|||${customer.name}` === detailedSalesComparisonCustomer))){
      select.value = detailedSalesComparisonCustomer;
    }else if(previewValue){
      select.value = previewValue;
    }else if(previousValue){
      select.value = previousValue;
    }
  });
}

const REPORT_GROUP_CONFIGS = [
  {
    id: 'slodycze',
    label: 'SLODYCZE I PRZEKASKI RUMUNIA - SLODYCZE RUMUNIA',
    dashboardLabel: 'SŁODYCZE RUMUNIA',
    groups: ['SŁODYCZE I PRZEKĄSKI RUMUNIA', 'SŁODYCZE RUMUNIA', 'PRZEKĄSKI RUMUNIA'],
    sources: [jsonUrl]
  },
  {
    id: 'puszki',
    label: 'PUSZKI I SLOIKI CHLODZONE RUMUNIA - PUSZKI I SLOIKI RUMUNIA',
    dashboardLabel: 'PUSZKI I SŁOIKI RUMUNIA',
    groups: ['PUSZKI I SŁOIKI CHŁODZONE RUMUNIA', 'PUSZKI I SŁOIKI RUMUNIA', 'SŁOIKI RUMUNIA'],
    sources: [jsonUrlPuszkiSloiki]
  },
  {
    id: 'kawa',
    label: 'HERBATA I KAWA RUMUNIA',
    dashboardLabel: 'HERBATA I KAWA RUMUNIA',
    groups: ['HERBATA I KAWA RUMUNIA'],
    sources: [jsonUrlKawyHerbaty]
  },
  {
    id: 'mieso',
    label: 'MIESO I WEDLINY RUMUNIA',
    dashboardLabel: 'MIĘSO I WĘDLINY RUMUNIA',
    groups: ['MIĘSO I WĘDLINY RUMUNIA'],
    sources: [jsonUrlMieso]
  },
  {
    id: 'nabial',
    label: 'NABIAL RUMUNIA - RUMUNIA CHLODNIA POZOSTALE',
    dashboardLabel: 'NABIAŁ RUMUNIA',
    groups: ['NABIAŁ RUMUNIA', 'RUMUNIA CHŁODNIA POZOSTAŁE'],
    sources: [jsonUrlNabial]
  },
  {
    id: 'napoje',
    label: 'NAPOJE BEZALKOHOLOWE RUMUNIA',
    dashboardLabel: 'NAPOJE BEZALKOHOLOWE RUMUNIA',
    groups: ['NAPOJE BEZALKOHOLOWE RUMUNIA'],
    sources: [jsonUrlNapoje]
  },
  {
    id: 'podstawowe',
    label: 'PRODUKTY PODSTAWOWE RUMUNIA',
    dashboardLabel: 'PRODUKTY PODSTAWOWE RUMUNIA',
    groups: ['PRODUKTY PODSTAWOWE RUMUNIA'],
    sources: [jsonUrlProduktyPodstawowe]
  },
  {
    id: 'przyprawy',
    label: 'PRZYPRAWY I DODATKI W PROSZKU RUMUNIA',
    dashboardLabel: 'PRZYPRAWY I DODATKI W PROSZKU RUMUNIA',
    groups: ['PRZYPRAWY I DODATKI W PROSZKU RUMUNIA', 'DODATKI DO POTRAW RUMUNIA'],
    sources: [jsonUrlPrzyprawyProszek]
  }
];

const REPORT_GROUP_SHORT_LABELS = {
  slodycze: 'Slodycze',
  puszki: 'Puszki',
  kawa: 'Kawa',
  mieso: 'Mieso',
  nabial: 'Nabial',
  napoje: 'Napoje',
  podstawowe: 'Podstawowe',
  przyprawy: 'Przyprawy'
};

function createDefaultReportLimits(){
  return REPORT_GROUP_CONFIGS.reduce((acc, group) => {
    acc[group.id] = 25;
    return acc;
  }, {});
}

let reportsGroupLimits = createDefaultReportLimits();

function setWorldFoodExpanded(expanded){
  if(worldFoodBtn){
    worldFoodBtn.classList.toggle('active', expanded);
    worldFoodBtn.setAttribute('aria-expanded', expanded ? 'true' : 'false');
  }
  if(worldFoodSubnav){
    worldFoodSubnav.classList.toggle('hidden', !expanded);
    worldFoodSubnav.setAttribute('aria-hidden', expanded ? 'false' : 'true');
  }
}

function setPrimaryNavState(mode){
  const isWorldFood = mode === 'world-food';
  const isMarkaWlasna = mode === 'marka-wlasna';
  setWorldFoodExpanded(isWorldFood);
  if(markaWlasnaBtn){
    markaWlasnaBtn.classList.toggle('active', isMarkaWlasna);
  }
  if(salesReportsBtn){
    salesReportsBtn.classList.toggle('active', mode === 'reports');
  }
  if(tasksBtn){
    tasksBtn.classList.toggle('active', mode === 'tasks');
  }
  if(reportsSubnav){
    reportsSubnav.classList.toggle('hidden', mode !== 'reports');
    reportsSubnav.setAttribute('aria-hidden', mode === 'reports' ? 'false' : 'true');
  }
}

function showTabSection(sectionId){
  document.querySelectorAll('.tab-content').forEach(section => section.classList.add('hidden'));
  const section = document.getElementById(sectionId);
  if(section){
    section.classList.remove('hidden');
  }
}

function activateTab(tabId){
  document.querySelectorAll('.tab').forEach(tab => {
    tab.classList.toggle('active', tab.dataset.tab === tabId);
  });
  setPrimaryNavState('world-food');
  showTabSection(tabId);
  currentImageBaseUrl = tabId === 'ukraina' ? imageBaseUrlUkraina : imageBaseUrlRumunia;
  window.imageBaseUrl = currentImageBaseUrl;
  clearAllContentContainers();
  document.querySelectorAll('.grid .card').forEach(card => card.classList.remove('active'));
  viewMode = 'table';

  if(tabId !== 'listing'){
    stopListingScanner();
  }
}

function getMarkaWlasnaCardConfig(slug){
  return window.MARKA_WLASNA_LAYOUT?.getCardBySlug?.(slug) || null;
}

function getMarkaWlasnaCardSource(slug){
  const card = getMarkaWlasnaCardConfig(slug);
  if(!card){
    return {
      name: 'Marka Wlasna',
      type: 'json',
      url: ''
    };
  }

  return {
    name: `Marka Wlasna - ${card.title}`,
    type: card.dataSourceType || 'storage-json',
    path: card.dataSourcePath || '',
    fallbackUrl: card.fallbackUrl || buildMarkaWlasnaPublicUrl(card.dataSourcePath || '')
  };
}

function buildMarkaWlasnaPublicUrl(path){
  const normalizedPath = String(path || '').trim().replace(/^\/+/, '');
  if(!normalizedPath) return '';
  return `https://firebasestorage.googleapis.com/v0/b/wf-reports-uk-prod/o/${encodeURIComponent(normalizedPath)}?alt=media&token=${encodeURIComponent(MARKA_WLASNA_PUBLIC_DOWNLOAD_TOKEN)}`;
}

function decorateMarkaWlasnaCards(cards){
  return (Array.isArray(cards) ? cards : []).map(card => {
    const dataSourcePath = String(card?.dataSourcePath || '').trim();
    const iconPath = String(card?.iconPath || '').trim();
    const fallbackUrl = String(card?.fallbackUrl || '').trim() || buildMarkaWlasnaPublicUrl(dataSourcePath);
    const iconUrl = String(card?.iconUrl || '').trim() || buildMarkaWlasnaPublicUrl(iconPath);

    return {
      ...(card && typeof card === 'object' ? card : {}),
      dataSourcePath,
      fallbackUrl,
      iconPath,
      iconUrl
    };
  });
}

async function ensureMarkaWlasnaCardsLoaded(forceRefresh = false){
  if(markaWlasnaCardsConfigCache && !forceRefresh){
    return markaWlasnaCardsConfigCache;
  }

  let cards = null;
  const publicCardsUrl = buildMarkaWlasnaPublicUrl(MARKA_WLASNA_CARDS_CONFIG_PATH);

  if(publicCardsUrl){
    try{
      cards = await loadJsonRowsFromUrl(publicCardsUrl);
    }catch(error){
      console.warn('Marka Wlasna cards public url fallback', error);
    }
  }

  if(!Array.isArray(cards) || !cards.length){
    try{
      cards = await loadJsonRowsFromUrl(MARKA_WLASNA_CARDS_FALLBACK_URL);
    }catch(error){
      console.error('Marka Wlasna cards fallback error', error);
      cards = [];
    }
  }

  const normalizedCards = window.MARKA_WLASNA_LAYOUT?.normalizeCards
    ? window.MARKA_WLASNA_LAYOUT.normalizeCards(cards)
    : (Array.isArray(cards) ? cards : []);
  const decoratedCards = decorateMarkaWlasnaCards(normalizedCards);

  markaWlasnaCardsConfigCache = decoratedCards;
  if(window.MARKA_WLASNA_LAYOUT?.setCards){
    window.MARKA_WLASNA_LAYOUT.setCards(decoratedCards);
  }else if(window.MARKA_WLASNA_LAYOUT?.ensureLayout){
    window.MARKA_WLASNA_LAYOUT.ensureLayout(decoratedCards);
  }

  return decoratedCards;
}

async function showMarkaWlasnaView(){
  if(window.MARKA_WLASNA_LAYOUT?.ensureLayout){
    window.MARKA_WLASNA_LAYOUT.ensureLayout();
  }

  const markaWlasnaSection = document.getElementById('marka-wlasna');
  if(markaWlasnaSection && !window.__wfMarkaWlasnaDelegatedBound){
    window.__wfMarkaWlasnaDelegatedBound = true;
    markaWlasnaSection.addEventListener('click', event => {
      const card = event.target.closest('.marka-wlasna-card[data-card-slug]');
      if(!card) return;
      void openMarkaWlasnaCard(card.dataset.cardSlug || '');
    });
  }
  setPrimaryNavState('marka-wlasna');
  showTabSection('marka-wlasna');
  setImageBase(imageBaseUrlMarkaWlasna);
  setFullData([]);
  isLoading = false;
  activeContainer = document.getElementById('marka-wlasna');
  currentCategoryName = 'Marka_Wlasna';
  currentCategorySlug = 'marka_wlasna';
  resetFilters();
  clearAllContentContainers();
  document.querySelectorAll('.grid .card').forEach(card => card.classList.remove('active'));
  stopListingScanner();
  await ensureMarkaWlasnaCardsLoaded(true);
}

function showTasksView(){
  setPrimaryNavState('tasks');
  showTabSection('tasks');
  setFullData([]);
  isLoading = false;
  activeContainer = document.getElementById('tasks');
  currentCategoryName = 'Zadania';
  currentCategorySlug = 'zadania';
  resetFilters();
  clearAllContentContainers();
  document.querySelectorAll('.grid .card').forEach(card => card.classList.remove('active'));
  stopListingScanner();
}

function setAuthBootState(isLoading, message = 'Sprawdzam sesję...'){
  if(authSplashMessage){
    authSplashMessage.textContent = message || 'Sprawdzam sesję...';
  }
  if(document.body){
    document.body.classList.toggle('auth-booting', Boolean(isLoading));
  }
  if(authSplash){
    authSplash.classList.toggle('hidden', !isLoading);
  }
}

function isQuickOrderPublicMode(){
  return Boolean(window.__WF_QUICK_ORDER_BOOT__?.isPublicMode);
}

function showApp(){
  if(isQuickOrderPublicMode()){
    setAuthBootState(false);
    if(loginView) loginView.classList.add('hidden');
    if(appView) appView.classList.add('hidden');
    return;
  }
  setAuthBootState(false);
  if(loginView) loginView.classList.add('hidden');
  if(appView) appView.classList.remove('hidden');
  showSecurityModalOnce();
}

function showLogin(){
  if(isQuickOrderPublicMode()){
    setAuthBootState(false);
    if(loginView) loginView.classList.add('hidden');
    if(appView) appView.classList.add('hidden');
    return;
  }
  setAuthBootState(false);
  if(appView) appView.classList.add('hidden');
  if(loginView) loginView.classList.remove('hidden');
}

function showSecurityModalOnce(){
  if(!securityModal) return;
  securityModal.classList.remove('hidden');
}

function resetAppState(){
  fullData = [];
  fullDataColumns = null;
  isLoading = false;
  expanded = false;
  viewMode = 'table';
  activeContainer = slodyczeContainer;
  currentCategoryName = 'Slodycze';
  currentCategorySlug = 'slodycze';
  selectedProducts = new Map();
  catalogBlobUrl = null;
  catalogLoading = false;
  catalogCoverDataUrl = null;
  catalogPriceMap = null;
  catalogPriceRows = null;
  adminCatalogPriceMapCache = null;
  adminCatalogPriceRowsCache = null;
  adminCatalogPriceLoadPromise = null;
  catalogOptionsOverride = null;
  markaWlasnaCardsConfigCache = null;
  importedIndexSet = null;
  importedIndexCount = 0;
  importedIndexFile = '';
  reportsSourceRows = [];
  reportsGeneratedRows = [];
  reportsImportFile = '';
  reportsSummary = [];
  reportsMode = 'top-sales';
  clientReportSourceRows = [];
  clientReportImportFile = '';
  clientReportSelectedRepresentative = '';
  clientReportSelectedCustomer = '';
  clientReportPurchasedRows = [];
  clientReportMissingRows = [];
  clientReportSummary = [];
  clientReportSelectedMissingIndexes = new Set();
  clientReportPurchasedFilters = { producer:'', ranking:'', group:'' };
  clientReportMissingFilters = { producer:'', ranking:'', group:'' };
  clientRecommendationIncludeCover = true;
  weeklySalesSourceRows = [];
  weeklySalesGeneratedRows = [];
  weeklySalesImportFile = '';
  weeklySalesSelectedRepresentative = '';
  weeklySalesSelectedCustomer = '';
  weeklySalesSelectedProducer = '';
  weeklySalesSelectedWeek = '';
  weeklySalesComparisonWindow = 1;
  weeklySalesOnlyLastWeek250 = false;
  weeklySalesRepComparison = false;
  weeklySalesSortOrder = 'sales-desc';
  weeklySalesErrorMessage = '';
  detailedSalesSourceRows = [];
  detailedSalesGeneratedRows = [];
  detailedSalesImportFile = '';
  detailedSalesCustomersCache = null;
  detailedSalesSelectedCustomer = '';
  detailedSalesComparisonCustomer = '';
  detailedSalesComparisonCustomerSearch = '';
  detailedSalesSelectedGroups = [];
  detailedSalesSortOrder = 'sales-desc';
  detailedSalesWeeksLimit = 1;
  detailedSalesLoadedWeeksLimit = 0;
  detailedSalesLoadedReportsCount = 0;
  detailedSalesAvailableReportsCount = 0;
  detailedSalesUpdatedAt = '';
  detailedSalesLoading = false;
  detailedSalesErrorMessage = '';
  clearDetailedSalesTopComparisonResultState();
  detailedSalesTopComparisonLoading = false;
  detailedSalesTopComparisonRequestId = 0;
  resetDetailedSalesTopCustomerPortfolioFilters();
  topSuggestionSourceRows = [];
  topSuggestionGeneratedRows = [];
  topSuggestionPurchasedRows = [];
  topSuggestionImportFile = '';
  topSuggestionSelectedRepresentative = '';
  topSuggestionSelectedCustomer = '';
  topSuggestionLimit = '';
  topSuggestionSummary = null;
  if(typeof adminWeeklySalesViewLoading !== 'undefined'){
    adminWeeklySalesViewLoading = false;
  }
  if(typeof adminWeeklySalesInitialized !== 'undefined'){
    adminWeeklySalesInitialized = false;
  }
  if(typeof adminWeeklySalesFocusedRepresentative !== 'undefined'){
    adminWeeklySalesFocusedRepresentative = '';
  }
  personalSalesConfigCache = new Map();
  personalSalesFolderEntriesCache = new Map();
  adminSalesFolderEntriesCache = new Map();
  adminAssignedSalesConfigsCache = null;
  weeklySalesParsedReportCache = new Map();
  weeklySalesParsedReportPromiseCache = new Map();
  detailedSalesParsedReportCache = new Map();
  detailedSalesParsedReportPromiseCache = new Map();
  detailedSalesGroupsDropdownOpen = false;
  reportsGroupLimits = createDefaultReportLimits();
  resetFilters();

  slodyczeContainer.innerHTML = '';
  miesoContainer.innerHTML = '';
  nabialContainer.innerHTML = '';
  napojeContainer.innerHTML = '';
  przyprawyProszekContainer.innerHTML = '';
  puszkiSloikiContainer.innerHTML = '';
  produktyPodstawoweContainer.innerHTML = '';
  kawyHerbatyContainer.innerHTML = '';
  if(topRumuniaContainer) topRumuniaContainer.innerHTML = '';
  if(importDaneContainer) importDaneContainer.innerHTML = '';
  if(reportsContainer) reportsContainer.innerHTML = '';
  if(ukrainaSlodyczeContainer) ukrainaSlodyczeContainer.innerHTML = '';
  if(ukrainaMiesoContainer) ukrainaMiesoContainer.innerHTML = '';
  if(ukrainaKawyContainer) ukrainaKawyContainer.innerHTML = '';
  if(ukrainaPuszkiContainer) ukrainaPuszkiContainer.innerHTML = '';
  if(ukrainaNapojeContainer) ukrainaNapojeContainer.innerHTML = '';
  if(ukrainaPrzyprawyContainer) ukrainaPrzyprawyContainer.innerHTML = '';
  if(ukrainaDodatkiContainer) ukrainaDodatkiContainer.innerHTML = '';
  if(ukrainaTopContainer) ukrainaTopContainer.innerHTML = '';
  if(window.MARKA_WLASNA_LAYOUT?.clearContents){
    window.MARKA_WLASNA_LAYOUT.clearContents();
  }

  document.querySelectorAll('.grid .card').forEach(c => c.classList.remove('active'));

  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  const defaultTab = document.querySelector('.tab[data-tab="rumunia"]');
  if(defaultTab) defaultTab.classList.add('active');
  setPrimaryNavState('world-food');

  document.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));
  const rumuniaSection = document.getElementById('rumunia');
  if(rumuniaSection) rumuniaSection.classList.remove('hidden');
  currentImageBaseUrl = imageBaseUrlRumunia;
  window.imageBaseUrl = currentImageBaseUrl;
  reportsBackgroundPrefetchStarted = false;
}

function clearAllContentContainers(){
  slodyczeContainer.innerHTML = '';
  miesoContainer.innerHTML = '';
  nabialContainer.innerHTML = '';
  napojeContainer.innerHTML = '';
  przyprawyProszekContainer.innerHTML = '';
  puszkiSloikiContainer.innerHTML = '';
  produktyPodstawoweContainer.innerHTML = '';
  kawyHerbatyContainer.innerHTML = '';
  if(topRumuniaContainer) topRumuniaContainer.innerHTML = '';
  if(importDaneContainer) importDaneContainer.innerHTML = '';
  if(reportsContainer) reportsContainer.innerHTML = '';
  if(ukrainaSlodyczeContainer) ukrainaSlodyczeContainer.innerHTML = '';
  if(ukrainaMiesoContainer) ukrainaMiesoContainer.innerHTML = '';
  if(ukrainaKawyContainer) ukrainaKawyContainer.innerHTML = '';
  if(ukrainaPuszkiContainer) ukrainaPuszkiContainer.innerHTML = '';
  if(ukrainaNapojeContainer) ukrainaNapojeContainer.innerHTML = '';
  if(ukrainaPrzyprawyContainer) ukrainaPrzyprawyContainer.innerHTML = '';
  if(ukrainaDodatkiContainer) ukrainaDodatkiContainer.innerHTML = '';
  if(ukrainaTopContainer) ukrainaTopContainer.innerHTML = '';
  if(window.MARKA_WLASNA_LAYOUT?.clearContents){
    window.MARKA_WLASNA_LAYOUT.clearContents();
  }
}

function setImageBase(base){
  currentImageBaseUrl = base;
  window.imageBaseUrl = currentImageBaseUrl;
}

function isConfiguredDataSource(source){
  return Boolean(String(source?.url || source?.path || source?.storagePath || '').trim());
}

async function loadJsonRowsFromUrl(url){
  const response = await fetch(url, { cache: 'no-store' });
  if(!response.ok){
    throw new Error(`Nie udało się pobrać danych (HTTP ${response.status}).`);
  }
  return response.json();
}

async function loadDataSourceRows(source){
  if(!isConfiguredDataSource(source)){
    throw new Error('Brak pliku źródłowego dla tej kategorii.');
  }
  if(source.type === 'storage-json'){
    try{
      return await loadStorageJson(source.path || source.storagePath);
    }catch(error){
      if(source.fallbackUrl){
        console.warn('Storage JSON fallback', source.path || source.storagePath, error);
        return loadJsonRowsFromUrl(source.fallbackUrl);
      }
      throw error;
    }
  }
  if(source.type === 'storage-xlsx'){
    return loadStorageXlsxAsJson(source.path || source.storagePath);
  }
  if(source.type === 'xlsx'){
    return loadXlsxAsJson(source.url);
  }
  return loadJsonRowsFromUrl(source.url);
}

function prepareCategoryView(container, name, slug){
  setFullData([]);
  isLoading = true;
  expanded = false;
  viewMode = 'table';
  activeContainer = container;
  currentCategoryName = name;
  currentCategorySlug = slug;
  resetFilters();
  clearAllContentContainers();
  render();
}

function setLoginError(message){
  if(loginError) loginError.textContent = message || '';
}

function getFriendlyAuthErrorMessage(error){
  const code = String(error?.code || '').trim();

  if(code === 'auth/network-request-failed'){
    return 'Brak połączenia z Firebase. Sprawdź hotspot, VPN i internet.';
  }
  if(code === 'auth/network-timeout'){
    return 'Logowanie przekroczyło czas oczekiwania. Sprawdź internet i spróbuj ponownie.';
  }
  if(code === 'auth/unauthorized-domain'){
    return 'Ten adres nie jest dozwolony w Firebase Auth (Authorized domains).';
  }
  if(
    code === 'auth/invalid-credential' ||
    code === 'auth/wrong-password' ||
    code === 'auth/user-not-found' ||
    code === 'auth/invalid-email'
  ){
    return 'Nieprawidłowy email lub hasło.';
  }
  if(code === 'auth/too-many-requests'){
    return 'Za dużo prób logowania. Spróbuj ponownie za chwilę.';
  }
  if(code === 'auth/user-disabled'){
    return 'To konto jest wyłączone.';
  }

  return code || error?.message || 'Błąd logowania';
}

function withTimeout(promise, timeoutMs, fallbackValue){
  return Promise.race([
    promise,
    new Promise(resolve => {
      setTimeout(() => resolve(fallbackValue), timeoutMs);
    })
  ]);
}

async function signInWithTimeout(email, password, timeoutMs = 15000){
  if(!auth){
    throw { code: 'auth/internal-error', message: 'Firebase Auth nie jest gotowy.' };
  }

  const signInResult = await Promise.race([
    auth.signInWithEmailAndPassword(email, password),
    new Promise((_, reject) => {
      setTimeout(() => {
        reject({ code: 'auth/network-timeout' });
      }, timeoutMs);
    })
  ]);

  if(typeof signInResult?.user?.getIdToken === 'function'){
    await signInResult.user.getIdToken(true);
  }

  return signInResult;
}

function normalizeUserRole(value, fallback = 'user'){
  return String(value || '').trim().toLowerCase() === 'admin' ? 'admin' : fallback;
}

async function hasAdminAccess(user, options = {}){
  if(!user) return false;

  try{
    const tokenResult = await user.getIdTokenResult(Boolean(options.forceRefresh));
    if(tokenResult?.claims?.admin === true){
      return true;
    }
  }catch(error){
    console.error('Admin claims read error', error);
  }

  try{
    if(db){
      const userDoc = await db.collection(USERS_COLLECTION).doc(user.uid).get();
      if(userDoc.exists && normalizeUserRole(userDoc.data()?.role, '') === 'admin'){
        return true;
      }
    }
  }catch(error){
    console.error('Admin Firestore role read error', error);
  }

  try{
    if(functionsService && typeof functionsService.httpsCallable === 'function'){
      const callable = functionsService.httpsCallable('adminEnsureAccess');
      const result = await callable({});
      return result?.data?.isAdmin === true;
    }
  }catch(error){
    console.error('Admin access ensure error', error);
  }

  return false;
}

function hidePersonalSalesInsight(){
  if(personalSalesInsight) personalSalesInsight.classList.add('hidden');
  if(personalSalesInsightContent) personalSalesInsightContent.innerHTML = '';
}

function syncPersonalSalesInsightCardHeights(){
  if(!personalSalesInsightContent) return;
  const layout = personalSalesInsightContent.querySelector('.sales-insight-layout');
  const leftCard = personalSalesInsightContent.querySelector('.sales-insight-shell');
  const rightCard = personalSalesInsightContent.querySelector('.sales-potential-shell');
  if(!layout || !leftCard || !rightCard) return;

  const leftHeight = Math.ceil(leftCard.getBoundingClientRect().height);
  if(leftHeight > 0){
    rightCard.style.height = `${leftHeight}px`;
  }
}

function setPersonalSalesInsightMarkup(markup){
  if(!personalSalesInsight || !personalSalesInsightContent) return;
  personalSalesInsightContent.innerHTML = markup;
  personalSalesInsight.classList.remove('hidden');
  requestAnimationFrame(() => {
    syncPersonalSalesInsightCardHeights();
    requestAnimationFrame(() => {
      syncPersonalSalesInsightCardHeights();
    });
  });
}

function formatSalesNumber(value){
  return new Intl.NumberFormat('pl-PL', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(Number.isFinite(value) ? value : 0);
}

function formatNumber(value, options = {}){
  const numericValue = Number(value || 0);
  const isInteger = Number.isInteger(numericValue);
  const minimumFractionDigits = Object.prototype.hasOwnProperty.call(options, 'minimumFractionDigits')
    ? options.minimumFractionDigits
    : (isInteger ? 0 : 2);
  const maximumFractionDigits = Object.prototype.hasOwnProperty.call(options, 'maximumFractionDigits')
    ? options.maximumFractionDigits
    : (isInteger ? 0 : 2);
  return new Intl.NumberFormat('pl-PL', {
    minimumFractionDigits,
    maximumFractionDigits
  }).format(Number.isFinite(numericValue) ? numericValue : 0);
}

function formatPercent(value, options = {}){
  return `${formatNumber(Number(value || 0), {
    minimumFractionDigits: Object.prototype.hasOwnProperty.call(options, 'minimumFractionDigits')
      ? options.minimumFractionDigits
      : 1,
    maximumFractionDigits: Object.prototype.hasOwnProperty.call(options, 'maximumFractionDigits')
      ? options.maximumFractionDigits
      : 1
  })}%`;
}

function normalizeHeaderKey(value){
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function findHeaderKey(headers, variants){
  const sourceHeaders = Array.isArray(headers) ? headers : [];
  const lookup = new Map();
  const normalizedHeaders = sourceHeaders.map(header => ({
    original: header,
    normalized: normalizeHeaderKey(header)
  }));

  sourceHeaders.forEach(header => {
    lookup.set(normalizeHeaderKey(header), header);
  });

  for(const variant of variants || []){
    const found = lookup.get(normalizeHeaderKey(variant));
    if(found) return found;
  }

  for(const variant of variants || []){
    const normalizedVariant = normalizeHeaderKey(variant);
    const found = normalizedHeaders.find(header => header.normalized.includes(normalizedVariant));
    if(found?.original) return found.original;
  }

  return '';
}

function formatSalesDelta(value){
  const prefix = value > 0 ? '+' : value < 0 ? '-' : '';
  return `${prefix}${formatSalesNumber(Math.abs(value))}`;
}

function formatPercentageDelta(value){
  const prefix = value > 0 ? '+' : value < 0 ? '-' : '';
  return `${prefix}${Math.abs(value).toFixed(1).replace('.', ',')}%`;
}

function formatInsightDate(value){
  if(!value) return 'brak daty';
  const date = new Date(value);
  if(Number.isNaN(date.getTime())) return 'brak daty';
  return date.toLocaleString('pl-PL');
}

function buildWeeklySalesComparisonMetrics(previousTotal, lastTotal){
  const normalizedPreviousTotal = Number(previousTotal || 0);
  const normalizedLastTotal = Number(lastTotal || 0);
  const difference = normalizedLastTotal - normalizedPreviousTotal;
  const percentChange = normalizedPreviousTotal !== 0
    ? (difference / normalizedPreviousTotal) * 100
    : (normalizedLastTotal === 0 ? 0 : 100);
  const direction = difference > 0.005 ? 'up' : difference < -0.005 ? 'down' : 'flat';

  return {
    previousTotal: normalizedPreviousTotal,
    lastTotal: normalizedLastTotal,
    difference,
    percentChange,
    direction
  };
}

function getSalesInsightStatusMeta(direction, options = {}){
  const isTeamScope = Boolean(options.teamScope);
  if(direction === 'up'){
    return {
      directionClass: 'is-up',
      statusClass: 'sales-insight-status--up',
      statusText: isTeamScope
        ? 'Sprzedaż zespołu rośnie względem poprzedniego tygodnia'
        : 'Sprzedaż rośnie względem poprzedniego tygodnia',
      statusIcon: '↗',
      trendKpiLabel: 'Wzrost'
    };
  }

  if(direction === 'down'){
    return {
      directionClass: 'is-down',
      statusClass: 'sales-insight-status--down',
      statusText: isTeamScope
        ? 'Sprzedaż zespołu spada względem poprzedniego tygodnia'
        : 'Sprzedaż spada względem poprzedniego tygodnia',
      statusIcon: '↘',
      trendKpiLabel: 'Spadek'
    };
  }

  return {
    directionClass: 'is-flat',
    statusClass: 'sales-insight-status--flat',
    statusText: isTeamScope
      ? 'Sprzedaż zespołu utrzymała ten sam poziom'
      : 'Sprzedaż utrzymała ten sam poziom',
    statusIcon: '→',
    trendKpiLabel: 'Stabilnie'
  };
}

function renderAdminSalesInsightLoading(){
  setPersonalSalesInsightMarkup(`
    <div class="sales-insight-layout sales-insight-layout--single">
      <div class="sales-insight-shell sales-insight-shell--admin">
        <div class="sales-insight-main">
          <span class="sales-insight-eyebrow">Panel administratora</span>
          <h2>Sprzedaż wszystkich przedstawicieli</h2>
          <p>Ładuję najnowszy raport tygodniowy administratora i porównuję jego ostatni tydzień z poprzednim.</p>
        </div>
        <div class="sales-insight-side">
          <div class="admin-user-empty">Pobieram raport administratora z Firebase...</div>
        </div>
      </div>
    </div>
  `);
}

function renderAdminSalesInsightEmpty(message){
  setPersonalSalesInsightMarkup(`
    <div class="sales-insight-empty">
      <strong>Panel administratora</strong><br>
      ${escapeHtml(message)}
    </div>
  `);
}

function buildAdminWeeklySalesInsightSummaryFromRows(rows, options = {}){
  const sourceRows = Array.isArray(rows) ? rows : [];
  const availableWeeks = typeof getWeeklySalesAvailableWeeksAsc === 'function'
    ? getWeeklySalesAvailableWeeksAsc(sourceRows)
    : Array.from(new Set(
      sourceRows.map(row => String(row.week || '').trim()).filter(Boolean)
    )).sort((a, b) => String(a || '').localeCompare(String(b || ''), 'pl', { numeric: true }));
  const lastWeek = availableWeeks[availableWeeks.length - 1] || '';
  const previousWeek = availableWeeks[availableWeeks.length - 2] || '';

  if(!lastWeek){
    throw new Error('Raport administratora nie zawiera tygodni do porównania.');
  }

  const representativesMap = new Map();
  sourceRows.forEach(row => {
    const week = String(row.week || '').trim();
    if(week !== lastWeek && week !== previousWeek) return;

    const representative = String(row.representative || '').trim() || 'Bez przedstawiciela';
    const current = representativesMap.get(representative) || {
      representative,
      previousTotal: 0,
      lastTotal: 0,
      previousWeek,
      lastWeek
    };
    const salesValue = Number(row.quantity || row.value || 0);

    if(week === previousWeek){
      current.previousTotal += salesValue;
    }else if(week === lastWeek){
      current.lastTotal += salesValue;
    }

    representativesMap.set(representative, current);
  });

  const representativeRows = Array.from(representativesMap.values())
    .map(row => ({
      ...row,
      ...buildWeeklySalesComparisonMetrics(row.previousTotal, row.lastTotal)
    }))
    .sort((a, b) => b.lastTotal - a.lastTotal || String(a.representative || '').localeCompare(String(b.representative || ''), 'pl'));

  if(!representativeRows.length){
    throw new Error('Raport administratora nie zawiera danych przedstawicieli do wyświetlenia.');
  }

  const previousTotal = representativeRows.reduce((sum, row) => sum + Number(row.previousTotal || 0), 0);
  const lastTotal = representativeRows.reduce((sum, row) => sum + Number(row.lastTotal || 0), 0);
  const metrics = buildWeeklySalesComparisonMetrics(previousTotal, lastTotal);
  const topRepresentative = representativeRows[0] || null;
  const updatedAt = String(options.updatedAt || '').trim();
  const fileName = String(options.fileName || '').trim();

  return {
    ...metrics,
    previousWeek,
    lastWeek,
    hasMixedWeeks: false,
    representativeRows,
    representativesCount: representativeRows.length,
    repsGrowingCount: representativeRows.filter(row => row.direction === 'up').length,
    repsDecliningCount: representativeRows.filter(row => row.direction === 'down').length,
    repsFlatCount: representativeRows.filter(row => row.direction === 'flat').length,
    topRepresentative,
    updatedAt,
    fileName
  };
}

function buildAdminWeeklySalesInsightSummary(representativeInsights){
  const entries = Array.isArray(representativeInsights) ? representativeInsights.filter(Boolean) : [];
  if(!entries.length){
    throw new Error('Nie znaleziono raportów tygodniowych przypisanych do przedstawicieli handlowych.');
  }

  const representativeRows = entries
    .map(entry => {
      const representative = String(entry?.config?.displayName || entry?.config?.storageFolder || '').trim() || 'Bez przedstawiciela';
      const previousWeek = String(entry?.insight?.previousWeek || '').trim();
      const lastWeek = String(entry?.insight?.lastWeek || '').trim();
      const metrics = buildWeeklySalesComparisonMetrics(
        entry?.insight?.previousTotal,
        entry?.insight?.lastTotal
      );

      return {
        representative,
        previousWeek,
        lastWeek,
        updatedAt: String(entry?.insight?.updatedAt || '').trim(),
        fileName: String(entry?.insight?.fileName || '').trim(),
        ...metrics
      };
    })
    .sort((a, b) => b.lastTotal - a.lastTotal || String(a.representative || '').localeCompare(String(b.representative || ''), 'pl'));

  const previousTotal = representativeRows.reduce((sum, row) => sum + Number(row.previousTotal || 0), 0);
  const lastTotal = representativeRows.reduce((sum, row) => sum + Number(row.lastTotal || 0), 0);
  const metrics = buildWeeklySalesComparisonMetrics(previousTotal, lastTotal);
  const topRepresentative = representativeRows[0] || null;
  const periodMap = new Map();

  representativeRows.forEach(row => {
    const key = `${row.previousWeek}|||${row.lastWeek}`;
    const current = periodMap.get(key) || {
      previousWeek: row.previousWeek,
      lastWeek: row.lastWeek,
      count: 0
    };
    current.count += 1;
    periodMap.set(key, current);
  });

  const dominantPeriod = Array.from(periodMap.values())
    .sort((a, b) => b.count - a.count || String(a.lastWeek || '').localeCompare(String(b.lastWeek || ''), 'pl', { numeric: true }))[0] || {
      previousWeek: '',
      lastWeek: '',
      count: 0
    };
  const hasMixedWeeks = periodMap.size > 1;
  const updatedAt = representativeRows
    .map(row => {
      const date = new Date(row.updatedAt);
      return Number.isNaN(date.getTime()) ? null : date;
    })
    .filter(Boolean)
    .sort((a, b) => b.getTime() - a.getTime())[0];

  return {
    ...metrics,
    previousWeek: dominantPeriod.previousWeek,
    lastWeek: dominantPeriod.lastWeek,
    hasMixedWeeks,
    representativeRows,
    representativesCount: representativeRows.length,
    repsGrowingCount: representativeRows.filter(row => row.direction === 'up').length,
    repsDecliningCount: representativeRows.filter(row => row.direction === 'down').length,
    repsFlatCount: representativeRows.filter(row => row.direction === 'flat').length,
    topRepresentative,
    updatedAt: updatedAt ? updatedAt.toISOString() : '',
    fileName: representativeRows.length === 1
      ? (representativeRows[0]?.fileName || '')
      : `${representativeRows.length} raportów PH`
  };
}

function renderAdminSalesInsight(summary){
  const previousWeekShortLabel = String(summary.previousWeek || '').trim();
  const lastWeekShortLabel = String(summary.lastWeek || '').trim();
  const previousWeekLabel = summary.hasMixedWeeks ? 'Różne zakresy tygodni' : formatWeeklySalesWeekLabel(summary.previousWeek);
  const lastWeekLabel = summary.hasMixedWeeks ? 'Różne zakresy tygodni' : formatWeeklySalesWeekLabel(summary.lastWeek);
  const {
    directionClass,
    statusClass,
    statusText,
    statusIcon,
    trendKpiLabel
  } = getSalesInsightStatusMeta(summary.direction, { teamScope: true });
  const maxValue = Math.max(summary.previousTotal, summary.lastTotal, 1);
  const previousWidth = Math.max((summary.previousTotal / maxValue) * 100, summary.previousTotal > 0 ? 8 : 0);
  const lastWidth = Math.max((summary.lastTotal / maxValue) * 100, summary.lastTotal > 0 ? 8 : 0);
  const representativeRowsHtml = summary.representativeRows.map((row, index) => {
    const rowStatusClass = row.direction === 'up'
      ? 'is-up'
      : row.direction === 'down'
        ? 'is-down'
        : 'is-flat';
    const rowStatusIcon = row.direction === 'up' ? '↗' : row.direction === 'down' ? '↘' : '→';
    const canOpenWeeklyReport = typeof openAdminWeeklySalesRepresentative === 'function';

    return `
      <article
        class="sales-insight-team-card ${rowStatusClass} ${canOpenWeeklyReport ? 'is-clickable' : ''}"
        ${canOpenWeeklyReport ? `onclick="openAdminWeeklySalesRepresentative('${escapeAttr(row.representative)}')"` : ''}
      >
        <div class="sales-insight-team-row">
          <span class="sales-insight-team-rank">#${index + 1}</span>
          <div class="sales-insight-team-copy">
            <strong class="sales-insight-team-name">${escapeHtml(row.representative)}</strong>
            <span class="sales-insight-team-meta">${escapeHtml(
              row.lastWeek && row.previousWeek
                ? `${row.previousWeek} → ${row.lastWeek}`
                : 'Porównanie dwóch ostatnich tygodni'
            )}</span>
            ${canOpenWeeklyReport ? '<span class="sales-insight-team-linkhint">Kliknij, żeby zobaczyć więcej szczegółów</span>' : ''}
          </div>
          <div class="sales-insight-team-side">
            <span class="sales-insight-team-side-value">${escapeHtml(formatSalesNumber(row.lastTotal))}</span>
            <span class="sales-insight-team-delta ${rowStatusClass}">${rowStatusIcon} ${escapeHtml(formatSalesDelta(row.difference))}</span>
            <span class="sales-insight-team-percent">${escapeHtml(formatPercentageDelta(row.percentChange))}</span>
          </div>
        </div>
      </article>
    `;
  }).join('');

  setPersonalSalesInsightMarkup(`
    <div class="sales-insight-layout sales-insight-layout--single">
      <div class="sales-insight-shell sales-insight-shell--admin ${directionClass}">
        <div class="sales-insight-main">
          <span class="sales-insight-eyebrow">Panel administratora</span>
          <h2>Sprzedaż wszystkich przedstawicieli</h2>
          <p>Porównanie dwóch ostatnich tygodni na podstawie zbiorczego raportu tygodniowego administratora.</p>
          <div class="sales-insight-status ${statusClass}">
            <span>${statusIcon}</span>
            <strong>${escapeHtml(statusText)}</strong>
          </div>
          <div class="sales-insight-kpis sales-insight-kpis--admin">
            <div class="sales-insight-kpi">
              <span class="sales-insight-kpi-label">Aktualny tydzień</span>
              <strong class="sales-insight-kpi-value">${escapeHtml(formatSalesNumber(summary.lastTotal))}</strong>
            </div>
            <div class="sales-insight-kpi">
              <span class="sales-insight-kpi-label">Poprzedni tydzień</span>
              <strong class="sales-insight-kpi-value">${escapeHtml(formatSalesNumber(summary.previousTotal))}</strong>
            </div>
            <div class="sales-insight-kpi">
              <span class="sales-insight-kpi-label">Przedstawiciele</span>
              <strong class="sales-insight-kpi-value">${formatNumber(summary.representativesCount)}</strong>
            </div>
            <div class="sales-insight-kpi">
              <span class="sales-insight-kpi-label">Trend zespołu</span>
              <strong class="sales-insight-kpi-value">${escapeHtml(trendKpiLabel)}</strong>
            </div>
          </div>
          <div class="sales-insight-actions">
            <button
              type="button"
              class="sales-insight-details-btn"
              onclick="openPersonalSalesReportDetails()"
              data-personal-sales-details-btn="1"
            >Pokaż szczegóły</button>
          </div>
          <div class="sales-insight-change">
            <div class="sales-insight-delta">${escapeHtml(formatSalesDelta(summary.difference))}</div>
            <div class="sales-insight-percent">${escapeHtml(formatPercentageDelta(summary.percentChange))}</div>
          </div>
          <div class="sales-insight-team">
            <div class="sales-insight-team-head">
              <span class="sales-insight-team-title">Przedstawiciele handlowi</span>
              <span class="sales-insight-team-subtitle">${escapeHtml(
                summary.hasMixedWeeks
                  ? 'Każdy PH: ostatni tydzień vs poprzedni'
                  : `${summary.previousWeek || 'brak'} → ${summary.lastWeek || 'brak'}`
              )}</span>
            </div>
            <div class="sales-insight-team-list">
              ${representativeRowsHtml}
            </div>
          </div>
        </div>
        <div class="sales-insight-side">
          <div class="sales-insight-bars">
            <div class="sales-insight-bar-card">
              <div class="sales-insight-bar-meta">
                <span class="sales-insight-bar-chip">${summary.hasMixedWeeks ? 'Poprzednie tygodnie' : 'Poprzedni tydzień'}</span>
                <span class="sales-insight-bar-label">${escapeHtml(summary.hasMixedWeeks ? 'Zespół' : previousWeekShortLabel)}</span>
                <span class="sales-insight-bar-range">${escapeHtml(
                  summary.hasMixedWeeks
                    ? 'Suma poprzednich tygodni ze wszystkich raportów PH.'
                    : previousWeekLabel.replace(previousWeekShortLabel, '').trim().replace(/^\(|\)$/g, '')
                )}</span>
              </div>
              <div class="sales-insight-bar-visual">
                <div class="sales-insight-bar-track">
                  <span class="sales-insight-bar-fill sales-insight-bar-fill--prev" style="width:${previousWidth.toFixed(2)}%"></span>
                </div>
                <span class="sales-insight-bar-value">${escapeHtml(formatSalesNumber(summary.previousTotal))}</span>
              </div>
            </div>
            <div class="sales-insight-bar-card">
              <div class="sales-insight-bar-meta">
                <span class="sales-insight-bar-chip sales-insight-bar-chip--accent">${summary.hasMixedWeeks ? 'Aktualne tygodnie' : 'Aktualny tydzień'}</span>
                <span class="sales-insight-bar-label">${escapeHtml(summary.hasMixedWeeks ? 'Zespół' : lastWeekShortLabel)}</span>
                <span class="sales-insight-bar-range">${escapeHtml(
                  summary.hasMixedWeeks
                    ? 'Suma ostatnich tygodni ze wszystkich raportów PH.'
                    : lastWeekLabel.replace(lastWeekShortLabel, '').trim().replace(/^\(|\)$/g, '')
                )}</span>
              </div>
              <div class="sales-insight-bar-visual">
                <div class="sales-insight-bar-track">
                  <span class="sales-insight-bar-fill sales-insight-bar-fill--last" style="width:${lastWidth.toFixed(2)}%"></span>
                </div>
                <span class="sales-insight-bar-value">${escapeHtml(formatSalesNumber(summary.lastTotal))}</span>
              </div>
            </div>
          </div>
          <div class="sales-insight-admin-side-grid">
            <div class="sales-insight-admin-side-card">
              <span class="sales-insight-admin-side-label">PH na plusie</span>
              <strong class="sales-insight-admin-side-value">${formatNumber(summary.repsGrowingCount)}</strong>
            </div>
            <div class="sales-insight-admin-side-card">
              <span class="sales-insight-admin-side-label">PH ze spadkiem</span>
              <strong class="sales-insight-admin-side-value">${formatNumber(summary.repsDecliningCount)}</strong>
            </div>
            <div class="sales-insight-admin-side-card">
              <span class="sales-insight-admin-side-label">Lider tygodnia</span>
              <strong class="sales-insight-admin-side-value">${escapeHtml(summary.topRepresentative?.representative || 'Brak danych')}</strong>
            </div>
          </div>
          <div class="sales-insight-footnote">
            Ostatnia aktualizacja: ${escapeHtml(formatInsightDate(summary.updatedAt))}
            ${summary.fileName ? `<br>Źródło: ${escapeHtml(summary.fileName)}` : ''}
          </div>
        </div>
      </div>
    </div>
  `);
}

function formatWeeksCountLabel(value){
  const count = Math.max(Number(value) || 0, 0);
  if(count === 1) return '1 tydzień';
  const mod10 = count % 10;
  const mod100 = count % 100;
  if(mod10 >= 2 && mod10 <= 4 && (mod100 < 12 || mod100 > 14)){
    return `${count} tygodnie`;
  }
  return `${count} tygodni`;
}



function getDetailedSalesCustomerKey(code, name){
  return `${String(code || '').trim()}|||${String(name || '').trim()}`;
}

function wrapDetailedSalesTopCompactContent(content){
  return `<div class="reports-compact-mode">${content}</div>`;
}

function getDetailedSalesTopGroupKey(value){
  return String(value || '').trim().toLowerCase();
}

function getDetailedSalesTopCoveragePercent(purchasedCount, offerCount){
  const purchased = Number(purchasedCount || 0);
  const offer = Number(offerCount || 0);
  if(!Number.isFinite(purchased) || !Number.isFinite(offer) || offer <= 0){
    return 0;
  }
  return (purchased / offer) * 100;
}

function getDetailedSalesTopCoverageTone(percent){
  if(percent < 30) return 'low';
  if(percent < 60) return 'medium';
  return 'high';
}

function getDetailedSalesTopCoverageLabel(percent){
  const tone = getDetailedSalesTopCoverageTone(percent);
  if(tone === 'low') return 'Niski udział';
  if(tone === 'medium') return 'Średni udział';
  return 'Wysoki udział';
}

function matchesDetailedSalesTopGroup(topGroup, selectedGroups){
  const normalizedTopGroup = normalizeHeaderKey(topGroup);
  const normalizedSelectedGroups = expandDetailedSalesSelectedGroups(selectedGroups)
    .map(value => normalizeHeaderKey(value))
    .filter(Boolean);
  if(!normalizedSelectedGroups.length) return true;
  return normalizedSelectedGroups.some(normalizedSelectedGroup => (
    normalizedTopGroup === normalizedSelectedGroup
    || normalizedTopGroup.includes(normalizedSelectedGroup)
    || normalizedSelectedGroup.includes(normalizedTopGroup)
  ));
}

function getDetailedSalesDashboardGroupConfig(value){
  const normalizedValue = normalizeHeaderKey(value);
  if(!normalizedValue) return null;
  return REPORT_GROUP_CONFIGS.find(group => (
    (group.groups || []).some(groupName => normalizeHeaderKey(groupName) === normalizedValue)
  )) || null;
}

function getDetailedSalesDashboardGroupLabel(value){
  const fallbackLabel = String(value || '').trim() || 'Bez grupy';
  return getDetailedSalesDashboardGroupConfig(value)?.dashboardLabel || fallbackLabel;
}

function getDetailedSalesShortGroupLabel(value){
  const groupConfig = getDetailedSalesDashboardGroupConfig(value);
  if(groupConfig?.id && REPORT_GROUP_SHORT_LABELS[groupConfig.id]){
    return REPORT_GROUP_SHORT_LABELS[groupConfig.id];
  }
  return String(value || '').trim() || 'Bez grupy';
}

function dedupeReportRows(rows){
  const bestByIndex = new Map();
  (Array.isArray(rows) ? rows : []).forEach(row => {
    const key = normalizeIndexValue(row?.index);
    if(!key) return;
    const current = bestByIndex.get(key);
    const ranking = Number.isFinite(Number(row?.ranking)) ? Number(row.ranking) : Number.POSITIVE_INFINITY;
    if(!current || ranking < current.ranking){
      bestByIndex.set(key, {
        ...row,
        ranking
      });
    }
  });
  return Array.from(bestByIndex.values());
}

function normalizeReportRow(row, headers){
  const indexKey = findHeaderKey(headers, ['indeks', 'index', 'id', 'numer katalogowy', 'sku number']);
  const nameKey = findHeaderKey(headers, ['nazwa', 'name']);
  const producerKey = findHeaderKey(headers, ['skrot producenta', 'skrót producenta', 'producent', 'producer']);
  const eanKey = findHeaderKey(headers, ['kod ean', 'ean']);
  const groupNameKey = findHeaderKey(headers, ['nazwa grupa', 'nazwa_grupa', 'grupa', 'group']);
  const rankingKey = findHeaderKey(headers, ['sprzedaż ilościowa', 'sprzedaz ilosciowa', 'ranking']);
  const productGroupKey = findHeaderKey(headers, ['grupa produktowa', 'grupa produktowa ']);

  const rankingRaw = row[rankingKey];
  const rankingNumber = Number(String(rankingRaw ?? '').replace(',', '.'));

  return {
    index: String(row[indexKey] ?? '').trim(),
    name: String(row[nameKey] ?? '').trim(),
    producer: String(row[producerKey] ?? '').trim(),
    ean: String(row[eanKey] ?? '').trim(),
    groupName: String(row[groupNameKey] ?? '').trim(),
    ranking: Number.isFinite(rankingNumber) ? rankingNumber : Number.POSITIVE_INFINITY,
    rankingLabel: String(rankingRaw ?? '').trim(),
    productGroup: String(row[productGroupKey] ?? '').trim()
  };
}

function generateReportsData(){
  const generated = [];
  reportsSummary = [];

  REPORT_GROUP_CONFIGS.forEach(group => {
    const limitRaw = Number(reportsGroupLimits[group.id]);
    const limit = Number.isFinite(limitRaw) && limitRaw >= 0 ? Math.floor(limitRaw) : 25;
    const sourceRows = reportsSourceRows.filter(row => group.groups.includes(row.productGroup));
    const uniqueRows = dedupeReportRows(sourceRows).sort((a, b) => {
      if(a.ranking !== b.ranking) return a.ranking - b.ranking;
      return String(a.name).localeCompare(String(b.name), 'pl');
    });
    const pickedRows = uniqueRows.slice(0, limit);

    reportsSummary.push({
      id: group.id,
      label: group.label,
      available: uniqueRows.length,
      selected: pickedRows.length,
      requested: limit
    });

    pickedRows.forEach(row => {
      generated.push({
        INDEKS: row.index,
        NAZWA: row.name,
        SKROT_PRODUCENTA: row.producer,
        'KOD EAN': row.ean || '',
        NAZWA_GRUPA: row.groupName,
        Ranking: row.rankingLabel || (Number.isFinite(row.ranking) ? String(row.ranking) : ''),
        'Grupa produktowa': group.label
      });
    });
  });

  reportsGeneratedRows = generated;
}

function pickReportFile(inputId){
  const input = document.getElementById(inputId);
  if(!input) return;
  input.value = '';
  input.click();
}

function renderReportLimits(){
  return REPORT_GROUP_CONFIGS.map(group => `
    <label class="reports-limit-card">
      <span class="reports-limit-title">${escapeHtml(group.label)}</span>
      <input
        type="number"
        min="0"
        value="${escapeAttr(reportsGroupLimits[group.id] ?? 25)}"
        onchange="setReportGroupLimit('${group.id}', this.value)"
      >
    </label>
  `).join('');
}

function renderTopSalesReportContent(){
  const summary = reportsSummary.length
    ? `<div class="reports-summary">
        ${reportsSummary.map(item => `
          <div class="reports-summary-card">
            <div class="reports-summary-title">${escapeHtml(item.label)}</div>
            <div class="reports-summary-meta">Wybrano ${item.selected} z ${item.available} pozycji</div>
          </div>
        `).join('')}
      </div>`
    : '';

  const rowsHtml = reportsGeneratedRows.map((row, index) => {
    const idx = String(row.INDEKS ?? '').trim();
    return `
      <tr>
        <td>${index + 1}</td>
        <td class="reports-image-cell">
          ${idx ? buildProductImageTag(idx, imageBaseUrlRumunia, 'reports-image') : ''}
        </td>
        <td>${escapeHtml(row.INDEKS ?? '')}</td>
        <td>${escapeHtml(row.NAZWA ?? '')}</td>
        <td>${escapeHtml(row.SKROT_PRODUCENTA ?? '')}</td>
        <td>${escapeHtml(row.NAZWA_GRUPA ?? '')}</td>
        <td>${escapeHtml(row.Ranking ?? '')}</td>
        <td>${escapeHtml(row['Grupa produktowa'] ?? '')}</td>
      </tr>
    `;
  }).join('');

  return `
    <div class="reports-toolbar">
      <div>
        <div class="import-title">Top sprzedaż</div>
        <div class="import-sub">Importuj Excel i generuj top indeksy według grup produktowych.</div>
      </div>
      <div class="reports-actions">
        <input id="reports-excel-input" class="import-input reports-file-input" type="file" accept=".xlsx,.xls" onchange="importReportsExcel()">
        <button class="btn-outline" onclick="pickReportFile('reports-excel-input')">Importuj Excel</button>
        <button class="btn-outline" onclick="exportReportsXlsx()" ${reportsGeneratedRows.length ? '' : 'disabled'}>Zapisz XLSX</button>
      </div>
      <div class="import-info">
        ${reportsImportFile ? 'Status: plik zaimportowany' : 'Status: oczekuje na import'}
      </div>
    </div>

    <div class="reports-limits">
      ${renderReportLimits()}
    </div>

    ${summary}

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Lp.</th>
            <th>Zdjęcie</th>
            <th>INDEKS</th>
            <th>NAZWA</th>
            <th>SKROT_PRODUCENTA</th>
            <th>NAZWA_GRUPA</th>
            <th>Ranking</th>
            <th>Grupa produktowa</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml || `<tr><td colspan="8" class="reports-empty">Zaimportuj plik Excel, aby wygenerować raport.</td></tr>`}
        </tbody>
      </table>
    </div>
  `;
}

function normalizeClientReportRow(row, headers){
  const indexKey = findHeaderKey(headers, ['indeks', 'index', 'id']);
  const nameKey = findHeaderKey(headers, ['nazwa', 'name']);
  const groupNameKey = findHeaderKey(headers, ['nazwa grupa', 'nazwa_grupa']);
  const repKey = findHeaderKey(headers, ['opiekun klienta', 'opiekun_klienta', 'przedstawiciel handlowy']);
  const customerCodeKey = findHeaderKey(headers, ['kod kh', 'kod_kh', 'numer klienta']);
  const customerNameKey = findHeaderKey(headers, ['nazwa kh', 'nazwa_kh', 'nazwa klienta']);

  return {
    index: String(row[indexKey] ?? '').trim(),
    name: String(row[nameKey] ?? '').trim(),
    groupName: String(row[groupNameKey] ?? '').trim(),
    representative: String(row[repKey] ?? '').trim(),
    customerCode: String(row[customerCodeKey] ?? '').trim(),
    customerName: String(row[customerNameKey] ?? '').trim()
  };
}

function normalizeWeeklySalesRow(row, headers){
  const indexKey = findHeaderKey(headers, ['indeks', 'index', 'id', 'numer katalogowy', 'sku number']);
  const nameKey = findHeaderKey(headers, ['nazwa', 'name', 'nazwa towaru', 'nazwa produktu']);
  const producerKey = findHeaderKey(headers, ['skrot producenta', 'skrót producenta', 'producent', 'producer']);
  const repKey = findHeaderKey(headers, ['opiekun klienta', 'opiekun_klienta', 'przedstawiciel handlowy', 'przedstawiciel', 'handlowiec']);
  const customerCodeKey = findHeaderKey(headers, ['kod kh', 'kod_kh', 'numer klienta', 'kod klienta']);
  const customerNameKey = findHeaderKey(headers, ['nazwa kh', 'nazwa_kh', 'nazwa klienta', 'klient']);
  const weekKey = findHeaderKey(headers, ['tydzień', 'tydzien', 'week', 'nr tygodnia', 'numer tygodnia']);
  const salesKey = findHeaderKey(headers, ['sprzedaż ilościowa', 'sprzedaz ilosciowa', 'ilość', 'ilosc', 'szt', 'quantity', 'qty']);
  const valueKey = findHeaderKey(headers, ['sprzedaż wartościowa', 'sprzedaz wartosciowa', 'wartość', 'wartosc', 'value']);

  const salesRaw = row[salesKey];
  const valueRaw = row[valueKey];
  const salesNumber = Number(String(salesRaw ?? '').replace(/\s+/g, '').replace(',', '.'));
  const valueNumber = Number(String(valueRaw ?? '').replace(/\s+/g, '').replace(',', '.'));

  return {
    representative: String(row[repKey] ?? '').trim() || 'Bez przedstawiciela',
    customerCode: String(row[customerCodeKey] ?? '').trim(),
    customerName: String(row[customerNameKey] ?? '').trim() || 'Bez nazwy klienta',
    index: String(row[indexKey] ?? '').trim(),
    name: String(row[nameKey] ?? '').trim(),
    producer: String(row[producerKey] ?? '').trim(),
    week: String(row[weekKey] ?? '').trim() || 'Bez tygodnia',
    quantity: Number.isFinite(salesNumber) ? salesNumber : 0,
    value: Number.isFinite(valueNumber) ? valueNumber : 0
  };
}

function buildClientReportImportFallback(fileName){
  const baseName = String(fileName || '')
    .replace(/\.[^.]+$/, '')
    .trim();
  const customerName = baseName || 'Import Excel';
  return {
    representative: 'Import Excel',
    customerCode: 'IMPORT',
    customerName
  };
}

async function loadOfferRowsForGroup(group){
  if(reportOfferRowsCache.has(group.id)) return reportOfferRowsCache.get(group.id);

  const allRows = [];
  for(const url of group.sources || []){
    const res = await fetch(url);
    const data = await res.json();
    (data || []).forEach(row => {
      const rankingRaw = getValueByVariants(row, ['ranking']);
      const rankingValue = Number(String(rankingRaw ?? '').replace(',', '.'));
      const normalized = {
        index: String(getValueByVariants(row, ['indeks', 'index', 'id', 'numer katalogowy', 'sku number']) || '').trim(),
        name: String(getValueByVariants(row, ['nazwa', 'name']) || '').trim(),
        producer: String(getValueByVariants(row, ['producent', 'producer']) || '').trim(),
        ean: String(getValueByVariants(row, ['kod ean', 'ean']) || '').trim(),
        ranking: Number.isFinite(rankingValue) ? rankingValue : Number.POSITIVE_INFINITY,
        rankingLabel: String(rankingRaw ?? '').trim(),
        productGroup: group.label
      };
      if(normalized.index) allRows.push(normalized);
    });
  }

  reportOfferRowsCache.set(group.id, allRows);
  return allRows;
}

function getClientReportRepresentatives(){
  return Array.from(new Set(clientReportSourceRows.map(row => row.representative).filter(Boolean)))
    .sort((a, b) => a.localeCompare(b, 'pl'));
}

function getClientReportCustomers(){
  if(!clientReportSelectedRepresentative) return [];
  const map = new Map();
  clientReportSourceRows
    .filter(row => row.representative === clientReportSelectedRepresentative)
    .forEach(row => {
      const key = `${row.customerCode}|||${row.customerName}`;
      if(!map.has(key)){
        map.set(key, { code: row.customerCode, name: row.customerName });
      }
    });
  return Array.from(map.values()).sort((a, b) => {
    const nameCompare = a.name.localeCompare(b.name, 'pl');
    return nameCompare || a.code.localeCompare(b.code, 'pl');
  });
}

function buildReportsAssistantCustomers(rows = clientReportSourceRows){
  const customerMap = new Map();

  (Array.isArray(rows) ? rows : []).forEach(row => {
    const representative = String(row.representative || '').trim();
    const customerCode = normalizeCustomerCodeValue(row.customerCode);
    const customerName = String(row.customerName || '').trim();
    const key = `${representative}|||${customerCode}|||${customerName}`;

    if(!String(key).replace(/\|/g, '').trim()) return;

    const current = customerMap.get(key) || {
      key,
      representative,
      customerCode,
      customerName,
      importedRowsCount: 0,
      uniqueIndexes: new Set()
    };

    current.importedRowsCount += 1;

    const normalizedIndex = normalizeIndexValue(row.index);
    if(normalizedIndex){
      current.uniqueIndexes.add(normalizedIndex);
    }

    customerMap.set(key, current);
  });

  return Array.from(customerMap.values())
    .map(item => ({
      key: item.key,
      representative: item.representative,
      customerCode: item.customerCode,
      customerName: item.customerName,
      importedRowsCount: item.importedRowsCount,
      uniqueIndexesCount: item.uniqueIndexes.size
    }))
    .sort((a, b) => {
      const codeCompare = String(a.customerCode || '').localeCompare(String(b.customerCode || ''), 'pl');
      if(codeCompare) return codeCompare;
      const nameCompare = String(a.customerName || '').localeCompare(String(b.customerName || ''), 'pl');
      if(nameCompare) return nameCompare;
      return String(a.representative || '').localeCompare(String(b.representative || ''), 'pl');
    });
}

function buildReportsAssistantWeeklyRows(){
  if(typeof weeklySalesSourceRows === 'undefined' || !Array.isArray(weeklySalesSourceRows)){
    return [];
  }
  return buildReportsAssistantSourceRows(weeklySalesSourceRows);
}

function mapReportsAssistantOfferRow(row){
  return {
    index: String(row.index || '').trim(),
    name: String(row.name || '').trim(),
    producer: String(row.producer || '').trim(),
    ean: String(row.ean || '').trim(),
    ranking: Number.isFinite(row.ranking) ? row.ranking : Number.POSITIVE_INFINITY,
    rankingLabel: String(row.rankingLabel || '').trim()
  };
}

function buildReportsAssistantGroups(sourceGroups = REPORT_GROUP_CONFIGS){
  return sourceGroups.map(group => ({
    id: group.id,
    label: group.label,
    dashboardLabel: group.dashboardLabel,
    shortLabel: REPORT_GROUP_SHORT_LABELS[group.id] || group.dashboardLabel || group.label
  }));
}

function buildReportsAssistantSourceRows(rows = []){
  return rows.map(row => ({
    representative: String(row.representative || '').trim(),
    customerCode: normalizeCustomerCodeValue(row.customerCode),
    customerName: String(row.customerName || '').trim(),
    index: String(row.index || '').trim(),
    name: String(row.name || '').trim(),
    producer: String(row.producer || '').trim(),
    groupName: String(row.groupName || '').trim(),
    week: String(row.week || '').trim(),
    quantity: Number.isFinite(Number(row.quantity)) ? Number(row.quantity) : 0,
    value: Number.isFinite(Number(row.value)) ? Number(row.value) : 0,
    reportLabel: String(row.reportLabel || '').trim(),
    reportOrder: Number.isFinite(Number(row.reportOrder)) ? Number(row.reportOrder) : 0
  }));
}

async function buildReportsAssistantClientGapSnapshot(){
  const [selectedCustomerCode = '', selectedCustomerName = ''] = String(clientReportSelectedCustomer || '').split('|||');
  const offerRowsByGroupEntries = await Promise.all(
    REPORT_GROUP_CONFIGS.map(async group => {
      const offerRows = dedupeReportRows(await loadOfferRowsForGroup(group))
        .sort((a, b) => {
          if(a.ranking !== b.ranking) return a.ranking - b.ranking;
          return String(a.name || '').localeCompare(String(b.name || ''), 'pl');
        })
        .map(mapReportsAssistantOfferRow);

      return [group.id, offerRows];
    })
  );

  return {
    mode: 'client-gap',
    importFile: String(clientReportImportFile || '').trim(),
    selectedRepresentative: String(clientReportSelectedRepresentative || '').trim(),
    selectedCustomerCode: normalizeCustomerCodeValue(selectedCustomerCode),
    selectedCustomerName: String(selectedCustomerName || '').trim(),
    groupLimits: { ...reportsGroupLimits },
    groups: buildReportsAssistantGroups(),
    customers: buildReportsAssistantCustomers(),
    sourceRows: buildReportsAssistantSourceRows(clientReportSourceRows),
    weeklySalesRows: buildReportsAssistantWeeklyRows(),
    offerRowsByGroup: Object.fromEntries(offerRowsByGroupEntries)
  };
}

async function buildReportsAssistantDetailedSalesSnapshot(){
  const selectedGroups = Array.isArray(detailedSalesSelectedGroups) ? [...detailedSalesSelectedGroups] : [];
  const activeGroupConfigs = REPORT_GROUP_CONFIGS.filter(group => (
    !selectedGroups.length
    || (group.groups || []).some(groupName => matchesDetailedSalesTopGroup(groupName, selectedGroups))
  ));

  const scopedTopRows = dedupeReportRows(
    (await loadTopRumuniaOfferRows()).filter(row => matchesDetailedSalesTopGroup(row.group, selectedGroups))
  ).sort((a, b) => {
    if(a.ranking !== b.ranking) return a.ranking - b.ranking;
    return String(a.name || '').localeCompare(String(b.name || ''), 'pl');
  });

  const offerRowsByGroup = Object.fromEntries(
    activeGroupConfigs.map(group => {
      const rows = scopedTopRows
        .filter(row => getDetailedSalesDashboardGroupConfig(row.group)?.id === group.id)
        .map(mapReportsAssistantOfferRow);
      return [group.id, rows];
    })
  );

  const customers = typeof getDetailedSalesCustomers === 'function'
    ? getDetailedSalesCustomers().map(customer => ({
      key: `${customer.code}|||${customer.name}`,
      representative: '',
      customerCode: normalizeCustomerCodeValue(customer.code),
      customerName: String(customer.name || '').trim(),
      importedRowsCount: 0,
      uniqueIndexesCount: 0
    }))
    : [];

  const [selectedCustomerCode = '', selectedCustomerName = ''] = String(detailedSalesComparisonCustomer || '').split('|||');

  return {
    mode: 'top-rumunia-comparison',
    importFile: String(detailedSalesImportFile || '').trim(),
    selectedRepresentative: '',
    selectedCustomerCode: normalizeCustomerCodeValue(selectedCustomerCode),
    selectedCustomerName: String(selectedCustomerName || '').trim(),
    groupLimits: {},
    groups: buildReportsAssistantGroups(activeGroupConfigs),
    customers,
    sourceRows: buildReportsAssistantSourceRows(detailedSalesSourceRows),
    weeklySalesRows: buildReportsAssistantWeeklyRows(),
    offerRowsByGroup
  };
}

async function buildReportsAssistantWeeklySalesSnapshot(){
  const [selectedCustomerCode = '', selectedCustomerName = ''] = String(weeklySalesSelectedCustomer || '').split('|||');
  const offerRowsByGroupEntries = await Promise.all(
    REPORT_GROUP_CONFIGS.map(async group => {
      const offerRows = dedupeReportRows(await loadOfferRowsForGroup(group))
        .sort((a, b) => {
          if(a.ranking !== b.ranking) return a.ranking - b.ranking;
          return String(a.name || '').localeCompare(String(b.name || ''), 'pl');
        })
        .map(mapReportsAssistantOfferRow);

      return [group.id, offerRows];
    })
  );

  const weeklyRows = buildReportsAssistantWeeklyRows();

  return {
    mode: 'weekly-sales',
    importFile: String(weeklySalesImportFile || '').trim(),
    selectedRepresentative: String(weeklySalesSelectedRepresentative || '').trim(),
    selectedCustomerCode: normalizeCustomerCodeValue(selectedCustomerCode),
    selectedCustomerName: String(selectedCustomerName || '').trim(),
    groupLimits: {},
    groups: buildReportsAssistantGroups(),
    customers: buildReportsAssistantCustomers(weeklySalesSourceRows),
    sourceRows: weeklyRows,
    weeklySalesRows: weeklyRows,
    offerRowsByGroup: Object.fromEntries(offerRowsByGroupEntries)
  };
}

async function getReportsAssistantSnapshot(){
  if(detailedSalesSourceRows.length){
    return buildReportsAssistantDetailedSalesSnapshot();
  }

  if(clientReportSourceRows.length){
    return buildReportsAssistantClientGapSnapshot();
  }

  if(typeof weeklySalesSourceRows !== 'undefined' && Array.isArray(weeklySalesSourceRows) && weeklySalesSourceRows.length){
    return buildReportsAssistantWeeklySalesSnapshot();
  }

  return {
    mode: reportsMode,
    importFile: '',
    selectedRepresentative: '',
    selectedCustomerCode: '',
    selectedCustomerName: '',
    groupLimits: {},
    groups: buildReportsAssistantGroups(),
    customers: [],
    sourceRows: [],
    weeklySalesRows: [],
    offerRowsByGroup: {}
  };
}

window.getReportsAssistantSnapshot = getReportsAssistantSnapshot;

function getTopSuggestionRepresentatives(){
  return Array.from(new Set(topSuggestionSourceRows.map(row => row.representative).filter(Boolean)))
    .sort((a, b) => a.localeCompare(b, 'pl'));
}

function getTopSuggestionCustomers(){
  if(!topSuggestionSelectedRepresentative) return [];
  const map = new Map();
  topSuggestionSourceRows
    .filter(row => row.representative === topSuggestionSelectedRepresentative)
    .forEach(row => {
      const key = `${row.customerCode}|||${row.customerName}`;
      if(!map.has(key)){
        map.set(key, { code: row.customerCode, name: row.customerName });
      }
    });
  return Array.from(map.values()).sort((a, b) => {
    const nameCompare = a.name.localeCompare(b.name, 'pl');
    return nameCompare || a.code.localeCompare(b.code, 'pl');
  });
}

async function generateTopSuggestionsReport(){
  topSuggestionGeneratedRows = [];
  topSuggestionPurchasedRows = [];
  topSuggestionSummary = null;
  if(!topSuggestionSelectedRepresentative || !topSuggestionSelectedCustomer) return;

  const [customerCode, customerName] = topSuggestionSelectedCustomer.split('|||');
  const customerRows = topSuggestionSourceRows.filter(row => (
    row.representative === topSuggestionSelectedRepresentative
    && row.customerCode === customerCode
    && row.customerName === customerName
  ));
  const latestWeeks = getLatestWeeksFromRows(customerRows, 3);
  const purchasedProducerMap = new Map();
  customerRows
    .filter(row => latestWeeks.includes(String(row.week || '')) && (row.quantity || row.value) > 0)
    .forEach(row => {
      const producer = String(row.producer || '').trim();
      const producerKey = normalizeProducerValue(producer);
      if(!producerKey) return;
      const current = purchasedProducerMap.get(producerKey) || {
        producerCode: row.index || '',
        producer,
        salesTotal: 0
      };
      current.salesTotal += row.quantity || row.value;
      purchasedProducerMap.set(producerKey, current);
    });
  const purchasedProducerSet = new Set(purchasedProducerMap.keys());

  const topRows = dedupeReportRows(await loadTopRumuniaOfferRows()).sort((a, b) => {
    if(a.ranking !== b.ranking) return a.ranking - b.ranking;
    return String(a.name).localeCompare(String(b.name), 'pl');
  });
  const topProducerMap = new Map();
  topRows.forEach(row => {
    const producer = String(row.producer || '').trim();
    const producerKey = normalizeProducerValue(producer);
    if(!producerKey) return;
    const current = topProducerMap.get(producerKey) || {
      producer,
      productCount: 0,
      bestRanking: row.ranking,
      bestRankingLabel: row.rankingLabel || (Number.isFinite(row.ranking) ? String(row.ranking) : ''),
      exampleIndex: row.index,
      exampleName: row.name
    };
    current.productCount += 1;
    if(row.ranking < current.bestRanking){
      current.bestRanking = row.ranking;
      current.bestRankingLabel = row.rankingLabel || (Number.isFinite(row.ranking) ? String(row.ranking) : '');
      current.exampleIndex = row.index;
      current.exampleName = row.name;
    }
    topProducerMap.set(producerKey, current);
  });

  const missingRows = Array.from(topProducerMap.entries())
    .filter(([producerKey]) => !purchasedProducerSet.has(producerKey))
    .map(([, row]) => row)
    .sort((a, b) => {
      if(a.bestRanking !== b.bestRanking) return a.bestRanking - b.bestRanking;
      return String(a.producer).localeCompare(String(b.producer), 'pl');
    });
  const limit = Math.max(0, Math.floor(Number(topSuggestionLimit)));
  const pickedRows = limit ? missingRows.slice(0, limit) : missingRows;

  topSuggestionPurchasedRows = Array.from(purchasedProducerMap.values())
    .sort((a, b) => b.salesTotal - a.salesTotal || String(a.producer).localeCompare(String(b.producer), 'pl'));
  topSuggestionGeneratedRows = pickedRows.map(row => ({
    producer: row.producer,
    productCount: row.productCount,
    bestRanking: row.bestRankingLabel,
    exampleIndex: row.exampleIndex,
    exampleName: row.exampleName
  }));
  topSuggestionSummary = {
    weeks: latestWeeks,
    purchasedCount: topSuggestionPurchasedRows.length,
    missingCount: missingRows.length,
    shownCount: topSuggestionGeneratedRows.length
  };
}

async function generateClientComparisonReport(){
  if(!clientReportSelectedRepresentative || !clientReportSelectedCustomer){
    clientReportPurchasedRows = [];
    clientReportMissingRows = [];
    clientReportSummary = [];
    clientReportSelectedMissingIndexes = new Set();
    return;
  }

  const [customerCode, customerName] = clientReportSelectedCustomer.split('|||');
  const clientRows = clientReportSourceRows.filter(row => (
    row.representative === clientReportSelectedRepresentative
    && row.customerCode === customerCode
    && row.customerName === customerName
  ));
  const purchasedIndexSet = new Set(clientRows.map(row => normalizeIndexValue(row.index)).filter(Boolean));
  const purchased = [];
  const missing = [];
  const summary = [];

  for(const group of REPORT_GROUP_CONFIGS){
    const limitRaw = Number(reportsGroupLimits[group.id]);
    const limit = Number.isFinite(limitRaw) && limitRaw >= 0 ? Math.floor(limitRaw) : 25;
    const offerSourceRows = await loadOfferRowsForGroup(group);
    const offerRows = dedupeReportRows(
      offerSourceRows.map(row => ({
        index: row.index,
        name: row.name,
        producer: row.producer,
        ean: row.ean,
        ranking: Number.isFinite(row.ranking) ? row.ranking : Number.POSITIVE_INFINITY,
        rankingLabel: row.rankingLabel,
        productGroup: group.label
      }))
    ).sort((a, b) => {
      if(a.ranking !== b.ranking) return a.ranking - b.ranking;
      return String(a.name).localeCompare(String(b.name), 'pl');
    });

    const purchasedFromOffer = offerRows.filter(row => purchasedIndexSet.has(normalizeIndexValue(row.index)));
    const missingTopRows = offerRows.filter(row => !purchasedIndexSet.has(normalizeIndexValue(row.index))).slice(0, limit);

    let purchasedCount = 0;
    let missingCount = 0;

    purchasedFromOffer.forEach(row => {
      purchased.push({
        INDEKS: row.index,
        NAZWA: row.name,
        SKROT_PRODUCENTA: row.producer,
        'KOD EAN': row.ean || '',
        Ranking: row.rankingLabel || (Number.isFinite(row.ranking) ? String(row.ranking) : ''),
        'Grupa produktowa': group.label
      });
      purchasedCount += 1;
    });

    missingTopRows.forEach(row => {
      missing.push({
        INDEKS: row.index,
        NAZWA: row.name,
        SKROT_PRODUCENTA: row.producer,
        'KOD EAN': row.ean || '',
        Ranking: row.rankingLabel || (Number.isFinite(row.ranking) ? String(row.ranking) : ''),
        'Grupa produktowa': group.label
      });
      missingCount += 1;
    });

    summary.push({
      label: group.label,
      purchased: purchasedCount,
      missing: missingCount,
      total: missingTopRows.length,
      offerTotal: offerRows.length
    });
  }

  clientReportPurchasedRows = purchased;
  clientReportMissingRows = missing;
  clientReportSummary = summary;
  clientReportSelectedMissingIndexes = new Set(
    clientReportMissingRows.map(row => normalizeIndexValue(row.INDEKS)).filter(Boolean)
  );
}

function renderClientRowsTable(rows, emptyText, options = {}){
  const { selectable = false } = options;
  const html = rows.map((row, index) => {
    const idx = String(row.INDEKS ?? '').trim();
    const normalizedIndex = normalizeIndexValue(idx);
    const checked = selectable && clientReportSelectedMissingIndexes.has(normalizedIndex) ? 'checked' : '';
    return `
      <tr>
        <td>${index + 1}</td>
        ${selectable ? `<td><input type="checkbox" data-index="${escapeAttr(idx)}" onchange="toggleClientMissingSelection(this)" ${checked}></td>` : ''}
        <td class="reports-image-cell">
          ${idx ? buildProductImageTag(idx, imageBaseUrlRumunia, 'reports-image') : ''}
        </td>
        <td>${escapeHtml(row.INDEKS ?? '')}</td>
        <td>${escapeHtml(row.NAZWA ?? '')}</td>
        <td>${escapeHtml(row.SKROT_PRODUCENTA ?? '')}</td>
        <td>${escapeHtml(row.Ranking ?? '')}</td>
        <td>${escapeHtml(row['Grupa produktowa'] ?? '')}</td>
      </tr>
    `;
  }).join('');

  return `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Lp.</th>
            ${selectable ? '<th>Wybór</th>' : ''}
            <th>Zdjęcie</th>
            <th>INDEKS</th>
            <th>NAZWA</th>
            <th>SKROT_PRODUCENTA</th>
            <th>Ranking</th>
            <th>Grupa produktowa</th>
          </tr>
        </thead>
        <tbody>
          ${html || `<tr><td colspan="${selectable ? 8 : 7}" class="reports-empty">${escapeHtml(emptyText)}</td></tr>`}
        </tbody>
      </table>
    </div>
  `;
}

function renderPurchasedProducerRowsTable(rows, emptyText){
  const html = rows.map((row, index) => `
    <tr>
      <td>${index + 1}</td>
      <td>${escapeHtml(row.producerCode || '')}</td>
      <td>${escapeHtml(row.producer || '')}</td>
      <td>${formatNumber(row.salesTotal || 0)}</td>
    </tr>
  `).join('');

  return `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Lp.</th>
            <th>Nr producenta</th>
            <th>Producent</th>
            <th>Sprzedaż</th>
          </tr>
        </thead>
        <tbody>
          ${html || `<tr><td colspan="4" class="reports-empty">${escapeHtml(emptyText)}</td></tr>`}
        </tbody>
      </table>
    </div>
  `;
}

function renderMissingProducerRowsTable(rows, emptyText){
  const html = rows.map((row, index) => `
    <tr>
      <td>${index + 1}</td>
      <td>${escapeHtml(row.producer || '')}</td>
      <td>${escapeHtml(row.productCount || '')}</td>
      <td>${escapeHtml(row.bestRanking || '')}</td>
      <td>${escapeHtml(row.exampleIndex || '')}</td>
      <td>${escapeHtml(row.exampleName || '')}</td>
    </tr>
  `).join('');

  return `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Lp.</th>
            <th>Producent</th>
            <th>Ile indeksów w Top</th>
            <th>Najwyższy ranking</th>
            <th>Przykładowy indeks</th>
            <th>Przykładowy produkt</th>
          </tr>
        </thead>
        <tbody>
          ${html || `<tr><td colspan="6" class="reports-empty">${escapeHtml(emptyText)}</td></tr>`}
        </tbody>
      </table>
    </div>
  `;
}

function getFilteredClientRows(rows, filters){
  return rows.filter(row => {
    if(filters.producer && !includesText(row.SKROT_PRODUCENTA, filters.producer)) return false;
    if(filters.ranking && !includesText(row.Ranking, filters.ranking)) return false;
    if(filters.group && !includesText(row['Grupa produktowa'], filters.group)) return false;
    return true;
  });
}

function renderClientTableFilters(type, filters){
  const rows = type === 'purchased' ? clientReportPurchasedRows : clientReportMissingRows;
  const producerOptions = Array.from(new Set(rows.map(row => String(row.SKROT_PRODUCENTA || '').trim()).filter(Boolean)))
    .sort((a, b) => a.localeCompare(b, 'pl'))
    .map(producer => `<option value="${escapeAttr(producer)}" ${filters.producer === producer ? 'selected' : ''}>${escapeHtml(producer)}</option>`)
    .join('');
  const groupOptions = Array.from(new Set(rows.map(row => String(row['Grupa produktowa'] || '').trim()).filter(Boolean)))
    .sort((a, b) => a.localeCompare(b, 'pl'))
    .map(group => `<option value="${escapeAttr(group)}" ${filters.group === group ? 'selected' : ''}>${escapeHtml(group)}</option>`)
    .join('');

  return `
    <div class="reports-table-filters">
      <select onchange="setClientTableFilter('${type}', 'producer', this.value)">
        <option value="">Wszyscy producenci</option>
        ${producerOptions}
      </select>
      <input type="text" value="${escapeAttr(filters.ranking)}" placeholder="Filtruj po rankingu" oninput="setClientTableFilter('${type}', 'ranking', this.value)">
      <select onchange="setClientTableFilter('${type}', 'group', this.value)">
        <option value="">Wszystkie grupy produktowe</option>
        ${groupOptions}
      </select>
    </div>
  `;
}

function renderClientComparisonContent(){
  const representatives = getClientReportRepresentatives();
  const customers = getClientReportCustomers();
  const filteredPurchasedRows = getFilteredClientRows(clientReportPurchasedRows, clientReportPurchasedFilters);
  const filteredMissingRows = getFilteredClientRows(clientReportMissingRows, clientReportMissingFilters);
  const summary = clientReportSummary.length
    ? `<div class="reports-summary">
        ${clientReportSummary.map(item => `
          <div class="reports-summary-card">
            <div class="reports-summary-title">${escapeHtml(item.label)}</div>
            <div class="reports-summary-meta">W ofercie: ${item.offerTotal} | Klient kupił: ${item.purchased} | Rekomendowane: ${item.missing}</div>
          </div>
        `).join('')}
      </div>`
    : '';

  const selectedMissingCount = clientReportSelectedMissingIndexes.size;

  return `
    <div class="reports-toolbar">
      <div>
        <div class="import-title">Potencjał klienta Top Rumunia</div>
        <div class="import-sub">Porównaj zakupy klienta do listy Top Rumunia według grup produktowych.</div>
      </div>
      <div class="reports-actions">
        <input id="client-reports-excel-input" class="import-input reports-file-input" type="file" accept=".xlsx,.xls" onchange="importClientReportExcel()">
        <button class="btn-outline" onclick="pickReportFile('client-reports-excel-input')">Importuj Excel</button>
        <button class="btn-outline" onclick="exportClientRecommendationXlsx()" ${selectedMissingCount ? '' : 'disabled'}>Zapisz XLSX</button>
        <button class="btn-outline" onclick="createClientRecommendationCatalog()" ${selectedMissingCount ? '' : 'disabled'}>Utwórz katalog</button>
      </div>
      <div class="import-info">
        ${clientReportImportFile ? 'Status: plik zaimportowany' : 'Status: oczekuje na import'}
        ${clientReportSelectedCustomer ? ` | Zaznaczone rekomendacje: ${selectedMissingCount}` : ''}
      </div>
    </div>

    <div class="reports-filters">
      <label class="reports-filter-card">
        <span class="reports-limit-title">Przedstawiciel handlowy</span>
        <select onchange="setClientReportRepresentative(this.value)">
          <option value="">Wybierz</option>
          ${representatives.map(rep => `<option value="${escapeAttr(rep)}" ${clientReportSelectedRepresentative === rep ? 'selected' : ''}>${escapeHtml(rep)}</option>`).join('')}
        </select>
      </label>
      <label class="reports-filter-card">
        <span class="reports-limit-title">Numer klienta</span>
        <select onchange="setClientReportCustomer(this.value)">
          <option value="">Wybierz</option>
          ${customers.map(customer => {
            const value = `${customer.code}|||${customer.name}`;
            return `<option value="${escapeAttr(value)}" ${clientReportSelectedCustomer === value ? 'selected' : ''}>${escapeHtml(customer.code || customer.name)}</option>`;
          }).join('')}
        </select>
      </label>
      <label class="reports-filter-card">
        <span class="reports-limit-title">Nazwa klienta</span>
        <select onchange="setClientReportCustomer(this.value)">
          <option value="">Wybierz</option>
          ${customers.map(customer => {
            const value = `${customer.code}|||${customer.name}`;
            return `<option value="${escapeAttr(value)}" ${clientReportSelectedCustomer === value ? 'selected' : ''}>${escapeHtml(customer.name || customer.code)}</option>`;
          }).join('')}
        </select>
      </label>
    </div>

    <div class="reports-limits">
      ${renderReportLimits()}
    </div>

    <label class="reports-cover-toggle">
      <input type="checkbox" ${clientRecommendationIncludeCover ? 'checked' : ''} onchange="setClientRecommendationCoverOption(this.checked)">
      <span>Dodaj okladke z komunikatem dla klienta podczas tworzenia katalogu</span>
    </label>

    ${summary}

    <div class="reports-two-columns">
      <div class="reports-column">
        <h3 class="reports-table-title">Kupione przez klienta</h3>
        ${renderClientTableFilters('purchased', clientReportPurchasedFilters)}
        ${renderClientRowsTable(filteredPurchasedRows, 'Brak kupionych indeksów z Top Rumunia dla wybranego klienta.')}
      </div>
      <div class="reports-column">
        <h3 class="reports-table-title">Brakujące z Top Rumunia</h3>
        <div class="reports-inline-actions">
          <button class="btn-outline" onclick="selectAllClientMissing()" ${clientReportMissingRows.length ? '' : 'disabled'}>Zaznacz wszystkie</button>
          <button class="btn-outline" onclick="clearAllClientMissing()" ${clientReportSelectedMissingIndexes.size ? '' : 'disabled'}>Odznacz wszystkie</button>
        </div>
        ${renderClientTableFilters('missing', clientReportMissingFilters)}
        ${renderClientRowsTable(filteredMissingRows, 'Klient kupił wszystkie wybrane indeksy z Top Rumunia.', { selectable: true })}
      </div>
    </div>
  `;
}

function renderTopSuggestionsContent(){
  const representatives = getTopSuggestionRepresentatives();
  const customers = getTopSuggestionCustomers();
  const summary = topSuggestionSummary
    ? `<div class="reports-summary">
        <div class="reports-summary-card">
          <div class="reports-summary-title">Ostatnie 3 tygodnie</div>
          <div class="reports-summary-meta">${topSuggestionSummary.weeks.map(escapeHtml).join(', ') || 'Brak sprzedaży'}</div>
        </div>
        <div class="reports-summary-card">
          <div class="reports-summary-title">Kupieni producenci</div>
          <div class="reports-summary-meta">${topSuggestionSummary.purchasedCount}</div>
        </div>
        <div class="reports-summary-card">
          <div class="reports-summary-title">Niekupieni producenci</div>
          <div class="reports-summary-meta">${topSuggestionSummary.shownCount} z ${topSuggestionSummary.missingCount}</div>
        </div>
      </div>`
    : '';

  return `
    <div class="reports-toolbar">
      <div>
        <div class="import-title">Propozycje Top Rumunia</div>
        <div class="import-sub">Porównaj zakupy klienta z ostatnich 3 tygodni z rankingiem Top Rumunia.</div>
      </div>
      <div class="reports-actions">
        <input id="top-suggestions-excel-input" class="import-input reports-file-input" type="file" accept=".xlsx,.xls" onchange="importTopSuggestionsExcel()">
        <button class="btn-outline" onclick="pickReportFile('top-suggestions-excel-input')">Importuj Excel</button>
        <button class="btn-outline" onclick="exportTopSuggestionsXlsx()" ${topSuggestionGeneratedRows.length ? '' : 'disabled'}>Zapisz XLSX</button>
      </div>
      <div class="import-info">
        ${topSuggestionImportFile ? 'Status: plik zaimportowany' : 'Status: oczekuje na import'}
      </div>
    </div>

    <div class="reports-filters">
      <label class="reports-filter-card">
        <span class="reports-limit-title">Przedstawiciel handlowy</span>
        <select onchange="setTopSuggestionRepresentative(this.value)">
          <option value="">Wybierz</option>
          ${representatives.map(rep => `<option value="${escapeAttr(rep)}" ${topSuggestionSelectedRepresentative === rep ? 'selected' : ''}>${escapeHtml(rep)}</option>`).join('')}
        </select>
      </label>
      <label class="reports-filter-card">
        <span class="reports-limit-title">Klient</span>
        <select onchange="setTopSuggestionCustomer(this.value)" ${topSuggestionSelectedRepresentative ? '' : 'disabled'}>
          <option value="">Wybierz</option>
          ${customers.map(customer => {
            const value = `${customer.code}|||${customer.name}`;
            const label = [customer.code, customer.name].filter(Boolean).join(' - ') || 'Bez nazwy klienta';
            return `<option value="${escapeAttr(value)}" ${topSuggestionSelectedCustomer === value ? 'selected' : ''}>${escapeHtml(label)}</option>`;
          }).join('')}
        </select>
      </label>
      <label class="reports-filter-card">
        <span class="reports-limit-title">Ile producentów pokazać</span>
        <input type="number" min="0" value="${escapeAttr(topSuggestionLimit)}" placeholder="Wszystkie" onchange="setTopSuggestionLimit(this.value)">
      </label>
    </div>

    ${summary}

    <div class="reports-two-columns">
      <div class="reports-column">
        <h3 class="reports-table-title">Kupieni producenci</h3>
        ${renderPurchasedProducerRowsTable(topSuggestionPurchasedRows, 'Brak kupionych producentów w ostatnich 3 tygodniach.')}
      </div>
      <div class="reports-column">
        <h3 class="reports-table-title">Producenci, których nie kupił</h3>
        ${renderMissingProducerRowsTable(topSuggestionGeneratedRows, 'Zaimportuj plik, wybierz PH i klienta, aby zobaczyć niekupionych producentów z Top Rumunia.')}
      </div>
    </div>
  `;
}

function getDetailedSalesTopImageUrl(index){
  const idx = String(index || '').trim();
  if(!idx) return '';
  const ext = getPreferredImageExt(idx, imageBaseUrlRumunia);
  return buildImageUrl(idx, ext, imageBaseUrlRumunia);
}

function buildDetailedSalesTopComparisonRow(row){
  return {
    INDEKS: row.index,
    NAZWA: row.name,
    SKROT_PRODUCENTA: row.producer,
    'KOD EAN': row.ean || '',
    Ranking: row.rankingLabel || (Number.isFinite(row.ranking) ? String(row.ranking) : ''),
    'Grupa produktowa': row.group || ''
  };
}

function isDetailedSalesTopMatch(topRow, purchasedIndexSet, purchasedNameSet, purchasedCompositeSet){
  const normalizedIndex = normalizeIndexValue(topRow.index);
  const normalizedName = normalizeDetailedSalesName(topRow.name);
  if(normalizedIndex && purchasedIndexSet.has(normalizedIndex)){
    return true;
  }
  if(normalizedIndex && normalizedName){
    if(purchasedCompositeSet.has(`${normalizedIndex}|||${normalizedName}`)){
      return true;
    }
  }
  return normalizedName ? purchasedNameSet.has(normalizedName) : false;
}

async function generateDetailedSalesTopComparison(options = {}){
  const suppressRender = Boolean(options.suppressRender);
  const requestId = ++detailedSalesTopComparisonRequestId;
  const previousSelectedGroupKey = detailedSalesTopSelectedGroupKey;
  const previousSelectedProducerKey = detailedSalesTopSelectedProducerKey;

  clearDetailedSalesTopComparisonResultState({ preserveCustomerSummaryRows: true });

  if(!detailedSalesSourceRows.length){
    detailedSalesTopCustomerSummaryRows = [];
    detailedSalesTopComparisonLoading = false;
    if(!suppressRender){
      renderReportsView();
    }
    return;
  }

  detailedSalesTopComparisonLoading = true;
  if(!suppressRender){
    renderReportsView();
  }

  try{
    const scopedTopRows = dedupeReportRows(
      (await loadTopRumuniaOfferRows()).filter(row => matchesDetailedSalesTopGroup(row.group, detailedSalesSelectedGroups))
    ).sort((a, b) => {
      if(a.ranking !== b.ranking) return a.ranking - b.ranking;
      return String(a.name).localeCompare(String(b.name), 'pl');
    });

    if(requestId !== detailedSalesTopComparisonRequestId){
      return;
    }

    detailedSalesTopCustomerSummaryRows = buildDetailedSalesTopCustomerSummaryRows(scopedTopRows);

    if(!detailedSalesComparisonCustomer){
      return;
    }

    const [customerCode, customerName] = detailedSalesComparisonCustomer.split('|||');
    const selectedCustomerRows = detailedSalesSourceRows.filter(row => (
      row.customerCode === customerCode && row.customerName === customerName
    ));
    const purchasedIndexSet = new Set(
      selectedCustomerRows.map(row => normalizeIndexValue(row.index)).filter(Boolean)
    );
    const purchasedNameSet = new Set(
      selectedCustomerRows.map(row => normalizeDetailedSalesName(row.name)).filter(Boolean)
    );
    const purchasedCompositeSet = new Set(
      selectedCustomerRows
        .map(row => {
          const normalizedIndex = normalizeIndexValue(row.index);
          const normalizedName = normalizeDetailedSalesName(row.name);
          if(!normalizedIndex || !normalizedName) return '';
          return `${normalizedIndex}|||${normalizedName}`;
        })
        .filter(Boolean)
    );

    const purchasedRows = [];
    const missingRows = [];

    scopedTopRows.forEach(row => {
      const enrichedRow = buildDetailedSalesTopComparisonRow(row);
      if(isDetailedSalesTopMatch(row, purchasedIndexSet, purchasedNameSet, purchasedCompositeSet)){
        purchasedRows.push(enrichedRow);
      }else{
        missingRows.push(enrichedRow);
      }
    });

    if(requestId !== detailedSalesTopComparisonRequestId){
      return;
    }

    detailedSalesTopPurchasedRows = purchasedRows;
    detailedSalesTopMissingRows = missingRows;
    const availableSelectionKeys = new Set(
      missingRows.map(getDetailedSalesTopMissingRowKey).filter(Boolean)
    );
    detailedSalesTopSelectedMissingKeys = new Set(
      Array.from(detailedSalesTopSelectedMissingKeys).filter(key => availableSelectionKeys.has(key))
    );
    detailedSalesTopGroupSummaryRows = buildDetailedSalesTopGroupSummaryRows(purchasedRows, missingRows);
    detailedSalesTopSelectedGroupKey = detailedSalesTopGroupSummaryRows.some(row => row.groupKey === previousSelectedGroupKey)
      ? previousSelectedGroupKey
      : (detailedSalesTopGroupSummaryRows[0]?.groupKey || '');
    detailedSalesTopProducerSummaryRows = buildDetailedSalesTopProducerSummaryRows(
      purchasedRows,
      missingRows,
      detailedSalesTopSelectedGroupKey
    );
    detailedSalesTopSelectedProducerKey = detailedSalesTopProducerSummaryRows.some(row => row.producerKey === previousSelectedProducerKey)
      ? previousSelectedProducerKey
      : (detailedSalesTopProducerSummaryRows[0]?.producerKey || '');
    detailedSalesTopSummary = {
      customerCode,
      customerName,
      groupLabel: getDetailedSalesSelectedGroupsLabel(3),
      purchasedCount: purchasedRows.length,
      missingCount: missingRows.length,
      offerCount: scopedTopRows.length
    };
  }catch(error){
    if(requestId !== detailedSalesTopComparisonRequestId){
      return;
    }
    console.error('Detailed sales Top Rumunia comparison error', error);
    clearDetailedSalesTopComparisonResultState();
    detailedSalesTopComparisonError = error?.message || 'Nie udało się porównać klienta z Top Rumunia.';
  }finally{
    if(requestId !== detailedSalesTopComparisonRequestId){
      return;
    }
    detailedSalesTopComparisonLoading = false;
    if(!suppressRender){
      renderReportsView();
    }
  }
}

function buildDetailedSalesTopCustomerSummaryRows(scopedTopRows){
  const topRows = Array.isArray(scopedTopRows) ? scopedTopRows : [];
  if(!topRows.length) return [];

  const groupOrder = new Map(REPORT_GROUP_CONFIGS.map((group, index) => [group.dashboardLabel, index]));

  const customerRowsMap = new Map();
  detailedSalesSourceRows.forEach(row => {
    const customerKey = getDetailedSalesCustomerKey(row.customerCode, row.customerName);
    if(!customerKey || customerKey === '|||') return;
    const current = customerRowsMap.get(customerKey) || {
      customerKey,
      customerCode: row.customerCode,
      customerName: row.customerName,
      rows: [],
      scopedSalesValue: 0
    };
    current.rows.push(row);
    if(matchesDetailedSalesRowGroup(row.groupName)){
      current.scopedSalesValue += Number(row.value || 0);
    }
    customerRowsMap.set(customerKey, current);
  });

  return Array.from(customerRowsMap.values())
    .map(customer => {
      const purchasedIndexSet = new Set(
        customer.rows.map(row => normalizeIndexValue(row.index)).filter(Boolean)
      );
      const purchasedNameSet = new Set(
        customer.rows.map(row => normalizeDetailedSalesName(row.name)).filter(Boolean)
      );
      const purchasedCompositeSet = new Set(
        customer.rows
          .map(row => {
            const normalizedIndex = normalizeIndexValue(row.index);
            const normalizedName = normalizeDetailedSalesName(row.name);
            if(!normalizedIndex || !normalizedName) return '';
            return `${normalizedIndex}|||${normalizedName}`;
          })
          .filter(Boolean)
      );

      let purchasedCount = 0;
      let missingCount = 0;
      const groupStats = new Map();

      topRows.forEach(topRow => {
        const isPurchased = isDetailedSalesTopMatch(topRow, purchasedIndexSet, purchasedNameSet, purchasedCompositeSet);
        const groupLabel = getDetailedSalesDashboardGroupLabel(topRow.group);
        const groupKey = getDetailedSalesTopGroupKey(groupLabel) || 'bez-grupy';
        const currentGroup = groupStats.get(groupKey) || {
          groupLabel,
          offerCount: 0,
          purchasedCount: 0
        };
        currentGroup.offerCount += 1;
        if(isPurchased){
          purchasedCount += 1;
          currentGroup.purchasedCount += 1;
        }else{
          missingCount += 1;
        }
        groupStats.set(groupKey, currentGroup);
      });

      const purchasedGroupCount = Array.from(groupStats.values()).filter(group => group.purchasedCount > 0).length;
      const missingOnlyGroupCount = Array.from(groupStats.values()).filter(group => group.purchasedCount === 0 && group.offerCount > 0).length;
      const coveragePercent = getDetailedSalesTopCoveragePercent(purchasedCount, topRows.length);
      const coverageTone = getDetailedSalesTopCoverageTone(coveragePercent);
      const groupBreakdown = Array.from(groupStats.values())
        .map(group => {
          const missingCount = Math.max(group.offerCount - group.purchasedCount, 0);
          const groupCoveragePercent = getDetailedSalesTopCoveragePercent(group.purchasedCount, group.offerCount);
          return {
            groupLabel: group.groupLabel,
            shortLabel: getDetailedSalesShortGroupLabel(group.groupLabel),
            offerCount: group.offerCount,
            purchasedCount: group.purchasedCount,
            missingCount,
            coveragePercent: groupCoveragePercent,
            coverageTone: getDetailedSalesTopCoverageTone(groupCoveragePercent)
          };
        })
        .sort((a, b) => {
          const aOrder = groupOrder.has(a.groupLabel) ? groupOrder.get(a.groupLabel) : Number.POSITIVE_INFINITY;
          const bOrder = groupOrder.has(b.groupLabel) ? groupOrder.get(b.groupLabel) : Number.POSITIVE_INFINITY;
          if(aOrder !== bOrder) return aOrder - bOrder;
          return String(a.groupLabel || '').localeCompare(String(b.groupLabel || ''), 'pl');
        });
      const topOpportunityGroup = groupBreakdown.reduce((best, group) => {
        if(!best) return group;
        if(group.missingCount !== best.missingCount) return group.missingCount > best.missingCount ? group : best;
        if(group.offerCount !== best.offerCount) return group.offerCount > best.offerCount ? group : best;
        return String(group.groupLabel || '').localeCompare(String(best.groupLabel || ''), 'pl') < 0 ? group : best;
      }, null);

      return {
        customerKey: customer.customerKey,
        customerCode: customer.customerCode,
        customerName: customer.customerName,
        offerCount: topRows.length,
        purchasedCount,
        missingCount,
        purchasedGroupCount,
        missingOnlyGroupCount,
        totalGroupCount: groupBreakdown.length,
        coveragePercent,
        coverageTone,
        salesValue: customer.scopedSalesValue,
        groupBreakdown,
        topOpportunityGroupLabel: topOpportunityGroup?.groupLabel || '',
        topOpportunityMissingCount: topOpportunityGroup?.missingCount || 0
      };
    })
    .sort((a, b) => {
      if(b.purchasedCount !== a.purchasedCount) return b.purchasedCount - a.purchasedCount;
      if(b.coveragePercent !== a.coveragePercent) return b.coveragePercent - a.coveragePercent;
      if(b.salesValue !== a.salesValue) return b.salesValue - a.salesValue;
      const nameCompare = String(a.customerName || '').localeCompare(String(b.customerName || ''), 'pl');
      if(nameCompare) return nameCompare;
      return String(a.customerCode || '').localeCompare(String(b.customerCode || ''), 'pl', { numeric: true });
    });
}

function buildDetailedSalesTopMeetingRecommendationRows(scopedTopRows){
  const topRows = Array.isArray(scopedTopRows) ? scopedTopRows : [];
  if(!topRows.length || !detailedSalesSourceRows.length) return [];

  const customerPurchaseMap = new Map();
  detailedSalesSourceRows.forEach(row => {
    const customerKey = getDetailedSalesCustomerKey(row.customerCode, row.customerName);
    if(!customerKey || customerKey === '|||') return;
    const current = customerPurchaseMap.get(customerKey) || {
      purchasedIndexSet: new Set(),
      purchasedNameSet: new Set(),
      purchasedCompositeSet: new Set()
    };
    const normalizedIndex = normalizeIndexValue(row.index);
    const normalizedName = normalizeDetailedSalesName(row.name);
    if(normalizedIndex){
      current.purchasedIndexSet.add(normalizedIndex);
    }
    if(normalizedName){
      current.purchasedNameSet.add(normalizedName);
    }
    if(normalizedIndex && normalizedName){
      current.purchasedCompositeSet.add(`${normalizedIndex}|||${normalizedName}`);
    }
    customerPurchaseMap.set(customerKey, current);
  });

  const totalCustomers = customerPurchaseMap.size;
  if(!totalCustomers) return [];

  return topRows
    .map(row => {
      let purchasedCustomers = 0;
      customerPurchaseMap.forEach(customer => {
        if(isDetailedSalesTopMatch(row, customer.purchasedIndexSet, customer.purchasedNameSet, customer.purchasedCompositeSet)){
          purchasedCustomers += 1;
        }
      });
      const missingCustomers = Math.max(totalCustomers - purchasedCustomers, 0);
      return {
        index: row.index,
        name: row.name,
        producer: row.producer,
        groupLabel: getDetailedSalesDashboardGroupLabel(row.group),
        shortGroupLabel: getDetailedSalesShortGroupLabel(row.group),
        ranking: Number.isFinite(row.ranking) ? row.ranking : Number.POSITIVE_INFINITY,
        rankingLabel: row.rankingLabel || (Number.isFinite(row.ranking) ? String(row.ranking) : ''),
        purchasedCustomers,
        missingCustomers,
        totalCustomers
      };
    })
    .filter(row => row.missingCustomers > 0)
    .sort((a, b) => {
      if(a.ranking !== b.ranking) return a.ranking - b.ranking;
      if(b.missingCustomers !== a.missingCustomers) return b.missingCustomers - a.missingCustomers;
      if(a.groupLabel !== b.groupLabel) return String(a.groupLabel || '').localeCompare(String(b.groupLabel || ''), 'pl');
      return String(a.name || '').localeCompare(String(b.name || ''), 'pl');
    })
    .slice(0, 40);
}

function getFilteredDetailedSalesTopRows(rows, filters){
  return rows.filter(row => {
    if(filters.producer && !includesText(row.SKROT_PRODUCENTA, filters.producer)) return false;
    if(filters.group && !includesText(row['Grupa produktowa'], filters.group)) return false;
    return true;
  });
}

function renderDetailedSalesTopTableFilters(type, rows, filters){
  const producerOptions = Array.from(new Set(
    (rows || []).map(row => String(row.SKROT_PRODUCENTA || '').trim()).filter(Boolean)
  ))
    .sort((a, b) => a.localeCompare(b, 'pl'))
    .map(producer => `<option value="${escapeAttr(producer)}" ${filters.producer === producer ? 'selected' : ''}>${escapeHtml(producer)}</option>`)
    .join('');
  const groupOptions = Array.from(new Set(
    (rows || []).map(row => String(row['Grupa produktowa'] || '').trim()).filter(Boolean)
  ))
    .sort((a, b) => a.localeCompare(b, 'pl'))
    .map(group => `<option value="${escapeAttr(group)}" ${filters.group === group ? 'selected' : ''}>${escapeHtml(group)}</option>`)
    .join('');

  return `
    <div class="reports-table-filters">
      <select onchange="setDetailedSalesTopTableFilter('${type}', 'producer', this.value)">
        <option value="">Wszyscy producenci</option>
        ${producerOptions}
      </select>
      <select onchange="setDetailedSalesTopTableFilter('${type}', 'group', this.value)">
        <option value="">Wszystkie grupy</option>
        ${groupOptions}
      </select>
    </div>
  `;
}

function getDetailedSalesTopMissingRowKey(row){
  const normalizedIndex = normalizeIndexValue(row?.INDEKS);
  const normalizedName = normalizeDetailedSalesName(row?.NAZWA);
  const producerKey = normalizeProducerValue(row?.SKROT_PRODUCENTA) || String(row?.SKROT_PRODUCENTA || '').trim().toLowerCase();
  const groupKey = String(row?.['Grupa produktowa'] || '').trim().toLowerCase();
  if(!normalizedIndex && !normalizedName) return '';
  return [normalizedIndex || '', normalizedName || '', producerKey || '', groupKey || ''].join('|||');
}

function isDetailedSalesTopMissingRowSelected(row){
  const rowKey = getDetailedSalesTopMissingRowKey(row);
  return !!rowKey && detailedSalesTopSelectedMissingKeys.has(rowKey);
}

function getSelectedDetailedSalesTopMissingRows(){
  return detailedSalesTopMissingRows.filter(row => isDetailedSalesTopMissingRowSelected(row));
}

function getDetailedSalesTopVisibleMissingRows(scope = 'filtered'){
  if(scope === 'producer'){
    return getDetailedSalesTopMissingRowsForSelectedProducer();
  }
  return getFilteredDetailedSalesTopRows(detailedSalesTopMissingRows, detailedSalesTopMissingFilters);
}

function getDetailedSalesTopRecommendationFileName(){
  const [customerCode = '', customerName = ''] = String(detailedSalesComparisonCustomer || '').split('|||');
  const safePart = (value, fallback) => String(value || fallback)
    .trim()
    .replace(/[\\/:*?"<>|]+/g, ' ')
    .replace(/\s+/g, ' ');
  return `${safePart(customerCode, 'Klient')} ${safePart(customerName, 'Bez nazwy')} - rekomendacje Top Rumunia`;
}

function getDetailedSalesTopMissingExportRows(){
  return getSelectedDetailedSalesTopMissingRows().map(row => ({
    INDEKS: row.INDEKS ?? '',
    NAZWA: row.NAZWA ?? '',
    SKROT_PRODUCENTA: row.SKROT_PRODUCENTA ?? '',
    Ranking: row.Ranking ?? '',
    'Grupa produktowa': row['Grupa produktowa'] ?? '',
    'Zdjęcie URL': getDetailedSalesTopImageUrl(row.INDEKS)
  }));
}

function renderDetailedSalesTopSelectionToolbar(summaryText, actionsHtml = '', className = ''){
  return `
    <div class="reports-selection-toolbar${className ? ` ${className}` : ''}">
      <div class="reports-selection-summary">${escapeHtml(summaryText)}</div>
      ${actionsHtml ? `<div class="reports-selection-actions">${actionsHtml}</div>` : ''}
    </div>
  `;
}

function renderDetailedSalesTopTableToolbar(summaryText, actionsHtml = ''){
  return `
    <div class="reports-selection-toolbar reports-selection-toolbar--table">
      <div class="reports-selection-summary">${escapeHtml(summaryText)}</div>
      ${actionsHtml
        ? `<div class="reports-selection-actions">${actionsHtml}</div>`
        : `<div class="reports-selection-actions reports-selection-actions--placeholder" aria-hidden="true"></div>`
      }
    </div>
  `;
}

function renderDetailedSalesTopInfoToolbar(rows, options = {}){
  const {
    emptyText = 'Brak pozycji w tym widoku.',
    label = 'Widoczne'
  } = options;
  const visibleRows = Array.isArray(rows) ? rows : [];
  const summaryText = visibleRows.length
    ? `${label}: ${visibleRows.length}`
    : emptyText;

  return renderDetailedSalesTopTableToolbar(summaryText);
}

function renderDetailedSalesTopExportToolbar(){
  const selectedCount = getSelectedDetailedSalesTopMissingRows().length;
  const canExportRepresentativePdf = !!detailedSalesSourceRows.length;
  const summaryText = selectedCount
    ? `Zaznaczone rekomendacje: ${selectedCount}`
    : 'Zaznacz brakujące produkty w tabelach poniżej, aby zapisać listę lub utworzyć katalog.';

  return renderDetailedSalesTopSelectionToolbar(summaryText, `
    <button class="btn-outline" onclick="clearAllDetailedSalesTopMissing()" ${selectedCount ? '' : 'disabled'}>Wyczyść zaznaczenie</button>
    <button class="btn-outline" onclick="exportDetailedSalesTopMissingCsv()" ${selectedCount ? '' : 'disabled'}>Zapisz CSV</button>
    <button class="btn-outline" onclick="exportDetailedSalesTopMissingXls()" ${selectedCount ? '' : 'disabled'}>Zapisz XLS</button>
    <button class="btn-outline" onclick="createDetailedSalesTopRecommendationCatalog()" ${selectedCount ? '' : 'disabled'}>Utwórz katalog</button>
    <button class="btn-outline" onclick="exportDetailedSalesRepresentativeSummaryPdf()" ${canExportRepresentativePdf ? '' : 'disabled'}>Raport PDF PH</button>
  `, 'reports-selection-toolbar--export');
}

function renderDetailedSalesTopQuickSelectionToolbar(rows, options = {}){
  const {
    scope = 'filtered',
    selectLabel = 'Zaznacz widoczne',
    clearLabel = 'Odznacz widoczne'
  } = options;
  const visibleRows = Array.isArray(rows) ? rows : [];
  const visibleSelectedCount = visibleRows.filter(row => isDetailedSalesTopMissingRowSelected(row)).length;
  const selectedCount = getSelectedDetailedSalesTopMissingRows().length;
  const summaryText = visibleRows.length
    ? `Widoczne: ${visibleRows.length} | zaznaczone w tym widoku: ${visibleSelectedCount} | łącznie: ${selectedCount}`
    : `Brak pozycji do zaznaczenia${selectedCount ? ` | łącznie zaznaczone: ${selectedCount}` : ''}`;

  return renderDetailedSalesTopTableToolbar(summaryText, `
    <button class="btn-outline" onclick="selectVisibleDetailedSalesTopMissing('${scope}')" ${visibleRows.length ? '' : 'disabled'}>${escapeHtml(selectLabel)}</button>
    <button class="btn-outline" onclick="clearVisibleDetailedSalesTopMissing('${scope}')" ${visibleSelectedCount ? '' : 'disabled'}>${escapeHtml(clearLabel)}</button>
  `);
}

function renderDetailedSalesTopRowsTable(rows, emptyText, options = {}){
  const { selectable = false } = options;
  const html = rows.map((row, index) => {
    const idx = String(row.INDEKS ?? '').trim();
    const rowKey = getDetailedSalesTopMissingRowKey(row);
    const canSelect = selectable && !!rowKey;
    const isSelected = canSelect && detailedSalesTopSelectedMissingKeys.has(rowKey);
    const rowClass = canSelect ? ` class="reports-selectable-row${isSelected ? ' is-active' : ''}"` : '';
    const rowAttrs = canSelect
      ? ` data-selection-key="${escapeAttr(rowKey)}" onclick="toggleDetailedSalesTopMissingSelectionByKey(this.getAttribute('data-selection-key'))"`
      : '';
    const checked = isSelected ? 'checked' : '';

    return `
      <tr${rowClass}${rowAttrs}>
        <td>${index + 1}</td>
        ${selectable ? `
          <td class="reports-selection-cell">
            <input
              type="checkbox"
              class="reports-row-checkbox"
              data-selection-key="${escapeAttr(rowKey)}"
              onchange="setDetailedSalesTopMissingSelectionByKey(this.getAttribute('data-selection-key'), this.checked)"
              onclick="event.stopPropagation()"
              ${checked}
              ${canSelect ? '' : 'disabled'}
            >
          </td>
        ` : ''}
        <td class="reports-image-cell">
          ${idx ? buildProductImageTag(idx, imageBaseUrlRumunia, 'reports-image') : ''}
        </td>
        <td>${escapeHtml(row.INDEKS ?? '')}</td>
        <td>${escapeHtml(row.NAZWA ?? '')}</td>
        <td>${escapeHtml(row.SKROT_PRODUCENTA ?? '')}</td>
        <td>${escapeHtml(row.Ranking ?? '')}</td>
        <td>${escapeHtml(row['Grupa produktowa'] ?? '')}</td>
      </tr>
    `;
  }).join('');

  return `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Lp.</th>
            ${selectable ? '<th>Wybór</th>' : ''}
            <th>Zdjęcie</th>
            <th>INDEKS</th>
            <th>NAZWA</th>
            <th>SKROT_PRODUCENTA</th>
            <th>Ranking</th>
            <th>Grupa produktowa</th>
          </tr>
        </thead>
        <tbody>
          ${html || `<tr><td colspan="${selectable ? 8 : 7}" class="reports-empty">${escapeHtml(emptyText)}</td></tr>`}
        </tbody>
      </table>
    </div>
  `;
}
function buildDetailedSalesTopGroupSummaryRows(purchasedRows, missingRows){
  const summaryMap = new Map();

  const ensureGroupRow = (row) => {
    const groupLabel = getDetailedSalesDashboardGroupLabel(row['Grupa produktowa']);
    const groupKey = getDetailedSalesTopGroupKey(groupLabel);
    if(!groupKey) return null;
    const current = summaryMap.get(groupKey) || {
      groupKey,
      groupLabel,
      purchasedIndexSet: new Set(),
      offerIndexSet: new Set()
    };
    summaryMap.set(groupKey, current);
    return current;
  };

  (purchasedRows || []).forEach(row => {
    const current = ensureGroupRow(row);
    const indexKey = getDetailedSalesTopMissingRowKey(row);
    if(!current || !indexKey) return;
    current.purchasedIndexSet.add(indexKey);
    current.offerIndexSet.add(indexKey);
  });

  (missingRows || []).forEach(row => {
    const current = ensureGroupRow(row);
    const indexKey = getDetailedSalesTopMissingRowKey(row);
    if(!current || !indexKey) return;
    current.offerIndexSet.add(indexKey);
  });

  return Array.from(summaryMap.values())
    .map(row => ({
      groupKey: row.groupKey,
      groupLabel: row.groupLabel,
      purchasedCount: row.purchasedIndexSet.size,
      offerCount: row.offerIndexSet.size,
      missingCount: Math.max(0, row.offerIndexSet.size - row.purchasedIndexSet.size)
    }))
    .sort((a, b) => {
      if(b.missingCount !== a.missingCount) return b.missingCount - a.missingCount;
      if(a.purchasedCount !== b.purchasedCount) return a.purchasedCount - b.purchasedCount;
      if(b.offerCount !== a.offerCount) return b.offerCount - a.offerCount;
      return String(a.groupLabel).localeCompare(String(b.groupLabel), 'pl');
    });
}

function renderDetailedSalesTopSectionSwitch(){
  return `
    <div class="reports-mode-switch reports-mode-switch--secondary">
      <button
        type="button"
        class="btn-outline ${detailedSalesTopSection === 'summary' ? 'reports-mode-active' : ''}"
        onclick="setDetailedSalesTopSection('summary')"
      >
        Skrót klienta
      </button>
      <button
        type="button"
        class="btn-outline ${detailedSalesTopSection === 'full' ? 'reports-mode-active' : ''}"
        onclick="setDetailedSalesTopSection('full')"
      >
        Pełna lista Top
      </button>
    </div>
  `;
}

function getDetailedSalesTopCustomerLabel(row){
  return [row?.customerCode, row?.customerName].filter(Boolean).join(' - ') || 'Bez nazwy klienta';
}

function sortDetailedSalesTopCustomerSummaryRows(rows){
  const source = Array.isArray(rows) ? [...rows] : [];
  const compareText = (a, b) => {
    const nameCompare = String(a?.customerName || '').localeCompare(String(b?.customerName || ''), 'pl');
    if(nameCompare) return nameCompare;
    return String(a?.customerCode || '').localeCompare(String(b?.customerCode || ''), 'pl', { numeric: true });
  };

  source.sort((a, b) => {
    switch(detailedSalesTopCustomerSort){
      case 'purchased-asc':
        if(a.purchasedCount !== b.purchasedCount) return a.purchasedCount - b.purchasedCount;
        break;
      case 'coverage-desc':
        if(a.coveragePercent !== b.coveragePercent) return b.coveragePercent - a.coveragePercent;
        break;
      case 'coverage-asc':
        if(a.coveragePercent !== b.coveragePercent) return a.coveragePercent - b.coveragePercent;
        break;
      case 'groups-desc':
        if(a.purchasedGroupCount !== b.purchasedGroupCount) return b.purchasedGroupCount - a.purchasedGroupCount;
        break;
      case 'sales-desc':
        if(a.salesValue !== b.salesValue) return b.salesValue - a.salesValue;
        break;
      case 'missing-desc':
        if(a.missingCount !== b.missingCount) return b.missingCount - a.missingCount;
        break;
      case 'purchased-desc':
      default:
        if(a.purchasedCount !== b.purchasedCount) return b.purchasedCount - a.purchasedCount;
        break;
    }

    if(b.coveragePercent !== a.coveragePercent) return b.coveragePercent - a.coveragePercent;
    if(b.salesValue !== a.salesValue) return b.salesValue - a.salesValue;
    return compareText(a, b);
  });

  return source;
}

function getFilteredDetailedSalesTopCustomerSummaryRows(){
  let rows = Array.isArray(detailedSalesTopCustomerSummaryRows)
    ? [...detailedSalesTopCustomerSummaryRows]
    : [];

  if(detailedSalesTopCustomerSearch){
    rows = rows.filter(row => includesText(
      `${row.customerCode || ''} ${row.customerName || ''}`,
      detailedSalesTopCustomerSearch
    ));
  }

  if(detailedSalesTopCustomerToneFilter){
    rows = rows.filter(row => row.coverageTone === detailedSalesTopCustomerToneFilter);
  }

  return sortDetailedSalesTopCustomerSummaryRows(rows);
}

function renderDetailedSalesTopCustomerPortfolioSection(){
  const allRows = Array.isArray(detailedSalesTopCustomerSummaryRows)
    ? detailedSalesTopCustomerSummaryRows
    : [];
  const visibleRows = getFilteredDetailedSalesTopCustomerSummaryRows();
  const summaryRows = visibleRows.length ? visibleRows : allRows;
  const isExpanded = Boolean(detailedSalesTopCustomerPortfolioExpanded);
  const strongestRow = [...summaryRows].sort((a, b) => b.coveragePercent - a.coveragePercent || b.purchasedCount - a.purchasedCount)[0] || null;
  const biggestRow = [...summaryRows].sort((a, b) => b.purchasedCount - a.purchasedCount || b.salesValue - a.salesValue)[0] || null;
  const potentialRow = [...summaryRows].sort((a, b) => b.missingCount - a.missingCount || a.coveragePercent - b.coveragePercent)[0] || null;
  const activeCustomerRow = allRows.find(row => row.customerKey === detailedSalesComparisonCustomer) || null;
  const toggleLabel = isExpanded
    ? 'Zwiń listę klientów'
    : 'Rozwiń listę klientów';
  const currentVisibleCount = isExpanded ? visibleRows.length : allRows.length;

  const rowsHtml = visibleRows.map(row => {
    const isActive = row.customerKey === detailedSalesComparisonCustomer;
    const activeClass = isActive ? ' is-active' : '';
    const tone = row.coverageTone || 'low';
    return `
      <button
        type="button"
        class="reports-overview-row reports-customer-row is-${tone}${activeClass}"
        onclick="setDetailedSalesComparisonCustomer(${escapeAttr(JSON.stringify(row.customerKey))})"
      >
        <span class="reports-overview-row-main">
          <span class="reports-customer-row-heading">
            <span class="reports-customer-row-code">${escapeHtml(row.customerCode || 'Brak kodu')}</span>
            <span class="reports-overview-row-title">${escapeHtml(row.customerName || 'Bez nazwy klienta')}</span>
          </span>
          <span class="reports-customer-row-sub">Sprzedaż w zakresie: ${formatNumber(row.salesValue)}</span>
          <span class="reports-overview-row-copy reports-overview-row-copy--stats">
            <span class="reports-mini-stat">
              <strong>${formatNumber(row.purchasedCount)}</strong>
              <em>kupione</em>
            </span>
            <span class="reports-mini-stat is-gap">
              <strong>${formatNumber(row.missingCount)}</strong>
              <em>braki</em>
            </span>
            <span class="reports-mini-stat">
              <strong>${formatNumber(row.purchasedGroupCount)}</strong>
              <em>kategorie</em>
            </span>
            <span class="reports-mini-stat">
              <strong>${formatNumber(row.offerCount)}</strong>
              <em>w top</em>
            </span>
          </span>
        </span>
        <span class="reports-overview-row-side">
          ${isActive ? `<span class="reports-active-pill">Wybrany klient</span>` : ''}
          <span class="reports-tone-badge is-${tone}">${escapeHtml(getDetailedSalesTopCoverageLabel(row.coveragePercent))}</span>
          <span class="reports-overview-row-percent">${escapeHtml(formatPercent(row.coveragePercent, { minimumFractionDigits: 1, maximumFractionDigits: 1 }))}</span>
          <span class="reports-progress-track">
            <span class="reports-progress-fill is-${tone}" style="width:${Math.max(0, Math.min(row.coveragePercent, 100))}%"></span>
          </span>
        </span>
      </button>
    `;
  }).join('');

  const emptyMessage = detailedSalesTopComparisonLoading && !allRows.length
    ? 'Liczę ranking klientów względem Top Rumunia.'
    : 'Brak klientów pasujących do wybranych filtrów.';

  return `
    <div class="reports-analysis-card reports-analysis-card--customer-portfolio${isExpanded ? ' is-expanded' : ' is-collapsed'}">
      <div class="reports-analysis-card-head reports-analysis-card-head--wide reports-analysis-card-head--portfolio">
        <div>
          <div class="reports-analysis-card-title">Portfel klientów z raportu</div>
          <div class="reports-analysis-card-sub">Lista jest na górze, ale schowana na start. Rozwiń ją, gdy chcesz szybko przełączyć się między klientami.</div>
        </div>
        <button
          type="button"
          class="btn-outline reports-customer-portfolio-toggle"
          onclick="toggleDetailedSalesTopCustomerPortfolioExpanded()"
          ${allRows.length ? '' : 'disabled'}
        >
          ${escapeHtml(toggleLabel)}
        </button>
      </div>

      <div class="reports-customer-portfolio-bar">
        <span class="reports-mini-stat">
          <strong>${formatNumber(currentVisibleCount)}</strong>
          <em>${isExpanded ? 'klientów w widoku' : 'klientów'}</em>
        </span>
        ${activeCustomerRow ? `
          <span class="reports-mini-stat">
            <strong>${escapeHtml(activeCustomerRow.customerCode || 'Klient')}</strong>
            <em>aktywny klient</em>
          </span>
        ` : ''}
        ${strongestRow ? `
          <span class="reports-mini-stat">
            <strong>${escapeHtml(formatPercent(strongestRow.coveragePercent, { minimumFractionDigits: 1, maximumFractionDigits: 1 }))}</strong>
            <em>najwyższy udział</em>
          </span>
        ` : ''}
        ${biggestRow ? `
          <span class="reports-mini-stat">
            <strong>${formatNumber(biggestRow.purchasedCount)}</strong>
            <em>max kupione z top</em>
          </span>
        ` : ''}
      </div>

      ${isExpanded ? `
        <div class="reports-summary reports-summary--compact">
          <div class="reports-summary-card">
            <div class="reports-summary-title">Klienci w widoku</div>
            <div class="reports-summary-meta">${formatNumber(visibleRows.length)} z ${formatNumber(allRows.length)}</div>
          </div>
          <div class="reports-summary-card">
            <div class="reports-summary-title">Najwyższy udział</div>
            <div class="reports-summary-meta">${strongestRow ? `${escapeHtml(getDetailedSalesTopCustomerLabel(strongestRow))} | ${escapeHtml(formatPercent(strongestRow.coveragePercent, { minimumFractionDigits: 1, maximumFractionDigits: 1 }))}` : 'Brak danych'}</div>
          </div>
          <div class="reports-summary-card">
            <div class="reports-summary-title">Najwięcej kupionych z Top</div>
            <div class="reports-summary-meta">${biggestRow ? `${escapeHtml(getDetailedSalesTopCustomerLabel(biggestRow))} | ${formatNumber(biggestRow.purchasedCount)}` : 'Brak danych'}</div>
          </div>
          <div class="reports-summary-card">
            <div class="reports-summary-title">Największy potencjał</div>
            <div class="reports-summary-meta">${potentialRow ? `${escapeHtml(getDetailedSalesTopCustomerLabel(potentialRow))} | braki ${formatNumber(potentialRow.missingCount)}` : 'Brak danych'}</div>
          </div>
        </div>

        <div class="reports-filters reports-filters--customer-portfolio">
          <label class="reports-filter-card">
            <span class="reports-limit-title">Szukaj klienta</span>
            <input
              type="text"
              value="${escapeAttr(detailedSalesTopCustomerSearch)}"
              placeholder="Kod lub nazwa klienta"
              oninput="setDetailedSalesTopCustomerSearch(this.value)"
            >
          </label>
          <label class="reports-filter-card">
            <span class="reports-limit-title">Filtr udziału</span>
            <select onchange="setDetailedSalesTopCustomerToneFilter(this.value)">
              <option value="">Wszyscy klienci</option>
              <option value="high" ${detailedSalesTopCustomerToneFilter === 'high' ? 'selected' : ''}>Wysoki udział</option>
              <option value="medium" ${detailedSalesTopCustomerToneFilter === 'medium' ? 'selected' : ''}>Średni udział</option>
              <option value="low" ${detailedSalesTopCustomerToneFilter === 'low' ? 'selected' : ''}>Niski udział</option>
            </select>
          </label>
          <label class="reports-filter-card">
            <span class="reports-limit-title">Sortowanie klientów</span>
            <select onchange="setDetailedSalesTopCustomerSort(this.value)">
              <option value="purchased-desc" ${detailedSalesTopCustomerSort === 'purchased-desc' ? 'selected' : ''}>Kupione z Top: od największej do najmniejszej</option>
              <option value="purchased-asc" ${detailedSalesTopCustomerSort === 'purchased-asc' ? 'selected' : ''}>Kupione z Top: od najmniejszej do największej</option>
              <option value="coverage-desc" ${detailedSalesTopCustomerSort === 'coverage-desc' ? 'selected' : ''}>Udział Top: od największego do najmniejszego</option>
              <option value="coverage-asc" ${detailedSalesTopCustomerSort === 'coverage-asc' ? 'selected' : ''}>Udział Top: od najmniejszego do największego</option>
              <option value="groups-desc" ${detailedSalesTopCustomerSort === 'groups-desc' ? 'selected' : ''}>Kupowane kategorie: od największej liczby</option>
              <option value="sales-desc" ${detailedSalesTopCustomerSort === 'sales-desc' ? 'selected' : ''}>Sprzedaż w zakresie: od największej do najmniejszej</option>
              <option value="missing-desc" ${detailedSalesTopCustomerSort === 'missing-desc' ? 'selected' : ''}>Potencjał: najwięcej braków</option>
            </select>
          </label>
        </div>

        <div class="reports-customer-ranking-list">
          ${rowsHtml || `<div class="reports-analysis-empty">${escapeHtml(emptyMessage)}</div>`}
        </div>
      ` : ''}
    </div>
  `;
}

function renderDetailedSalesTopFullComparisonSection(filteredPurchasedRows, filteredMissingRows){
  return `
    <div class="reports-two-columns">
      <div class="reports-column">
        <h3 class="reports-table-title">Kupione z Top Rumunia</h3>
        ${renderDetailedSalesTopTableFilters('purchased', detailedSalesTopPurchasedRows, detailedSalesTopPurchasedFilters)}
        ${renderDetailedSalesTopInfoToolbar(filteredPurchasedRows, {
          emptyText: 'Brak kupionych pozycji w tym widoku.',
          label: 'Widoczne'
        })}
        ${renderDetailedSalesTopRowsTable(filteredPurchasedRows, 'Klient nie kupił żadnego indeksu z Top Rumunia w wybranym zakresie.')}
      </div>
      <div class="reports-column">
        <h3 class="reports-table-title">Brakujące z Top Rumunia</h3>
        ${renderDetailedSalesTopTableFilters('missing', detailedSalesTopMissingRows, detailedSalesTopMissingFilters)}
        ${renderDetailedSalesTopQuickSelectionToolbar(filteredMissingRows, {
          scope: 'filtered',
          selectLabel: 'Zaznacz widoczne',
          clearLabel: 'Odznacz widoczne'
        })}
        ${renderDetailedSalesTopRowsTable(filteredMissingRows, 'Klient kupił wszystkie indeksy z Top Rumunia w wybranym zakresie.', { selectable: true })}
      </div>
    </div>
  `;
}

function buildDetailedSalesTopProducerSummaryRows(purchasedRows, missingRows, selectedGroupKey = ''){
  const summaryMap = new Map();
  const normalizedGroupKey = getDetailedSalesTopGroupKey(selectedGroupKey);

  const ensureProducerSummary = (row) => {
    const rowGroupKey = getDetailedSalesTopGroupKey(getDetailedSalesDashboardGroupLabel(row['Grupa produktowa']));
    if(normalizedGroupKey && rowGroupKey !== normalizedGroupKey) return;
    const producer = String(row.SKROT_PRODUCENTA || '').trim() || 'Bez producenta';
    const producerKey = normalizeProducerValue(producer) || producer.toLowerCase();
    if(!producerKey) return null;

    const current = summaryMap.get(producerKey) || {
      producerKey,
      producer,
      purchasedIndexSet: new Set(),
      missingIndexSet: new Set()
    };
    summaryMap.set(producerKey, current);
    return current;
  };

  (purchasedRows || []).forEach(row => {
    const current = ensureProducerSummary(row);
    const indexKey = getDetailedSalesTopMissingRowKey(row);
    if(!current || !indexKey) return;
    current.purchasedIndexSet.add(indexKey);
  });

  (missingRows || []).forEach(row => {
    const current = ensureProducerSummary(row);
    const indexKey = getDetailedSalesTopMissingRowKey(row);
    if(!current || !indexKey) return;
    current.missingIndexSet.add(indexKey);
  });

  return Array.from(summaryMap.values())
    .map(row => ({
      producerKey: row.producerKey,
      producer: row.producer,
      purchasedCount: row.purchasedIndexSet.size,
      missingCount: row.missingIndexSet.size,
      totalCount: row.purchasedIndexSet.size + row.missingIndexSet.size
    }))
    .sort((a, b) => {
      if(b.purchasedCount !== a.purchasedCount) return b.purchasedCount - a.purchasedCount;
      if(b.missingCount !== a.missingCount) return b.missingCount - a.missingCount;
      return String(a.producer).localeCompare(String(b.producer), 'pl');
    });
}

function getDetailedSalesTopMissingRowsForSelectedProducer(){
  if(!detailedSalesTopSelectedGroupKey || !detailedSalesTopSelectedProducerKey){
    return [];
  }
  return detailedSalesTopMissingRows.filter(row => {
    const groupKey = getDetailedSalesTopGroupKey(getDetailedSalesDashboardGroupLabel(row['Grupa produktowa']));
    const producer = String(row.SKROT_PRODUCENTA || '').trim() || 'Bez producenta';
    const producerKey = normalizeProducerValue(producer) || producer.toLowerCase();
    return groupKey === detailedSalesTopSelectedGroupKey && producerKey === detailedSalesTopSelectedProducerKey;
  });
}

function getDetailedSalesTopPurchasedRowsForSelectedProducer(){
  if(!detailedSalesTopSelectedGroupKey || !detailedSalesTopSelectedProducerKey){
    return [];
  }
  return detailedSalesTopPurchasedRows.filter(row => {
    const groupKey = getDetailedSalesTopGroupKey(getDetailedSalesDashboardGroupLabel(row['Grupa produktowa']));
    const producer = String(row.SKROT_PRODUCENTA || '').trim() || 'Bez producenta';
    const producerKey = normalizeProducerValue(producer) || producer.toLowerCase();
    return groupKey === detailedSalesTopSelectedGroupKey && producerKey === detailedSalesTopSelectedProducerKey;
  });
}

function renderDetailedSalesTopGroupButtons(rows){
  return (rows || []).map(row => {
    const activeClass = row.groupKey === detailedSalesTopSelectedGroupKey ? ' is-active' : '';
    const serializedKey = escapeAttr(JSON.stringify(row.groupKey));
    return `
      <button
        type="button"
        class="reports-producer-chip${activeClass}"
        onclick="setDetailedSalesTopGroup(${serializedKey})"
      >
        <span class="reports-producer-chip-name">${escapeHtml(row.groupLabel)}</span>
        <span class="reports-producer-chip-meta">Kupione: ${row.purchasedCount} | Dostępne: ${row.offerCount}</span>
      </button>
    `;
  }).join('');
}

function renderDetailedSalesTopProducerButtons(rows){
  return (rows || []).map(row => {
    const activeClass = row.producerKey === detailedSalesTopSelectedProducerKey ? ' is-active' : '';
    const serializedKey = escapeAttr(JSON.stringify(row.producerKey));
    return `
      <button
        type="button"
        class="reports-producer-chip${activeClass}"
        onclick="setDetailedSalesTopProducer(${serializedKey})"
      >
        <span class="reports-producer-chip-name">${escapeHtml(row.producer)}</span>
        <span class="reports-producer-chip-meta">Kupione: ${row.purchasedCount} | Braki: ${row.missingCount}</span>
      </button>
    `;
  }).join('');
}

function renderDetailedSalesTopProducerComparisonSection(){
  const allGroupRows = detailedSalesTopGroupSummaryRows;
  const purchasedGroupRows = allGroupRows.filter(row => row.purchasedCount > 0);
  const nonPurchasedGroupRows = allGroupRows.filter(row => row.purchasedCount === 0 && row.offerCount > 0);
  const overviewGroupRows = allGroupRows;
  const selectedGroupRow = allGroupRows.find(row => row.groupKey === detailedSalesTopSelectedGroupKey) || null;
  const selectedProducerRow = detailedSalesTopProducerSummaryRows.find(row => row.producerKey === detailedSalesTopSelectedProducerKey) || null;
  const purchasedRows = getDetailedSalesTopPurchasedRowsForSelectedProducer();
  const missingRows = getDetailedSalesTopMissingRowsForSelectedProducer();
  const filteredPurchasedRows = getFilteredDetailedSalesTopRows(purchasedRows, detailedSalesTopPurchasedFilters);
  const filteredMissingRows = getFilteredDetailedSalesTopRows(missingRows, detailedSalesTopMissingFilters);

  if(!allGroupRows.length){
    return '';
  }

  const clientLabel = detailedSalesTopSummary
    ? [detailedSalesTopSummary.customerCode, detailedSalesTopSummary.customerName].filter(Boolean).join(' - ')
    : 'Wybrany klient';
  const totalPurchasedCount = detailedSalesTopSummary?.purchasedCount || 0;
  const totalOfferCount = detailedSalesTopSummary?.offerCount || 0;
  const totalMissingCount = typeof detailedSalesTopSummary?.missingCount === 'number'
    ? detailedSalesTopSummary.missingCount
    : Math.max(0, totalOfferCount - totalPurchasedCount);
  const totalCoveragePercent = getDetailedSalesTopCoveragePercent(
    totalPurchasedCount,
    totalOfferCount
  );
  const totalTone = getDetailedSalesTopCoverageTone(totalCoveragePercent);
  const highlightSourceRows = allGroupRows;
  const topOpportunityRow = [...highlightSourceRows].sort((a, b) => {
    if(b.missingCount !== a.missingCount) return b.missingCount - a.missingCount;
    return a.purchasedCount - b.purchasedCount;
  })[0] || null;
  const remainingGroupsCount = Math.max(0, allGroupRows.length - purchasedGroupRows.length);
  const purchasedProducerCount = detailedSalesTopProducerSummaryRows.filter(row => row.purchasedCount > 0).length;
  const missingOnlyProducerCount = detailedSalesTopProducerSummaryRows.filter(row => row.purchasedCount === 0 && row.missingCount > 0).length;
  const overviewRowsHtml = overviewGroupRows.map(row => {
    const coveragePercent = getDetailedSalesTopCoveragePercent(row.purchasedCount, row.offerCount);
    const tone = getDetailedSalesTopCoverageTone(coveragePercent);
    const isActive = row.groupKey === detailedSalesTopSelectedGroupKey;
    const activeClass = isActive ? ' is-active' : '';
    const isMissingOnly = row.purchasedCount === 0 && row.offerCount > 0;

    return `
      <button
        type="button"
        class="reports-overview-row is-${tone}${activeClass}"
        onclick="setDetailedSalesTopGroup(${escapeAttr(JSON.stringify(row.groupKey))})"
      >
        <span class="reports-overview-row-main">
          <span class="reports-overview-row-title">${escapeHtml(row.groupLabel)}</span>
          <span class="reports-overview-row-copy reports-overview-row-copy--stats">
            <span class="reports-mini-stat">
              <strong>${formatNumber(row.purchasedCount)}</strong>
              <em>kupione</em>
            </span>
            <span class="reports-mini-stat">
              <strong>${formatNumber(row.offerCount)}</strong>
              <em>w ofercie</em>
            </span>
            <span class="reports-mini-stat is-gap">
              <strong>${formatNumber(row.missingCount)}</strong>
              <em>braki</em>
            </span>
            ${isMissingOnly ? `<span class="reports-status-pill is-opportunity">Nie kupuje</span>` : ''}
          </span>
        </span>
        <span class="reports-overview-row-side">
          ${isActive ? `<span class="reports-active-pill">Wybrana kategoria</span>` : ''}
          <span class="reports-tone-badge is-${tone}">${escapeHtml(getDetailedSalesTopCoverageLabel(coveragePercent))}</span>
          <span class="reports-overview-row-percent">${escapeHtml(formatPercent(coveragePercent, { minimumFractionDigits: 1, maximumFractionDigits: 1 }))}</span>
          <span class="reports-progress-track">
            <span class="reports-progress-fill is-${tone}" style="width:${Math.max(0, Math.min(coveragePercent, 100))}%"></span>
          </span>
        </span>
      </button>
    `;
  }).join('');
  const producerRowsHtml = detailedSalesTopProducerSummaryRows.length
    ? detailedSalesTopProducerSummaryRows.map(row => {
      const producerCoveragePercent = getDetailedSalesTopCoveragePercent(row.purchasedCount, row.totalCount || (row.purchasedCount + row.missingCount));
      const tone = getDetailedSalesTopCoverageTone(producerCoveragePercent);
      const isActive = row.producerKey === detailedSalesTopSelectedProducerKey;
      const activeClass = isActive ? ' is-active' : '';
      const isMissingOnly = row.purchasedCount === 0 && row.missingCount > 0;
      return `
        <button
          type="button"
          class="reports-producer-row is-${tone}${activeClass}"
          onclick="setDetailedSalesTopProducer(${escapeAttr(JSON.stringify(row.producerKey))})"
        >
          <span class="reports-producer-row-main">
            <span class="reports-producer-row-title">${escapeHtml(row.producer)}</span>
            <span class="reports-producer-row-copy">
              <span class="reports-mini-stat">
                <strong>${formatNumber(row.purchasedCount)}</strong>
                <em>kupione</em>
              </span>
              <span class="reports-mini-stat is-gap">
                <strong>${formatNumber(row.missingCount)}</strong>
                <em>braki</em>
              </span>
              ${isMissingOnly ? `<span class="reports-status-pill is-opportunity">Nie kupuje</span>` : ''}
            </span>
          </span>
          <span class="reports-producer-row-side">
            ${isActive ? `<span class="reports-active-pill">Wybrany producent</span>` : ''}
            <span class="reports-producer-row-percent">${escapeHtml(formatPercent(producerCoveragePercent, { minimumFractionDigits: 1, maximumFractionDigits: 1 }))}</span>
          </span>
        </button>
      `;
    }).join('')
    : '';
  const producersIntro = selectedGroupRow
    ? `W tej kategorii w ofercie jest ${formatNumber(detailedSalesTopProducerSummaryRows.length)} producentów. Klient kupuje ${formatNumber(purchasedProducerCount)}, a ${formatNumber(missingOnlyProducerCount)} nie kupuje jeszcze wcale.`
    : 'Wybierz kategorię, aby zobaczyć producentów kupowanych przez klienta.';
  const selectedGroupCoveragePercent = selectedGroupRow
    ? getDetailedSalesTopCoveragePercent(selectedGroupRow.purchasedCount, selectedGroupRow.offerCount)
    : 0;
  const selectedGroupTone = getDetailedSalesTopCoverageTone(selectedGroupCoveragePercent);
  const selectedProducerCoveragePercent = selectedProducerRow
    ? getDetailedSalesTopCoveragePercent(selectedProducerRow.purchasedCount, selectedProducerRow.totalCount || (selectedProducerRow.purchasedCount + selectedProducerRow.missingCount))
    : 0;
  const selectedProducerTone = getDetailedSalesTopCoverageTone(selectedProducerCoveragePercent);

  return `
    <div class="reports-subsection reports-dashboard-section">
      <div class="reports-executive-hero is-${totalTone}">
        <div class="reports-executive-copy">
          <div class="reports-executive-kicker">Skrót klienta</div>
          <div class="reports-executive-title">${escapeHtml(clientLabel)}</div>
          <div class="reports-executive-text">
            Klient kupuje ${formatNumber(totalPurchasedCount)} z ${formatNumber(totalOfferCount)} indeksów Top. Do domknięcia pozostaje ${formatNumber(totalMissingCount)} pozycji.
          </div>
          <div class="reports-executive-progress">
            <span class="reports-tone-badge is-${totalTone}">${escapeHtml(getDetailedSalesTopCoverageLabel(totalCoveragePercent))}</span>
            <span class="reports-executive-progress-value">${escapeHtml(formatPercent(totalCoveragePercent, { minimumFractionDigits: 1, maximumFractionDigits: 1 }))}</span>
            <span class="reports-progress-track">
              <span class="reports-progress-fill is-${totalTone}" style="width:${Math.max(0, Math.min(totalCoveragePercent, 100))}%"></span>
            </span>
          </div>
        </div>
        <div class="reports-executive-metrics">
          <div class="reports-executive-metric">
            <span class="reports-executive-metric-label">Kupowane kategorie</span>
            <span class="reports-executive-metric-value">${formatNumber(purchasedGroupRows.length)} / ${formatNumber(allGroupRows.length)}</span>
          </div>
          <div class="reports-executive-metric">
            <span class="reports-executive-metric-label">Kupione</span>
            <span class="reports-executive-metric-value">${formatNumber(totalPurchasedCount)}</span>
          </div>
          <div class="reports-executive-metric">
            <span class="reports-executive-metric-label">Niekupowane grupy</span>
            <span class="reports-executive-metric-value">${formatNumber(nonPurchasedGroupRows.length)}</span>
          </div>
        </div>
      </div>

      <div class="reports-priority-panel is-${topOpportunityRow ? getDetailedSalesTopCoverageTone(getDetailedSalesTopCoveragePercent(topOpportunityRow.purchasedCount, topOpportunityRow.offerCount)) : 'low'}">
        <div class="reports-priority-panel-copy">
          <div class="reports-priority-panel-kicker">Największa szansa</div>
          <div class="reports-priority-panel-title">${escapeHtml(topOpportunityRow?.groupLabel || 'Brak największej szansy do wskazania')}</div>
          ${topOpportunityRow ? `
            <div class="reports-priority-panel-stats">
              <span class="reports-mini-stat">
                <strong>${formatNumber(topOpportunityRow.purchasedCount)}</strong>
                <em>kupione</em>
              </span>
              <span class="reports-mini-stat">
                <strong>${formatNumber(topOpportunityRow.offerCount)}</strong>
                <em>w ofercie</em>
              </span>
              <span class="reports-mini-stat is-gap">
                <strong>${formatNumber(topOpportunityRow.missingCount)}</strong>
                <em>braki</em>
              </span>
            </div>
          ` : ''}
          <div class="reports-priority-panel-text">
            ${topOpportunityRow
              ? 'To kategoria z największą liczbą brakujących indeksów w koszyku klienta.'
              : 'Brak danych do wskazania priorytetu.'
            }
          </div>
        </div>
        <div class="reports-priority-panel-side">
          <div class="reports-priority-panel-note">
            ${purchasedGroupRows.length
              ? `Klient kupuje ${formatNumber(purchasedGroupRows.length)} kategorii z rankingu. ${remainingGroupsCount ? `${formatNumber(remainingGroupsCount)} kategorii nadal nie kupuje.` : 'Klient ma aktywność we wszystkich kategoriach.'}`
              : 'Klient nie kupuje jeszcze żadnej kategorii z rankingu Top Rumunia.'
            }
          </div>
          ${topOpportunityRow ? `
            <button
              type="button"
              class="btn-outline"
              onclick="setDetailedSalesTopGroup(${escapeAttr(JSON.stringify(topOpportunityRow.groupKey))})"
            >
              Przejdź do tej kategorii
            </button>
          ` : ''}
        </div>
      </div>

      <div class="reports-analysis-dual-panels">
        <div class="reports-analysis-card reports-analysis-card--groups-panel">
          <div class="reports-analysis-card-head">
            <div>
              <div class="reports-analysis-card-title">Kategorie z oferty Top</div>
              <div class="reports-analysis-card-sub">Widzisz tu także grupy, których klient jeszcze nie kupuje. Kliknij kategorię, aby zejść do producentów i rekomendacji.</div>
            </div>
          </div>
          <div class="reports-overview-list reports-overview-list--panel">
            ${overviewRowsHtml}
          </div>
        </div>

        <div class="reports-analysis-card reports-analysis-card--producers-panel">
          <div class="reports-analysis-card-head reports-analysis-card-head--wide">
            <div>
              <div class="reports-analysis-card-title">${escapeHtml(selectedGroupRow ? `Producenci w kategorii: ${selectedGroupRow.groupLabel}` : 'Producenci w kategorii')}</div>
              <div class="reports-analysis-card-sub">${escapeHtml(producersIntro)}</div>
            </div>
            ${selectedGroupRow ? `
              <div class="reports-header-metrics">
                <div class="reports-header-metric">
                  <span class="reports-header-metric-label">W ofercie</span>
                  <span class="reports-header-metric-value">${formatNumber(detailedSalesTopProducerSummaryRows.length)}</span>
                </div>
                <div class="reports-header-metric">
                  <span class="reports-header-metric-label">Kupowani</span>
                  <span class="reports-header-metric-value">${formatNumber(purchasedProducerCount)}</span>
                </div>
                <div class="reports-header-metric">
                  <span class="reports-header-metric-label">Niekupowani</span>
                  <span class="reports-header-metric-value">${formatNumber(missingOnlyProducerCount)}</span>
                </div>
                <div class="reports-header-metric is-${selectedGroupTone}">
                  <span class="reports-header-metric-label">Udział</span>
                  <span class="reports-header-metric-value">${escapeHtml(formatPercent(selectedGroupCoveragePercent, { minimumFractionDigits: 1, maximumFractionDigits: 1 }))}</span>
                </div>
              </div>
            ` : ''}
          </div>
          <div class="reports-producer-list reports-producer-list--panel">
            ${producerRowsHtml || `<div class="reports-analysis-empty">Brak producentów do pokazania dla wybranej kategorii.</div>`}
          </div>
        </div>
      </div>

      <div class="reports-analysis-card reports-analysis-card--detail">
        ${selectedProducerRow ? `
          <div class="reports-analysis-card-head">
            <div>
              <div class="reports-analysis-card-title">${escapeHtml(selectedProducerRow.producer)}</div>
              <div class="reports-analysis-card-sub">${escapeHtml(selectedGroupRow?.groupLabel || '')} | kupione ${formatNumber(selectedProducerRow.purchasedCount)} | braki ${formatNumber(selectedProducerRow.missingCount)}</div>
            </div>
            <div class="reports-tone-badge is-${selectedProducerTone}">${escapeHtml(formatPercent(selectedProducerCoveragePercent, { minimumFractionDigits: 1, maximumFractionDigits: 1 }))}</div>
          </div>
          <div class="reports-summary reports-summary--compact">
            <div class="reports-summary-card">
              <div class="reports-summary-title">Grupa</div>
              <div class="reports-summary-meta">${escapeHtml(selectedGroupRow?.groupLabel || 'Wybierz grupę')}</div>
            </div>
            <div class="reports-summary-card">
              <div class="reports-summary-title">Producent</div>
              <div class="reports-summary-meta">${escapeHtml(selectedProducerRow.producer)}</div>
            </div>
            <div class="reports-summary-card">
              <div class="reports-summary-title">Kupione produkty</div>
              <div class="reports-summary-meta">${formatNumber(selectedProducerRow.purchasedCount)}</div>
            </div>
            <div class="reports-summary-card">
              <div class="reports-summary-title">Brakujące produkty</div>
              <div class="reports-summary-meta">${formatNumber(selectedProducerRow.missingCount)}</div>
            </div>
          </div>
          <div class="reports-two-columns">
            <div class="reports-column">
              <h3 class="reports-table-title">Kupione produkty</h3>
              ${renderDetailedSalesTopTableFilters('purchased', purchasedRows, detailedSalesTopPurchasedFilters)}
              ${renderDetailedSalesTopInfoToolbar(filteredPurchasedRows, {
                emptyText: 'Brak kupionych produktów dla wybranego producenta.',
                label: 'Widoczne'
              })}
              ${renderDetailedSalesTopRowsTable(filteredPurchasedRows, 'Brak kupionych produktów dla wybranego producenta.')}
            </div>
            <div class="reports-column">
              <h3 class="reports-table-title">Brakujące produkty</h3>
              ${renderDetailedSalesTopTableFilters('missing', missingRows, detailedSalesTopMissingFilters)}
              ${renderDetailedSalesTopQuickSelectionToolbar(filteredMissingRows, {
                scope: 'producer',
                selectLabel: 'Zaznacz produkty producenta',
                clearLabel: 'Odznacz produkty producenta'
              })}
              ${renderDetailedSalesTopRowsTable(filteredMissingRows, 'Brak brakujących produktów dla wybranego producenta.', { selectable: true })}
            </div>
          </div>
        ` : `
          <div class="reports-analysis-empty">
            Wybierz producenta z listy po prawej, aby zobaczyć kupione produkty oraz listę rekomendacji dla tego producenta.
          </div>
        `}
      </div>
    </div>
  `;
}

async function loadTopRumuniaOfferRows(){
  if(topRumuniaOfferCache) return topRumuniaOfferCache;
  const res = await fetch(jsonUrlTopRumunia);
  const data = await res.json();
  topRumuniaOfferCache = (data || []).map((row, rowIndex) => {
    const rankingRaw = getValueByVariants(row, ['ranking']);
    const rankingValue = Number(String(rankingRaw ?? '').replace(',', '.'));
    return {
      index: String(getValueByVariants(row, ['indeks', 'index', 'id', 'numer katalogowy', 'sku number']) || '').trim(),
      name: String(getValueByVariants(row, ['nazwa', 'name']) || '').trim(),
      producer: String(getValueByVariants(row, ['producent', 'producer', 'skrot producenta', 'skrót producenta', 'skrot_producenta', 'skrót_producenta', 'SKROT_PRODUCENTA']) || '').trim(),
      ean: String(getValueByVariants(row, ['kod ean', 'ean']) || '').trim(),
      group: String(getValueByVariants(row, ['grupa', 'group', 'nazwa grupa', 'nazwa_grupa', 'NAZWA_GRUPA']) || '').trim(),
      ranking: Number.isFinite(rankingValue) ? rankingValue : rowIndex + 1,
      rankingLabel: String(rankingRaw ?? (rowIndex + 1)).trim()
    };
  }).filter(row => row.index);
  return topRumuniaOfferCache;
}

function renderDetailedSalesTopComparisonContent(){
  const comparisonCustomers = getDetailedSalesCustomers();
  const filteredPurchasedRows = getFilteredDetailedSalesTopRows(detailedSalesTopPurchasedRows, detailedSalesTopPurchasedFilters);
  const filteredMissingRows = getFilteredDetailedSalesTopRows(detailedSalesTopMissingRows, detailedSalesTopMissingFilters);
  const producerComparisonHtml = renderDetailedSalesTopProducerComparisonSection();
  const fullComparisonHtml = renderDetailedSalesTopFullComparisonSection(filteredPurchasedRows, filteredMissingRows);
  const customerPortfolioHtml = renderDetailedSalesTopCustomerPortfolioSection();
  const sectionSwitchHtml = renderDetailedSalesTopSectionSwitch();
  const exportToolbarHtml = renderDetailedSalesTopExportToolbar();
  const hasComparisonCustomer = comparisonCustomers.some(customer => {
    return `${customer.code}|||${customer.name}` === detailedSalesComparisonCustomer;
  });
  const comparisonFilterHtml = `
    <div class="reports-filters">
      <label class="reports-filter-card">
        <span class="reports-limit-title">Klient do porównania</span>
        <input
          type="text"
          value="${escapeAttr(detailedSalesComparisonCustomerSearch)}"
          oninput="setDetailedSalesComparisonCustomerSearch(this.value)"
          placeholder="Szukaj po numerze lub nazwie, np. 338"
          data-detailed-sales-comparison-search
          ${comparisonCustomers.length ? '' : 'disabled'}
        >
        <select onchange="setDetailedSalesComparisonCustomer(this.value)" data-detailed-sales-comparison-select ${comparisonCustomers.length ? '' : 'disabled'}>
          ${buildDetailedSalesComparisonCustomerOptionsHtml(comparisonCustomers)}
        </select>
      </label>
    </div>
  `;

  if(!hasComparisonCustomer){
    return wrapDetailedSalesTopCompactContent(`
      <div class="reports-toolbar">
        <div>
          <div class="import-title">Porównanie z Top Rumunia</div>
          <div class="import-sub">Wybierz klienta w filtrze poniżej, aby porównać jego zakupy z rankingiem Top Rumunia.</div>
        </div>
      </div>
      ${comparisonFilterHtml}
      ${customerPortfolioHtml}
    `);
  }

  if(detailedSalesTopComparisonLoading && !detailedSalesTopPurchasedRows.length && !detailedSalesTopMissingRows.length){
    return wrapDetailedSalesTopCompactContent(`
      <div class="reports-toolbar">
        <div>
          <div class="import-title">Porównanie z Top Rumunia</div>
          <div class="import-sub">Analizuję kupione i brakujące indeksy klienta względem rankingu Top Rumunia.</div>
        </div>
        <div class="import-info">Status: ładowanie porównania</div>
      </div>
      ${comparisonFilterHtml}
      ${customerPortfolioHtml}
    `);
  }

  const summary = detailedSalesTopSummary
    ? `
      <div class="reports-summary">
        <div class="reports-summary-card">
          <div class="reports-summary-title">Klient</div>
          <div class="reports-summary-meta">${escapeHtml([detailedSalesTopSummary.customerCode, detailedSalesTopSummary.customerName].filter(Boolean).join(' - ') || 'Wybrany klient')}</div>
        </div>
        <div class="reports-summary-card">
          <div class="reports-summary-title">Zakres grupy</div>
          <div class="reports-summary-meta">${escapeHtml(detailedSalesTopSummary.groupLabel)}</div>
        </div>
        <div class="reports-summary-card">
          <div class="reports-summary-title">Kupione z Top</div>
          <div class="reports-summary-meta">${detailedSalesTopSummary.purchasedCount}</div>
        </div>
        <div class="reports-summary-card">
          <div class="reports-summary-title">Brakujące z Top</div>
          <div class="reports-summary-meta">${detailedSalesTopSummary.missingCount} z ${detailedSalesTopSummary.offerCount}</div>
        </div>
      </div>
    `
    : '';

  const topComparisonStatusHtml = detailedSalesTopComparisonLoading
    ? '<div class="import-info">Status: odświeżam porównanie</div>'
    : '';
  const infoHtml = detailedSalesTopComparisonError
    ? `<div class="reports-toolbar"><div><div class="import-title">Porównanie z Top Rumunia</div><div class="import-sub">${escapeHtml(detailedSalesTopComparisonError)}</div></div></div>`
    : `
      <div class="reports-toolbar">
        <div>
          <div class="import-title">Porównanie z Top Rumunia</div>
          <div class="import-sub">Najpierw sprawdź skrót klienta, a potem przełącz się na pełną listę Top, jeśli chcesz zobaczyć wszystkie kupione i brakujące indeksy.</div>
        </div>
        ${topComparisonStatusHtml}
      </div>
    `;

  return wrapDetailedSalesTopCompactContent(`
    ${infoHtml}
    ${comparisonFilterHtml}
    ${customerPortfolioHtml}
    ${detailedSalesTopSection === 'full' ? summary : ''}
    ${sectionSwitchHtml}
    ${exportToolbarHtml}
    ${detailedSalesTopSection === 'full' ? fullComparisonHtml : producerComparisonHtml}
  `);
}

function getDetailedSalesFileInfoHtml(){
  if(detailedSalesImportFile){
    return `Zakres: ${escapeHtml(formatWeeksCountLabel(detailedSalesLoadedReportsCount || detailedSalesWeeksLimit))}${detailedSalesAvailableReportsCount ? ` | dostępne: ${escapeHtml(formatWeeksCountLabel(detailedSalesAvailableReportsCount))}` : ''}<br>Ostatni plik: ${escapeHtml(detailedSalesImportFile)}<br>Aktualizacja: ${escapeHtml(formatInsightDate(detailedSalesUpdatedAt))}${detailedSalesLoading ? '<br>Status: odświeżam zakres raportu...' : ''}`;
  }
  return `Status: ${escapeHtml(detailedSalesErrorMessage || 'brak przypisanego pliku')}`;
}

function renderDetailedSalesTopComparisonStandaloneContent(){
  const hasData = detailedSalesSourceRows.length > 0;
  const groups = getDetailedSalesGroups();
  const comparisonCustomers = getDetailedSalesCustomers();
  const filteredPurchasedRows = getFilteredDetailedSalesTopRows(detailedSalesTopPurchasedRows, detailedSalesTopPurchasedFilters);
  const filteredMissingRows = getFilteredDetailedSalesTopRows(detailedSalesTopMissingRows, detailedSalesTopMissingFilters);
  const producerComparisonHtml = renderDetailedSalesTopProducerComparisonSection();
  const fullComparisonHtml = renderDetailedSalesTopFullComparisonSection(filteredPurchasedRows, filteredMissingRows);
  const customerPortfolioHtml = renderDetailedSalesTopCustomerPortfolioSection();
  const sectionSwitchHtml = renderDetailedSalesTopSectionSwitch();
  const exportToolbarHtml = renderDetailedSalesTopExportToolbar();
  const hasComparisonCustomer = comparisonCustomers.some(customer => {
    return `${customer.code}|||${customer.name}` === detailedSalesComparisonCustomer;
  });
  if(detailedSalesLoading && !hasData){
    return wrapDetailedSalesTopCompactContent(`
      <div class="reports-toolbar">
        <div>
          <div class="import-title">Potencjał klienta Top Rumunia</div>
          <div class="import-sub">Pobieram raport per indeks przypisany do Twojego konta z Firebase.</div>
        </div>
        <div class="import-info">Status: ładowanie raportu</div>
      </div>
    `);
  }

  const filtersHtml = `
    <div class="reports-filters">
      <label class="reports-filter-card">
        <span class="reports-limit-title">Ile tygodni pokazać</span>
        <input
          type="number"
          min="1"
          step="1"
          value="${escapeAttr(detailedSalesWeeksLimit)}"
          onchange="setDetailedSalesWeeksLimit(this.value)"
          ${detailedSalesLoading ? 'disabled' : ''}
        >
      </label>
      ${renderDetailedSalesGroupsFilter(groups, { disabled: !hasData })}
      <label class="reports-filter-card">
        <span class="reports-limit-title">Klient do porównania</span>
        <input
          type="text"
          value="${escapeAttr(detailedSalesComparisonCustomerSearch)}"
          oninput="setDetailedSalesComparisonCustomerSearch(this.value)"
          placeholder="Szukaj po numerze lub nazwie, np. 338"
          data-detailed-sales-comparison-search
          ${comparisonCustomers.length ? '' : 'disabled'}
        >
        <select onchange="setDetailedSalesComparisonCustomer(this.value)" data-detailed-sales-comparison-select ${comparisonCustomers.length ? '' : 'disabled'}>
          ${buildDetailedSalesComparisonCustomerOptionsHtml(comparisonCustomers)}
        </select>
      </label>
    </div>
  `;

  if(!hasComparisonCustomer){
    const emptyMessage = detailedSalesErrorMessage
      || (hasData ? 'Wybierz klienta, aby porównać jego zakupy z rankingiem Top Rumunia.' : 'Brak szczegółowego raportu sprzedaży dla Twojego konta.');

    return wrapDetailedSalesTopCompactContent(`
      <div class="reports-toolbar">
        <div>
          <div class="import-title">Potencjał klienta Top Rumunia</div>
          <div class="import-sub">Raport na wzór widoku admina, oparty o dane per indeks przypisane do Twojego konta.</div>
        </div>
      </div>
      ${filtersHtml}
      ${customerPortfolioHtml}
      <div class="reports-toolbar">
        <div>
          <div class="import-title">Porównanie klienta z Top Rumunia</div>
          <div class="import-sub">${escapeHtml(emptyMessage)}</div>
        </div>
      </div>
    `);
  }

  const summary = detailedSalesTopSummary
    ? `
      <div class="reports-summary">
        <div class="reports-summary-card">
          <div class="reports-summary-title">Klient</div>
          <div class="reports-summary-meta">${escapeHtml([detailedSalesTopSummary.customerCode, detailedSalesTopSummary.customerName].filter(Boolean).join(' - ') || 'Wybrany klient')}</div>
        </div>
        <div class="reports-summary-card">
          <div class="reports-summary-title">Zakres grupy</div>
          <div class="reports-summary-meta">${escapeHtml(detailedSalesTopSummary.groupLabel)}</div>
        </div>
        <div class="reports-summary-card">
          <div class="reports-summary-title">Kupione z Top</div>
          <div class="reports-summary-meta">${detailedSalesTopSummary.purchasedCount}</div>
        </div>
        <div class="reports-summary-card">
          <div class="reports-summary-title">Brakujące z Top</div>
          <div class="reports-summary-meta">${detailedSalesTopSummary.missingCount} z ${detailedSalesTopSummary.offerCount}</div>
        </div>
      </div>
    `
    : '';

  const topComparisonStatusHtml = detailedSalesTopComparisonLoading
    ? '<div class="import-info">Status: odświeżam porównanie</div>'
    : '';
  const infoHtml = detailedSalesTopComparisonError
    ? `<div class="reports-toolbar"><div><div class="import-title">Porównanie klienta z Top Rumunia</div><div class="import-sub">${escapeHtml(detailedSalesTopComparisonError)}</div></div></div>`
    : `
      <div class="reports-toolbar">
        <div>
          <div class="import-title">Potencjał klienta Top Rumunia</div>
          <div class="import-sub">Skrót klienta pokaże kategorie i producentów z największą szansą na wzrost. Pełna lista pokaże wszystkie indeksy z Top Rumunia.</div>
        </div>
        ${topComparisonStatusHtml}
      </div>
    `;

  return wrapDetailedSalesTopCompactContent(`
    ${infoHtml}
    ${filtersHtml}
    ${customerPortfolioHtml}
    ${detailedSalesTopSection === 'full' ? summary : ''}
    ${sectionSwitchHtml}
    ${exportToolbarHtml}
    ${detailedSalesTopSection === 'full' ? fullComparisonHtml : producerComparisonHtml}
  `);
}

function renderReportsSubnav(){
  if(!reportsSubnav) return;
  const areButtonsDisabled = phReportsAutoLoading || detailedSalesLoading;
  const tabs = isCurrentUserAdmin()
    ? [
      { mode: 'admin-weekly-sales', label: 'Sprzedaż per tydzień' },
      { mode: 'top-sales', label: 'Top sprzedaż' },
      { mode: 'client-gap', label: 'Potencjał klienta' },
      { mode: 'top-suggestions', label: 'Propozycje Top' }
    ]
    : [
      { mode: 'weekly-sales', label: 'Sprzedaż tygodniowa' },
      { mode: 'detailed-sales', label: 'Sprzedaż szczegółowa' },
      { mode: 'top-rumunia-comparison', label: 'Potencjał klienta' }
    ];

  reportsSubnav.innerHTML = tabs.map(tab => `
    <button
      type="button"
      class="tab reports-nav-tab ${reportsMode === tab.mode ? 'active' : ''}"
      onclick="setReportsMode('${tab.mode}')"
      ${areButtonsDisabled ? 'disabled' : ''}
    >
      ${escapeHtml(tab.label)}
    </button>
  `).join('');
}

function isReportsModeAllowed(mode, isAdminReports = isCurrentUserAdmin()){
  const allowedModes = isAdminReports
    ? ['admin-weekly-sales', 'top-sales', 'client-gap', 'top-suggestions']
    : ['weekly-sales', 'detailed-sales', 'top-rumunia-comparison'];
  return allowedModes.includes(mode);
}

async function ensurePhReportModeDataLoaded(mode){
  if(isCurrentUserAdmin()) return;
  if(mode === 'detailed-sales' || mode === 'top-rumunia-comparison'){
    await ensurePhDetailedSalesDataLoaded();
    return;
  }
  await ensurePhWeeklySalesDataLoaded();
}

async function prefetchPhReportsDataInBackground(){
  if(isCurrentUserAdmin()) return;
  if(reportsBackgroundPrefetchStarted) return;
  reportsBackgroundPrefetchStarted = true;

  try{
    await ensurePhWeeklySalesDataLoaded();
  }catch(error){
    console.error('Background weekly sales prefetch error', error);
  }

  try{
    await ensurePhDetailedSalesDataLoaded();
  }catch(error){
    console.error('Background detailed sales prefetch error', error);
  }
}

function renderReportsView(){
  if(!reportsContainer) return;
  const isAdminReports = isCurrentUserAdmin();
  const contentByMode = {
    'admin-weekly-sales': typeof renderAdminWeeklySalesContent === 'function' ? renderAdminWeeklySalesContent : null,
    'top-sales': typeof renderTopSalesReportContent === 'function' ? renderTopSalesReportContent : null,
    'client-gap': typeof renderClientComparisonContent === 'function' ? renderClientComparisonContent : null,
    'weekly-sales': typeof renderWeeklySalesContent === 'function' ? renderWeeklySalesContent : null,
    'detailed-sales': typeof renderDetailedSalesContent === 'function' ? renderDetailedSalesContent : null,
    'top-rumunia-comparison': typeof renderDetailedSalesTopComparisonStandaloneContent === 'function' ? renderDetailedSalesTopComparisonStandaloneContent : null,
    'top-suggestions': typeof renderTopSuggestionsContent === 'function' ? renderTopSuggestionsContent : null
  };
  const fallbackMode = isAdminReports ? 'top-sales' : 'weekly-sales';
  if(!isReportsModeAllowed(reportsMode, isAdminReports)){
    reportsMode = fallbackMode;
  }
  const renderContent = contentByMode[reportsMode]
    || contentByMode[fallbackMode]
    || (() => `
      <div class="reports-toolbar">
        <div>
          <div class="import-title">Raport chwilowo niedostępny</div>
          <div class="import-sub">Brakuje podpiętego renderera dla tego widoku po ostatnim cofnięciu zmian.</div>
        </div>
      </div>
    `);

  renderReportsSubnav();

  reportsContainer.innerHTML = `
    <div class="reports-panel">
      ${renderContent()}
    </div>
  `;
  registerLazyImages(reportsContainer);
  syncDetailedSalesComparisonCustomerControls();
  document.dispatchEvent(new CustomEvent('wf:reports-rendered', {
    detail: {
      mode: reportsMode
    }
  }));
  document.dispatchEvent(new CustomEvent('wf:quick-order-refresh'));
}

function openReportsView(){
  const isAdminReports = isCurrentUserAdmin();
  document.querySelectorAll('.tab').forEach(tab => tab.classList.remove('active'));
  setPrimaryNavState('reports');
  showTabSection('reports-sales');
  setActiveCard('reports-card');
  setImageBase(imageBaseUrlRumunia);
  setFullData([]);
  isLoading = false;
  activeContainer = reportsContainer;
  currentCategoryName = 'Raporty';
  currentCategorySlug = 'raporty';
  resetFilters();
  clearAllContentContainers();
  if(!isAdminReports){
    if(!isReportsModeAllowed(reportsMode, false)){
      reportsMode = 'weekly-sales';
    }
    weeklySalesRepComparison = false;
  }
  renderReportsView();
  if(!isAdminReports){
    void ensurePhReportModeDataLoaded(reportsMode);
    setTimeout(() => {
      void prefetchPhReportsDataInBackground();
    }, 0);
  }else if(typeof ensureAdminReportModeDataLoaded === 'function'){
    void ensureAdminReportModeDataLoaded(reportsMode);
  }
}

async function importReportsExcel(){
  const input = document.getElementById('reports-excel-input');
  const file = input?.files?.[0];
  if(!file) return;

  try{
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    const headers = rows.length ? Object.keys(rows[0]) : [];
    reportsSourceRows = rows
      .map(row => normalizeReportRow(row, headers))
      .filter(row => row.index && row.productGroup);
    reportsImportFile = file.name;
    generateReportsData();
  }catch(e){
    console.error('Reports import error', e);
    reportsSourceRows = [];
    reportsGeneratedRows = [];
    reportsImportFile = '';
    reportsSummary = [];
  }

  renderReportsView();
}

function setReportGroupLimit(groupId, value){
  const parsed = Number(value);
  reportsGroupLimits[groupId] = Number.isFinite(parsed) && parsed >= 0 ? Math.floor(parsed) : 0;
  if(reportsMode === 'top-sales' && reportsSourceRows.length){
    generateReportsData();
  }
  if(reportsMode === 'client-gap' && clientReportSourceRows.length){
    generateClientComparisonReport().then(renderReportsView);
    return;
  }
  renderReportsView();
}

function setDetailedSalesTopTableFilter(type, key, value){
  if(type === 'purchased'){
    detailedSalesTopPurchasedFilters[key] = value;
  }else{
    detailedSalesTopMissingFilters[key] = value;
  }
  renderReportsView();
}

function setDetailedSalesTopSection(section){
  detailedSalesTopSection = section === 'full' ? 'full' : 'summary';
  renderReportsView();
}

function toggleDetailedSalesTopCustomerPortfolioExpanded(){
  detailedSalesTopCustomerPortfolioExpanded = !detailedSalesTopCustomerPortfolioExpanded;
  renderReportsView();
}

function setDetailedSalesTopCustomerSearch(value){
  detailedSalesTopCustomerSearch = String(value || '');
  renderReportsView();
}

function setDetailedSalesTopCustomerSort(value){
  const allowedValues = new Set([
    'purchased-desc',
    'purchased-asc',
    'coverage-desc',
    'coverage-asc',
    'groups-desc',
    'sales-desc',
    'missing-desc'
  ]);
  detailedSalesTopCustomerSort = allowedValues.has(value)
    ? value
    : 'purchased-desc';
  renderReportsView();
}

function setDetailedSalesTopCustomerToneFilter(value){
  detailedSalesTopCustomerToneFilter = ['high', 'medium', 'low'].includes(value)
    ? value
    : '';
  renderReportsView();
}

function setDetailedSalesTopGroup(groupKey){
  detailedSalesTopSelectedGroupKey = getDetailedSalesTopGroupKey(groupKey);
  detailedSalesTopProducerSummaryRows = buildDetailedSalesTopProducerSummaryRows(
    detailedSalesTopPurchasedRows,
    detailedSalesTopMissingRows,
    detailedSalesTopSelectedGroupKey
  );
  detailedSalesTopSelectedProducerKey = detailedSalesTopProducerSummaryRows[0]?.producerKey || '';
  renderReportsView();
}

function setDetailedSalesTopProducer(producerKey){
  detailedSalesTopSelectedProducerKey = String(producerKey || '').trim();
  renderReportsView();
}

async function setDetailedSalesComparisonCustomer(value, options = {}){
  const preserveSearch = Boolean(options.preserveSearch);
  if(String(detailedSalesComparisonCustomer || '') === String(value || '')){
    return;
  }
  detailedSalesComparisonCustomer = value;
  if(value && !preserveSearch){
    const [customerCode = '', customerName = ''] = String(value).split('|||');
    detailedSalesComparisonCustomerSearch = [customerCode, customerName].filter(Boolean).join(' - ');
  }
  clearDetailedSalesTopComparisonResultState({ preserveCustomerSummaryRows: true });
  renderReportsView();
  await generateDetailedSalesTopComparison();
}

function setDetailedSalesComparisonCustomerSearch(value){
  detailedSalesComparisonCustomerSearch = String(value || '');
  syncDetailedSalesComparisonCustomerControls();
  const customers = typeof getDetailedSalesCustomers === 'function' ? getDetailedSalesCustomers() : [];
  const autoMatchValue = resolveDetailedSalesComparisonCustomerAutoMatch(customers, detailedSalesComparisonCustomerSearch);
  if(autoMatchValue && autoMatchValue !== detailedSalesComparisonCustomer){
    void setDetailedSalesComparisonCustomer(autoMatchValue, { preserveSearch: true });
  }
}

async function setReportsMode(mode){
  const isAdminReports = isCurrentUserAdmin();
  reportsMode = isReportsModeAllowed(mode, isAdminReports)
    ? mode
    : (isAdminReports ? 'top-sales' : 'weekly-sales');
  if(reportsMode === 'top-rumunia-comparison'){
    detailedSalesTopSection = 'summary';
  }
  renderReportsView();
  if(!isAdminReports){
    void ensurePhReportModeDataLoaded(reportsMode);
  }else if(typeof ensureAdminReportModeDataLoaded === 'function'){
    void ensureAdminReportModeDataLoaded(reportsMode);
  }
}

function setDetailedSalesTopMissingSelectionByKey(rowKey, checked){
  const normalizedKey = String(rowKey || '').trim();
  if(!normalizedKey) return;
  if(checked){
    detailedSalesTopSelectedMissingKeys.add(normalizedKey);
  }else{
    detailedSalesTopSelectedMissingKeys.delete(normalizedKey);
  }
  renderReportsView();
}

function toggleDetailedSalesTopMissingSelectionByKey(rowKey){
  const normalizedKey = String(rowKey || '').trim();
  if(!normalizedKey) return;
  if(detailedSalesTopSelectedMissingKeys.has(normalizedKey)){
    detailedSalesTopSelectedMissingKeys.delete(normalizedKey);
  }else{
    detailedSalesTopSelectedMissingKeys.add(normalizedKey);
  }
  renderReportsView();
}

function selectVisibleDetailedSalesTopMissing(scope = 'filtered'){
  const rows = getDetailedSalesTopVisibleMissingRows(scope);
  rows.forEach(row => {
    const rowKey = getDetailedSalesTopMissingRowKey(row);
    if(rowKey){
      detailedSalesTopSelectedMissingKeys.add(rowKey);
    }
  });
  renderReportsView();
}

function clearVisibleDetailedSalesTopMissing(scope = 'filtered'){
  const visibleKeys = new Set(
    getDetailedSalesTopVisibleMissingRows(scope).map(getDetailedSalesTopMissingRowKey).filter(Boolean)
  );
  if(!visibleKeys.size) return;
  detailedSalesTopSelectedMissingKeys = new Set(
    Array.from(detailedSalesTopSelectedMissingKeys).filter(key => !visibleKeys.has(key))
  );
  renderReportsView();
}

function clearAllDetailedSalesTopMissing(){
  if(!detailedSalesTopSelectedMissingKeys.size) return;
  detailedSalesTopSelectedMissingKeys = new Set();
  renderReportsView();
}

function exportDetailedSalesTopMissingCsv(){
  const exportRows = getDetailedSalesTopMissingExportRows();
  if(!exportRows.length) return;
  const cols = ['INDEKS', 'NAZWA', 'SKROT_PRODUCENTA', 'Ranking', 'Grupa produktowa', 'Zdjęcie URL'];
  let csv = '\uFEFF' + cols.join(',') + '\n';
  exportRows.forEach(row => {
    csv += cols
      .map(col => `"${String(row[col] ?? '').replace(/"/g, '""')}"`)
      .join(',') + '\n';
  });
  download(csv, `${getDetailedSalesTopRecommendationFileName()}.csv`, 'text/csv;charset=utf-8;');
}

function exportDetailedSalesTopMissingXls(){
  const exportRows = getDetailedSalesTopMissingExportRows();
  if(!exportRows.length) return;
  const cols = ['INDEKS', 'NAZWA', 'SKROT_PRODUCENTA', 'Ranking', 'Grupa produktowa', 'Zdjęcie URL'];
  const rows = [cols, ...exportRows.map(row => cols.map(col => row[col] ?? ''))];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'rekomendacje');
  try{
    const out = XLSX.write(wb, { bookType: 'xls', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.ms-excel' });
    downloadBlob(blob, `${getDetailedSalesTopRecommendationFileName()}.xls`);
  }catch(error){
    console.error('Detailed sales Top XLS export error', error);
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    downloadBlob(blob, `${getDetailedSalesTopRecommendationFileName()}.xlsx`);
  }
}

async function createDetailedSalesTopRecommendationCatalog(){
  const selectedRows = getSelectedDetailedSalesTopMissingRows();
  if(!selectedRows.length) return;
  selectedProducts = new Map();
  selectedRows.forEach(row => {
    const idx = normalizeIndexValue(row.INDEKS);
    if(idx){
      selectedProducts.set(idx, {
        INDEKS: row.INDEKS,
        NAZWA: row.NAZWA,
        PRODUCENT: row.SKROT_PRODUCENTA,
        GRUPA: row['Grupa produktowa'],
        'KOD EAN': row['KOD EAN'] || ''
      });
    }
  });
  setFullData(Array.from(selectedProducts.values()), ['INDEKS', 'NAZWA', 'PRODUCENT', 'GRUPA', 'KOD EAN']);
  activeContainer = reportsContainer;
  currentImageBaseUrl = imageBaseUrlRumunia;
  window.imageBaseUrl = currentImageBaseUrl;
  currentCategorySlug = getDetailedSalesTopRecommendationFileName();
  catalogOptionsOverride = {
    groupByField: 'GRUPA',
    sectionTitle: 'Grupa produktowa'
  };
  await createCatalog();
}

function normalizePdfReportText(value){
  return String(value ?? '')
    .replace(/[ąćęłńóśźżĄĆĘŁŃÓŚŹŻ]/g, char => ({
      'ą': 'a',
      'ć': 'c',
      'ę': 'e',
      'ł': 'l',
      'ń': 'n',
      'ó': 'o',
      'ś': 's',
      'ź': 'z',
      'ż': 'z',
      'Ą': 'A',
      'Ć': 'C',
      'Ę': 'E',
      'Ł': 'L',
      'Ń': 'N',
      'Ó': 'O',
      'Ś': 'S',
      'Ź': 'Z',
      'Ż': 'Z'
    }[char] || char))
    .replace(/[–—]/g, '-')
    .replace(/[„”]/g, '"')
    .replace(/[’]/g, '\'')
    .replace(/[\u0000-\u001f\u007f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function getDetailedSalesRepresentativeSummaryFileName(representativeLabel){
  const safePart = (value, fallback) => String(value || fallback)
    .trim()
    .replace(/[\\/:*?"<>|]+/g, ' ')
    .replace(/\s+/g, ' ');
  return `${safePart(representativeLabel, 'PH')} - raport PH Top Rumunia`;
}

async function exportDetailedSalesRepresentativeSummaryPdf(){
  if(!window.jspdf || !detailedSalesSourceRows.length) return;

  const [topOfferRows, representativeConfig, logoImage] = await Promise.all([
    loadTopRumuniaOfferRows(),
    getPersonalSalesConfig(auth?.currentUser?.uid).catch(() => null),
    loadImageAsDataUrl(maspoLogoUrl).catch(() => null)
  ]);
  const scopedTopRows = dedupeReportRows(
    (topOfferRows || []).filter(row => matchesDetailedSalesTopGroup(row.group, detailedSalesSelectedGroups))
  ).sort((a, b) => {
    if(a.ranking !== b.ranking) return a.ranking - b.ranking;
    return String(a.name || '').localeCompare(String(b.name || ''), 'pl');
  });
  const customerRows = buildDetailedSalesTopCustomerSummaryRows(scopedTopRows);
  if(!customerRows.length) return;

  const meetingRows = buildDetailedSalesTopMeetingRecommendationRows(scopedTopRows);
  const weeklySummary = buildWeeklySalesRepresentativePdfSummary();
  const representativeLabel = String(
    representativeConfig?.displayName
    || weeklySummary.representativeLabel
    || weeklySalesSelectedRepresentative
    || 'Przedstawiciel handlowy'
  ).trim();
  const generatedAtLabel = new Intl.DateTimeFormat('pl-PL', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  }).format(new Date());
  const topScopeLabel = `${formatWeeksCountLabel(detailedSalesLoadedReportsCount || detailedSalesWeeksLimit)} | ${getDetailedSalesSelectedGroupsLabel(3)}`;
  const strongestRow = [...customerRows].sort((a, b) => b.coveragePercent - a.coveragePercent || b.purchasedCount - a.purchasedCount)[0] || null;
  const biggestSalesRow = [...customerRows].sort((a, b) => b.salesValue - a.salesValue || b.purchasedCount - a.purchasedCount)[0] || null;
  const potentialRow = [...customerRows].sort((a, b) => b.missingCount - a.missingCount || b.salesValue - a.salesValue)[0] || null;
  const sortedCustomerRows = [...customerRows].sort((a, b) => {
    if(b.salesValue !== a.salesValue) return b.salesValue - a.salesValue;
    if(b.missingCount !== a.missingCount) return b.missingCount - a.missingCount;
    return String(a.customerName || '').localeCompare(String(b.customerName || ''), 'pl');
  });
  const totalScopedSales = customerRows.reduce((sum, row) => sum + Number(row.salesValue || 0), 0);
  const averageCoverage = customerRows.reduce((sum, row) => sum + Number(row.coveragePercent || 0), 0) / (customerRows.length || 1);
  const activeTopCustomers = customerRows.filter(row => row.purchasedCount > 0).length;
  const noTopCustomers = customerRows.filter(row => row.purchasedCount === 0).length;
  const topMeetingRows = meetingRows.slice(0, 3);

  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({ orientation: 'p', unit: 'pt', format: 'a4' });
  const pageW = pdf.internal.pageSize.getWidth();
  const pageH = pdf.internal.pageSize.getHeight();
  const margin = 34;
  const contentW = pageW - (margin * 2);
  const footerY = pageH - 18;
  let y = margin;

  const safeText = value => normalizePdfReportText(value);
  const metricCardW = (contentW - 12) / 2;
  const metricCardH = 58;

  const drawWrappedText = (text, x, startY, width, lineHeight = 12) => {
    const lines = pdf.splitTextToSize(safeText(text), width);
    pdf.text(lines, x, startY);
    return startY + (lines.length * lineHeight);
  };

  const startPage = () => {
    if(pdf.getNumberOfPages() > 0 && y !== margin){
      pdf.addPage();
    }
    y = margin;
    if(logoImage?.dataUrl){
      pdf.addImage(logoImage.dataUrl, logoImage.format, margin, y, 92, 28);
    }
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(20);
    pdf.setTextColor(15, 27, 45);
    pdf.text(safeText('Raport PH - Top Rumunia'), margin + 104, y + 18);
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(10);
    pdf.setTextColor(90);
    pdf.text(safeText(`${representativeLabel} | ${generatedAtLabel}`), margin + 104, y + 36);
    pdf.setDrawColor(214, 222, 233);
    pdf.setLineWidth(1);
    pdf.line(margin, y + 50, pageW - margin, y + 50);
    pdf.setTextColor(0);
    y += 70;
  };

  const ensureSpace = height => {
    if(y + height > pageH - 30){
      startPage();
    }
  };

  const drawMetricCard = (x, top, label, value, note = '') => {
    pdf.setDrawColor(207, 221, 237);
    pdf.setFillColor(248, 251, 255);
    pdf.roundedRect(x, top, metricCardW, metricCardH, 10, 10, 'FD');
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(9);
    pdf.setTextColor(88, 104, 129);
    pdf.text(safeText(label), x + 12, top + 16);
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(15);
    pdf.setTextColor(15, 27, 45);
    pdf.text(safeText(value), x + 12, top + 35);
    if(note){
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(8);
      pdf.setTextColor(95, 111, 137);
      pdf.text(safeText(note), x + 12, top + 49);
    }
    pdf.setTextColor(0);
  };

  const drawSectionTitle = (title, subtitle = '') => {
    ensureSpace(subtitle ? 36 : 24);
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(14);
    pdf.setTextColor(15, 27, 45);
    pdf.text(safeText(title), margin, y);
    y += 14;
    if(subtitle){
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(9);
      pdf.setTextColor(95, 111, 137);
      y = drawWrappedText(subtitle, margin, y, contentW, 11);
    }
    pdf.setTextColor(0);
    y += 6;
  };

  const drawBulletLines = items => {
    items.filter(Boolean).forEach(item => {
      const bulletLines = pdf.splitTextToSize(safeText(item), contentW - 12);
      ensureSpace((bulletLines.length * 12) + 6);
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(10);
      pdf.text('-', margin, y);
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(10);
      pdf.text(bulletLines, margin + 12, y);
      y += (bulletLines.length * 12) + 2;
    });
  };

  startPage();

  drawSectionTitle('Szybki obraz PH', `Zakres Top: ${topScopeLabel}${weeklySummary.activeWeekLabel ? ` | Sprzedaz tygodniowa: ${weeklySummary.activeWeekLabel}` : ''}`);
  drawMetricCard(margin, y, 'Klienci w raporcie', `${formatNumber(customerRows.length)}`, `${formatNumber(activeTopCustomers)} kupuje cos z Top`);
  drawMetricCard(margin + metricCardW + 12, y, 'Suma sprzedazy w zakresie', `${formatNumber(totalScopedSales)}`, 'Sprzedaz z aktualnego zakresu grup');
  y += metricCardH + 10;
  drawMetricCard(margin, y, 'Sredni udzial Top', `${formatPercent(averageCoverage, { minimumFractionDigits: 1, maximumFractionDigits: 1 })}`, `Klienci bez zakupow z Top: ${formatNumber(noTopCustomers)}`);
  drawMetricCard(
    margin + metricCardW + 12,
    y,
    'Sprzedaz tygodniowa',
    `${formatNumber(weeklySummary.activeSalesTotal)}`,
    weeklySummary.previousWeekLabel
      ? `${weeklySummary.previousWeekLabel}: ${formatNumber(weeklySummary.previousSalesTotal)} | trend ${formatNumber(weeklySummary.trendValue)}`
      : 'Brak poprzedniego tygodnia do porownania'
  );
  y += metricCardH + 18;

  drawSectionTitle('Najwazniejsze informacje dla PH');
  drawBulletLines([
    strongestRow ? `Najwyzszy udzial Top ma klient ${getDetailedSalesTopCustomerLabel(strongestRow)} - ${formatPercent(strongestRow.coveragePercent, { minimumFractionDigits: 1, maximumFractionDigits: 1 })}.` : '',
    biggestSalesRow ? `Najwyzsza sprzedaz w aktualnym zakresie jest u klienta ${getDetailedSalesTopCustomerLabel(biggestSalesRow)} - ${formatNumber(biggestSalesRow.salesValue)}.` : '',
    potentialRow ? `Najwiekszy potencjal domkniecia oferty ma klient ${getDetailedSalesTopCustomerLabel(potentialRow)} - braki ${formatNumber(potentialRow.missingCount)}; najwieksza szansa: ${potentialRow.topOpportunityGroupLabel || 'brak wskazanej grupy'}.` : '',
    weeklySummary.topCustomers.length ? `Najmocniejszy klient w sprzedazy tygodniowej: ${safeText(getDetailedSalesTopCustomerLabel(weeklySummary.topCustomers[0]))} - ${formatNumber(weeklySummary.topCustomers[0].salesValue)}.` : '',
    topMeetingRows.length ? `Na najblizsze spotkania warto przygotowac przede wszystkim: ${topMeetingRows.map(row => `#${row.rankingLabel || row.ranking} ${safeText(row.name)}`).join(', ')}.` : ''
  ]);

  if(weeklySummary.topCustomers.length){
    drawSectionTitle('Top klienci z raportu tygodniowego');
    drawBulletLines(weeklySummary.topCustomers.map((row, index) => `${index + 1}. ${getDetailedSalesTopCustomerLabel(row)} - ${formatNumber(row.salesValue)}`));
  }

  drawSectionTitle('Skrot klientow', 'Kazdy blok pokazuje sprzedaz w zakresie, udzial oferty Top, liczbe kupowanych grup oraz rozbicie po wszystkich kategoriach.');
  sortedCustomerRows.forEach(row => {
    const breakdownLine = (Array.isArray(row.groupBreakdown) ? row.groupBreakdown : [])
      .map(group => `${group.shortLabel} ${formatNumber(group.purchasedCount)}/${formatNumber(group.offerCount)} (${formatPercent(group.coveragePercent, { minimumFractionDigits: 1, maximumFractionDigits: 1 })})`)
      .join(' | ');
    const lineOne = `Sprzedaz w zakresie: ${formatNumber(row.salesValue)} | Kupione z Top: ${formatNumber(row.purchasedCount)}/${formatNumber(row.offerCount)} | Udzial: ${formatPercent(row.coveragePercent, { minimumFractionDigits: 1, maximumFractionDigits: 1 })} | Kategorie: ${formatNumber(row.purchasedGroupCount)}/${formatNumber(row.totalGroupCount || row.purchasedGroupCount + row.missingOnlyGroupCount)}`;
    const lineTwo = `Najwieksza szansa: ${row.topOpportunityGroupLabel || 'Brak wskazanej grupy'} | Braki w tej grupie: ${formatNumber(row.topOpportunityMissingCount)}`;
    const lineOneLines = pdf.splitTextToSize(safeText(lineOne), contentW - 24);
    const lineTwoLines = pdf.splitTextToSize(safeText(lineTwo), contentW - 24);
    const breakdownLines = pdf.splitTextToSize(safeText(breakdownLine), contentW - 66);
    const blockHeight = 36 + (lineOneLines.length * 11) + 4 + (lineTwoLines.length * 11) + 6 + (breakdownLines.length * 11) + 12;
    ensureSpace(blockHeight + 10);

    const blockTop = y;
    pdf.setDrawColor(226, 232, 240);
    pdf.setFillColor(255, 255, 255);
    pdf.roundedRect(margin, blockTop, contentW, blockHeight, 10, 10, 'FD');
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(12);
    pdf.setTextColor(15, 27, 45);
    pdf.text(safeText(getDetailedSalesTopCustomerLabel(row)), margin + 12, blockTop + 18);
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(9);
    pdf.setTextColor(71, 85, 105);
    let blockY = blockTop + 34;
    pdf.text(lineOneLines, margin + 12, blockY);
    blockY += lineOneLines.length * 11 + 4;
    pdf.text(lineTwoLines, margin + 12, blockY);
    blockY += lineTwoLines.length * 11 + 4;
    pdf.setFont('helvetica', 'bold');
    pdf.setTextColor(51, 65, 85);
    pdf.text(safeText('Grupy:'), margin + 12, blockY);
    pdf.setFont('helvetica', 'normal');
    pdf.setTextColor(71, 85, 105);
    pdf.text(breakdownLines, margin + 54, blockY);
    pdf.setTextColor(0);
    y = blockTop + blockHeight + 10;
  });

  drawSectionTitle('40 indeksow do rozmowy z klientami', 'Najwyzej rankingowe indeksy z calego portfela PH, ktorych nadal brakuje u najwiekszej liczby klientow.');
  const tableColumns = {
    ranking: margin,
    index: margin + 44,
    product: margin + 114,
    group: margin + 360,
    missing: margin + 452
  };
  const drawMeetingTableHeader = () => {
    ensureSpace(24);
    pdf.setFillColor(241, 245, 249);
    pdf.rect(margin, y, contentW, 18, 'F');
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(8);
    pdf.setTextColor(71, 85, 105);
    pdf.text('Ranking', tableColumns.ranking + 4, y + 12);
    pdf.text('Indeks', tableColumns.index + 4, y + 12);
    pdf.text('Produkt / producent', tableColumns.product + 4, y + 12);
    pdf.text('Grupa', tableColumns.group + 4, y + 12);
    pdf.text('Brakuje u', tableColumns.missing + 4, y + 12);
    pdf.setTextColor(0);
    y += 24;
  };

  if(!meetingRows.length){
    drawBulletLines(['Brak indeksow do wskazania w aktualnym zakresie raportu.']);
  }else{
    drawMeetingTableHeader();
    meetingRows.forEach(row => {
      const productText = `${row.name} | ${row.producer || 'Bez producenta'}`;
      const productLines = pdf.splitTextToSize(safeText(productText), 230).slice(0, 2);
      const rowHeight = Math.max(20, 8 + (productLines.length * 10));
      if(y + rowHeight > pageH - 30){
        startPage();
        drawSectionTitle('40 indeksow do rozmowy z klientami', 'Kontynuacja listy produktow do spotkan z klientami.');
        drawMeetingTableHeader();
      }
      pdf.setDrawColor(235, 241, 245);
      pdf.line(margin, y + rowHeight, pageW - margin, y + rowHeight);
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(8.5);
      pdf.text(safeText(row.rankingLabel || row.ranking), tableColumns.ranking + 4, y + 10);
      pdf.text(safeText(row.index), tableColumns.index + 4, y + 10);
      pdf.text(productLines, tableColumns.product + 4, y + 10);
      pdf.text(safeText(row.shortGroupLabel), tableColumns.group + 4, y + 10);
      pdf.text(safeText(`${formatNumber(row.missingCustomers)}/${formatNumber(row.totalCustomers)}`), tableColumns.missing + 4, y + 10);
      y += rowHeight + 4;
    });
  }

  const totalPages = pdf.getNumberOfPages();
  for(let page = 1; page <= totalPages; page += 1){
    pdf.setPage(page);
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(9);
    pdf.setTextColor(120);
    pdf.text(`Strona ${page} / ${totalPages}`, pageW - margin - 48, footerY);
    pdf.setTextColor(0);
  }

  const blob = pdf.output('blob');
  downloadBlob(blob, `${getDetailedSalesRepresentativeSummaryFileName(representativeLabel)}.pdf`);
}

async function importClientReportExcel(){
  const input = document.getElementById('client-reports-excel-input');
  const file = input?.files?.[0];
  if(!file) return;

  try{
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    const headers = rows.length ? Object.keys(rows[0]) : [];
    const fallbackMeta = buildClientReportImportFallback(file.name);
    const normalizedRows = rows
      .map(row => normalizeClientReportRow(row, headers))
      .filter(row => row.index);

    clientReportSourceRows = normalizedRows.map(row => {
      const representative = row.representative || fallbackMeta.representative;
      const customerCode = row.customerCode || fallbackMeta.customerCode;
      const customerName = row.customerName || fallbackMeta.customerName;
      return {
        ...row,
        representative,
        customerCode,
        customerName
      };
    });

    clientReportImportFile = file.name;
    clientReportSelectedRepresentative = '';
    clientReportSelectedCustomer = '';
    clientReportPurchasedRows = [];
    clientReportMissingRows = [];
    clientReportSummary = [];
    clientReportSelectedMissingIndexes = new Set();
    clientReportPurchasedFilters = { producer:'', ranking:'', group:'' };
    clientReportMissingFilters = { producer:'', ranking:'', group:'' };

    const representatives = getClientReportRepresentatives();
    if(representatives.length === 1){
      clientReportSelectedRepresentative = representatives[0];
      const customers = getClientReportCustomers();
      if(customers.length === 1){
        const onlyCustomer = customers[0];
        clientReportSelectedCustomer = `${onlyCustomer.code}|||${onlyCustomer.name}`;
        await generateClientComparisonReport();
      }
    }
  }catch(e){
    console.error('Client report import error', e);
    clientReportSourceRows = [];
    clientReportImportFile = '';
    clientReportSelectedRepresentative = '';
    clientReportSelectedCustomer = '';
    clientReportPurchasedRows = [];
    clientReportMissingRows = [];
    clientReportSummary = [];
    clientReportSelectedMissingIndexes = new Set();
    clientReportPurchasedFilters = { producer:'', ranking:'', group:'' };
    clientReportMissingFilters = { producer:'', ranking:'', group:'' };
  }

  renderReportsView();
}

async function importTopSuggestionsExcel(){
  const input = document.getElementById('top-suggestions-excel-input');
  const file = input?.files?.[0];
  if(!file) return;

  try{
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const sheetName = wb.SheetNames.includes('Export') ? 'Export' : wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    topSuggestionSourceRows = normalizeWeeklySalesRowsFromMatrix(matrix);
    topSuggestionImportFile = file.name;
    topSuggestionSelectedRepresentative = '';
    topSuggestionSelectedCustomer = '';
    topSuggestionLimit = '';
    topSuggestionGeneratedRows = [];
    topSuggestionPurchasedRows = [];
    topSuggestionSummary = null;

    const representatives = getTopSuggestionRepresentatives();
    if(representatives.length === 1){
      topSuggestionSelectedRepresentative = representatives[0];
    }
  }catch(e){
    console.error('Top suggestions import error', e);
    topSuggestionSourceRows = [];
    topSuggestionGeneratedRows = [];
    topSuggestionPurchasedRows = [];
    topSuggestionImportFile = '';
    topSuggestionSelectedRepresentative = '';
    topSuggestionSelectedCustomer = '';
    topSuggestionLimit = '';
    topSuggestionSummary = null;
  }

  renderReportsView();
}

async function setClientReportRepresentative(value){
  clientReportSelectedRepresentative = value;
  clientReportSelectedCustomer = '';
  clientReportPurchasedRows = [];
  clientReportMissingRows = [];
  clientReportSummary = [];
  clientReportSelectedMissingIndexes = new Set();
  renderReportsView();
}

async function setClientReportCustomer(value){
  clientReportSelectedCustomer = value;
  await generateClientComparisonReport();
  renderReportsView();
}

async function setTopSuggestionRepresentative(value){
  topSuggestionSelectedRepresentative = value;
  topSuggestionSelectedCustomer = '';
  topSuggestionGeneratedRows = [];
  topSuggestionPurchasedRows = [];
  topSuggestionSummary = null;
  renderReportsView();
}

async function setTopSuggestionCustomer(value){
  topSuggestionSelectedCustomer = value;
  await generateTopSuggestionsReport();
  renderReportsView();
}

async function setTopSuggestionLimit(value){
  topSuggestionLimit = value;
  await generateTopSuggestionsReport();
  renderReportsView();
}

function toggleClientMissingSelection(checkbox){
  const index = normalizeIndexValue(checkbox?.getAttribute('data-index'));
  if(!index) return;
  if(checkbox.checked){
    clientReportSelectedMissingIndexes.add(index);
  }else{
    clientReportSelectedMissingIndexes.delete(index);
  }
}

function selectAllClientMissing(){
  clientReportSelectedMissingIndexes = new Set(
    clientReportMissingRows.map(row => normalizeIndexValue(row.INDEKS)).filter(Boolean)
  );
  renderReportsView();
}

function clearAllClientMissing(){
  clientReportSelectedMissingIndexes = new Set();
  renderReportsView();
}

function setClientTableFilter(type, key, value){
  if(type === 'purchased'){
    clientReportPurchasedFilters[key] = value;
  }else{
    clientReportMissingFilters[key] = value;
  }
  renderReportsView();
}

function setClientRecommendationCoverOption(checked){
  clientRecommendationIncludeCover = !!checked;
}

function getSelectedClientRecommendationRows(){
  return clientReportMissingRows.filter(row =>
    clientReportSelectedMissingIndexes.has(normalizeIndexValue(row.INDEKS))
  );
}

function getClientRecommendationFileName(){
  const [customerCode, customerName] = (clientReportSelectedCustomer || '').split('|||');
  const safePart = (value, fallback) => String(value || fallback)
    .trim()
    .replace(/[\\/:*?"<>|]+/g, ' ')
    .replace(/\s+/g, ' ');
  return `${safePart(customerCode, 'Klient')} ${safePart(customerName, 'Bez nazwy')} - Top Rumunia`;
}

function getClientRecommendationCustomerInfo(){
  const [customerCode, customerName] = (clientReportSelectedCustomer || '').split('|||');
  return {
    code: String(customerCode || '').trim(),
    name: String(customerName || '').trim()
  };
}

function getClientQuickOrderContext(){
  return {
    customer: getClientRecommendationCustomerInfo(),
    representativeName: String(clientReportSelectedRepresentative || '').trim(),
    offerTitle: getClientRecommendationFileName(),
    message: getClientRecommendationMessage(),
    items: getSelectedClientRecommendationRows().map(row => withQuickOrderPriceData({
      index: row.INDEKS,
      name: row.NAZWA,
      producer: row.SKROT_PRODUCENTA,
      group: row['Grupa produktowa'],
      ean: row['KOD EAN'] || '',
      ranking: row.Ranking || '',
      imageBaseUrl: imageBaseUrlRumunia
    }))
  };
}

window.getClientQuickOrderContext = getClientQuickOrderContext;

function getDetailedSalesTopQuickOrderCustomerInfo(){
  const [customerCode = '', customerName = ''] = String(detailedSalesComparisonCustomer || '').split('|||');
  return {
    code: String(customerCode || '').trim(),
    name: String(customerName || '').trim()
  };
}

function getDetailedSalesTopQuickOrderContext(){
  return {
    customer: getDetailedSalesTopQuickOrderCustomerInfo(),
    representativeName: '',
    offerTitle: getDetailedSalesTopRecommendationFileName(),
    message: getClientRecommendationMessage(),
    items: getSelectedDetailedSalesTopMissingRows().map(row => withQuickOrderPriceData({
      index: row.INDEKS,
      name: row.NAZWA,
      producer: row.SKROT_PRODUCENTA,
      group: row['Grupa produktowa'],
      ean: row['KOD EAN'] || '',
      ranking: row.Ranking || '',
      imageBaseUrl: imageBaseUrlRumunia
    }))
  };
}

window.getDetailedSalesTopQuickOrderContext = getDetailedSalesTopQuickOrderContext;

function getQuickOrderPriceData(index){
  const normalizedIndex = normalizeIndexValue(index);
  if(!normalizedIndex) return null;

  const sourceMap = catalogPriceMap instanceof Map && catalogPriceMap.size
    ? catalogPriceMap
    : (adminCatalogPriceMapCache instanceof Map && adminCatalogPriceMapCache.size ? adminCatalogPriceMapCache : null);

  if(!sourceMap || !sourceMap.has(normalizedIndex)){
    return null;
  }

  const entry = sourceMap.get(normalizedIndex) || {};
  const price = String(entry?.price || '').trim();
  const unit = String(entry?.unit || '').trim();
  if(!price) return null;

  return {
    price,
    unit,
    priceNumeric: String(entry?.priceNumeric || '').trim(),
    currencySymbol: String(entry?.currencySymbol || '').trim(),
    minimumOrderQuantity: String(entry?.minimumOrderQuantity || '').trim(),
    minimumOrderUnit: String(entry?.minimumOrderUnit || '').trim(),
    minimumOrderValue: String(entry?.minimumOrderValue || '').trim(),
    name: String(entry?.name || '').trim(),
    producer: String(entry?.producer || '').trim(),
    ean: String(entry?.ean || '').trim()
  };
}

function withQuickOrderPriceData(item){
  const baseItem = item && typeof item === 'object' ? { ...item } : {};
  const priceData = getQuickOrderPriceData(baseItem.index || baseItem.INDEKS || '');
  if(!priceData){
    return baseItem;
  }

  const enrichedItem = {
    ...baseItem,
    price: priceData.price,
    unit: priceData.unit
  };

  [
    'priceNumeric',
    'currencySymbol',
    'minimumOrderQuantity',
    'minimumOrderUnit',
    'minimumOrderValue'
  ].forEach(key => {
    if(String(priceData[key] || '').trim()){
      enrichedItem[key] = priceData[key];
    }
  });

  if(!String(enrichedItem.name || enrichedItem.NAZWA || '').trim() && priceData.name){
    enrichedItem.name = priceData.name;
  }
  if(!String(enrichedItem.producer || enrichedItem.PRODUCENT || enrichedItem.SKROT_PRODUCENTA || '').trim() && priceData.producer){
    enrichedItem.producer = priceData.producer;
  }
  if(!String(enrichedItem.ean || enrichedItem['KOD EAN'] || '').trim() && priceData.ean){
    enrichedItem.ean = priceData.ean;
  }

  return enrichedItem;
}

window.withQuickOrderPriceData = withQuickOrderPriceData;

function getGenericQuickOrderContext(){
  const items = Array.from(selectedProducts.values()).map(row => withQuickOrderPriceData({
    index: row?.INDEKS || row?.Indeks || row?.index || row?.ID || row?.id || row?.['Numer katalogowy'] || '',
    name: row?.NAZWA || row?.Nazwa || row?.name || '',
    producer: row?.PRODUCENT || row?.Producent || row?.producer || row?.SKROT_PRODUCENTA || row?.Skrot_producenta || '',
    group: row?.GRUPA || row?.Grupa || row?.group || row?.['Grupa produktowa'] || row?.['Grupa Produktowa'] || row?.NAZWA_GRUPA || '',
    ean: row?.['KOD EAN'] || row?.['Kod EAN'] || row?.ean || '',
    ranking: row?.Ranking || row?.ranking || '',
    imageBaseUrl: currentImageBaseUrl
  })).filter(item => String(item.index || '').trim() && String(item.name || '').trim());

  return {
    customer: {
      code: '',
      name: ''
    },
    representativeName: '',
    offerTitle: String(currentCategoryName || currentCategorySlug || 'Szybkie zamówienie').trim() || 'Szybkie zamówienie',
    message: {
      pl: '',
      en: ''
    },
    items
  };
}

window.getGenericQuickOrderContext = getGenericQuickOrderContext;

function getClientRecommendationMessage(){
  return {
    pl: 'Przygotowalismy dla Ciebie wyselekcjonowana liste produktow z oferty rumunskiej na podstawie ostatnich Twoich zakupow. Produkty podzielilismy na grupy i dobralismy je tak, aby wspieraly rotacje, uzupelnialy braki w ofercie i odpowiadaly na realne potrzeby klientow rumunskich. Oferta bedzie dostepna na mastersale.eu. Szczegoly uzyskasz u swojego opiekuna handlowego.',
    en: 'We have prepared a curated list of Romanian products based on your recent purchases. We grouped the products and selected them to support turnover, fill assortment gaps, and meet the real needs of Romanian shoppers. The offer will be available at mastersale.eu. For details, please contact your account manager.'
  };
}

function openClientRecommendationMessageModal(){
  if(!reportsMessageModal || !reportsMessageText) return;
  const message = getClientRecommendationMessage();
  reportsMessageText.innerHTML = `
    <div class="reports-message-hero">
      <div class="reports-message-kicker">World Food Rumunia</div>
      <div class="reports-message-sub">Rekomendacja produktowa dla klienta</div>
    </div>
    <div class="reports-message-section">
      <div class="reports-message-label">PL</div>
      <p>${escapeHtml(message.pl)}</p>
    </div>
    <div class="reports-message-section">
      <div class="reports-message-label">EN</div>
      <p>${escapeHtml(message.en)}</p>
    </div>
  `;
  reportsMessageModal.classList.remove('hidden');
}

function getClientRecommendationMessageText(){
  const message = getClientRecommendationMessage();
  return `WORLD FOOD RUMUNIA

PL:
${message.pl}

EN:
${message.en}`;
}

async function createClientRecommendationCoverDataUrl(){
  const message = getClientRecommendationMessage();
  const customer = getClientRecommendationCustomerInfo();
  const logoImage = await loadImageAsDataUrl(maspoLogoUrl).catch(() => null);
  const canvas = document.createElement('canvas');
  const width = 1600;
  const height = 2260;
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext('2d');

  const gradient = ctx.createLinearGradient(0, 0, width, height);
  gradient.addColorStop(0, '#f8fbff');
  gradient.addColorStop(1, '#eef7f1');
  ctx.fillStyle = gradient;
  ctx.fillRect(0, 0, width, height);

  ctx.fillStyle = '#ffffff';
  roundRect(ctx, 70, 70, width - 140, height - 140, 48);
  ctx.fill();

  ctx.fillStyle = '#002b7f';
  ctx.fillRect(90, 90, (width - 180) / 3, 26);
  ctx.fillStyle = '#fcd116';
  ctx.fillRect(90 + (width - 180) / 3, 90, (width - 180) / 3, 26);
  ctx.fillStyle = '#ce1126';
  ctx.fillRect(90 + ((width - 180) / 3) * 2, 90, (width - 180) / 3, 26);

  const circleX = width - 250;
  const circleY = 260;
  ctx.beginPath();
  ctx.arc(circleX, circleY, 120, 0, Math.PI * 2);
  ctx.closePath();
  ctx.fillStyle = '#f4f8ff';
  ctx.fill();
  ctx.lineWidth = 6;
  ctx.strokeStyle = '#dbeafe';
  ctx.stroke();

  ctx.fillStyle = '#002b7f';
  ctx.fillRect(circleX - 72, circleY - 56, 48, 112);
  ctx.fillStyle = '#fcd116';
  ctx.fillRect(circleX - 24, circleY - 56, 48, 112);
  ctx.fillStyle = '#ce1126';
  ctx.fillRect(circleX + 24, circleY - 56, 48, 112);

  ctx.fillStyle = '#0f1b2d';
  ctx.font = '700 84px Georgia';
  ctx.fillText('World Food Rumunia', 110, 220);

  ctx.fillStyle = '#168a3f';
  ctx.font = '700 28px Arial';
  ctx.fillText('TOP RUMUNIA', 112, 276);

  ctx.fillStyle = '#475569';
  ctx.font = '500 32px Arial';
  ctx.fillText('Rekomendacja produktowa dla klienta', 110, 324);

  const hasCustomerInfo = !!(customer.code || customer.name);
  const customerBoxY = 352;
  const customerBoxH = 116;
  const plBoxY = hasCustomerInfo ? 500 : 390;
  const plTextY = hasCustomerInfo ? 645 : 535;
  const enBoxY = hasCustomerInfo ? 1160 : 1050;
  const enTextY = hasCustomerInfo ? 1298 : 1188;
  const ctaY = hasCustomerInfo ? 1880 : 1770;

  if(customer.code || customer.name){
    ctx.fillStyle = '#ffffff';
    roundRect(ctx, 110, customerBoxY, 860, customerBoxH, 28);
    ctx.fill();
    ctx.strokeStyle = '#dbe7f3';
    ctx.lineWidth = 2;
    ctx.stroke();

    ctx.fillStyle = '#64748b';
    ctx.font = '600 22px Arial';
    ctx.fillText('Klient', 144, customerBoxY + 38);

    ctx.fillStyle = '#0f1b2d';
    ctx.font = '700 34px Arial';
    const customerLine = [customer.code, customer.name].filter(Boolean).join(' - ');
    drawWrappedText(ctx, customerLine, 144, customerBoxY + 84, 780, 40);
  }

  ctx.fillStyle = 'rgba(0, 43, 127, 0.08)';
  ctx.beginPath();
  ctx.arc(1180, 560, 150, 0, Math.PI * 2);
  ctx.fill();
  ctx.fillStyle = '#002b7f';
  ctx.fillRect(1098, 500, 58, 122);
  ctx.fillStyle = '#fcd116';
  ctx.fillRect(1156, 500, 58, 122);
  ctx.fillStyle = '#ce1126';
  ctx.fillRect(1214, 500, 58, 122);

  ctx.fillStyle = '#ffffff';
  roundRect(ctx, 90, plBoxY, width - 180, 610, 36);
  ctx.fill();
  ctx.strokeStyle = '#d8e7dc';
  ctx.lineWidth = 3;
  ctx.stroke();

  ctx.fillStyle = '#002b7f';
  roundRect(ctx, 132, plBoxY + 52, 220, 16, 8);
  ctx.fill();
  ctx.fillStyle = '#0f1b2d';
  ctx.font = '400 38px Arial';
  drawWrappedText(ctx, message.pl, 140, plTextY, width - 320, 54);

  ctx.fillStyle = '#ffffff';
  roundRect(ctx, 90, enBoxY, width - 180, 540, 36);
  ctx.fill();
  ctx.strokeStyle = '#d8e7dc';
  ctx.lineWidth = 3;
  ctx.stroke();

  ctx.fillStyle = '#ce1126';
  roundRect(ctx, 132, enBoxY + 52, 220, 16, 8);
  ctx.fill();
  ctx.fillStyle = '#0f1b2d';
  ctx.font = '400 38px Arial';
  drawWrappedText(ctx, message.en, 140, enTextY, width - 320, 54);

  ctx.fillStyle = '#102c6b';
  roundRect(ctx, 90, ctaY, width - 180, 170, 34);
  ctx.fill();
  ctx.fillStyle = '#ffffff';
  ctx.font = '700 46px Arial';
  ctx.fillText('Oferta bedzie dostepna na mastersale.eu', 145, ctaY + 80);
  ctx.font = '400 28px Arial';
  ctx.fillText('Szczegoly uzyskasz u swojego opiekuna handlowego.', 145, ctaY + 134);

  if(logoImage?.dataUrl){
    const img = new Image();
    await new Promise((resolve, reject) => {
      img.onload = resolve;
      img.onerror = reject;
      img.src = logoImage.dataUrl;
    }).catch(() => null);

    if(img.naturalWidth && img.naturalHeight){
      const maxW = 280;
      const maxH = 82;
      const scale = Math.min(maxW / img.naturalWidth, maxH / img.naturalHeight, 1);
      const drawW = img.naturalWidth * scale;
      const drawH = img.naturalHeight * scale;
      const drawX = width - drawW - 150;
      const drawY = 2088;
      ctx.drawImage(img, drawX, drawY, drawW, drawH);

      ctx.fillStyle = '#64748b';
      ctx.font = '600 22px Arial';
      ctx.fillText('World Food Partner', drawX + 20, Math.min(drawY + drawH + 28, 2230));
    }
  }

  ctx.fillStyle = '#64748b';
  ctx.font = '400 24px Arial';
  ctx.fillText('World Food | Romanian assortment recommendations', 110, 2230);

  return canvas.toDataURL('image/jpeg', 0.92);
}

function roundRect(ctx, x, y, width, height, radius){
  ctx.beginPath();
  ctx.moveTo(x + radius, y);
  ctx.lineTo(x + width - radius, y);
  ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
  ctx.lineTo(x + width, y + height - radius);
  ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
  ctx.lineTo(x + radius, y + height);
  ctx.quadraticCurveTo(x, y + height, x, y + height - radius);
  ctx.lineTo(x, y + radius);
  ctx.quadraticCurveTo(x, y, x + radius, y);
  ctx.closePath();
}

function drawWrappedText(ctx, text, x, y, maxWidth, lineHeight){
  const words = String(text || '').split(/\s+/);
  let line = '';
  let currentY = y;
  words.forEach(word => {
    const testLine = line ? `${line} ${word}` : word;
    if(ctx.measureText(testLine).width > maxWidth && line){
      ctx.fillText(line, x, currentY);
      line = word;
      currentY += lineHeight;
    }else{
      line = testLine;
    }
  });
  if(line){
    ctx.fillText(line, x, currentY);
  }
}

function exportClientRecommendationXlsx(){
  const selectedRows = getSelectedClientRecommendationRows();
  if(!selectedRows.length) return;
  const cols = ['INDEKS', 'NAZWA', 'SKROT_PRODUCENTA', 'KOD EAN', 'Ranking', 'Grupa produktowa'];
  const rows = [cols, ...selectedRows.map(row => cols.map(col => row[col] ?? ''))];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'rekomendowane');
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  downloadBlob(blob, `${getClientRecommendationFileName()}.xlsx`);
  openClientRecommendationMessageModal();
}

async function createClientRecommendationCatalog(){
  const selectedRows = getSelectedClientRecommendationRows();
  if(!selectedRows.length) return;
  selectedProducts = new Map();
  selectedRows.forEach(row => {
    const idx = normalizeIndexValue(row.INDEKS);
    if(idx){
      selectedProducts.set(idx, {
        INDEKS: row.INDEKS,
        NAZWA: row.NAZWA,
        PRODUCENT: row.SKROT_PRODUCENTA,
        GRUPA: row['Grupa produktowa'],
        'KOD EAN': row['KOD EAN'] || ''
      });
    }
  });
  setFullData(Array.from(selectedProducts.values()), ['INDEKS', 'NAZWA', 'PRODUCENT', 'GRUPA', 'KOD EAN']);
  activeContainer = reportsContainer;
  currentCategorySlug = getClientRecommendationFileName();
  const coverDataUrl = clientRecommendationIncludeCover
    ? await createClientRecommendationCoverDataUrl()
    : null;
  catalogOptionsOverride = {
    groupByField: 'GRUPA',
    sectionTitle: 'Grupa produktowa',
    coverDataUrl
  };
  await createCatalog();
}

function exportReportsXlsx(){
  if(!reportsGeneratedRows.length) return;
  const cols = ['INDEKS', 'NAZWA', 'SKROT_PRODUCENTA', 'KOD EAN', 'NAZWA_GRUPA', 'Ranking', 'Grupa produktowa'];
  const rows = [cols, ...reportsGeneratedRows.map(row => cols.map(col => row[col] ?? ''))];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'top sprzedaz');
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  downloadBlob(blob, 'top_sprzedaz.xlsx');
}

function exportTopSuggestionsXlsx(){
  if(!topSuggestionGeneratedRows.length && !topSuggestionPurchasedRows.length) return;
  const rows = [
    ['Typ', 'Nr producenta', 'Producent', 'Sprzedaż', 'Ile indeksów w Top', 'Najwyższy ranking', 'Przykładowy indeks', 'Przykładowy produkt'],
    ...topSuggestionPurchasedRows.map(row => [
      'Kupiony',
      row.producerCode || '',
      row.producer || '',
      row.salesTotal || 0,
      '',
      '',
      '',
      ''
    ]),
    ...topSuggestionGeneratedRows.map(row => [
      'Niekupiony',
      '',
      row.producer || '',
      '',
      row.productCount || '',
      row.bestRanking || '',
      row.exampleIndex || '',
      row.exampleName || ''
    ])
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'producenci top');
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  downloadBlob(blob, 'producenci_top_rumunia.xlsx');
}

function render(){
  if(!fullData.length && !isLoading) return;
  clearAllContentContainers();

  const cols = getCurrentColumns();
  const producerKey = findColumn(cols, ['producent', 'producer']);
  const groupKey = findColumn(cols, ['grupa', 'group']);
  const nameKey = findColumn(cols, ['nazwa', 'name']);
  const eanKey = findColumn(cols, ['kod ean', 'ean']);
  const indexKey = findColumn(cols, ['indeks', 'index', 'id', 'numer katalogowy', 'sku number']);

  const skeletonCols = cols.length ? cols : ['INDEKS','NAZWA','RANKING','GRUPA','PRODUCENT','KOD EAN'];
  const skeletonRows = Array.from({ length: 8 }).map(() => `
      <tr class="skeleton-row">
        ${skeletonCols.map(() => `<td><div class="skeleton-line"></div></td>`).join('')}
      </tr>
    `).join('');

  let html = `
    ${viewMode === 'table' ? `
      <div class="import-bar">
        <div>
          <div class="import-title">Import Excel</div>
          <div class="import-sub">Filtrowanie listy po indeksach z pliku</div>
        </div>
        <div class="import-actions">
          <input id="index-import-input" class="import-input" type="file" accept=".xlsx,.xls">
          <button class="btn-outline" onclick="importIndexExcel()">Importuj Excel</button>
          <button class="btn-outline" onclick="clearIndexImport()">Wyczyść import</button>
        </div>
        <div class="import-info">
          ${importedIndexSet ? `Zaimportowano: ${importedIndexCount} (${escapeHtml(importedIndexFile)})` : 'Brak importu'}
        </div>
      </div>
      <div class="actions">
        <button onclick="exportCSV()">Zapisz CSV</button>
        <button onclick="exportXLS()">Zapisz XLS</button>
        <button onclick="createCatalog()">Utwórz katalog</button>
        <button onclick="clearCatalogData()" ${selectedProducts.size || catalogBlobUrl ? '' : 'disabled'}>Wyczyść dane</button>
        ${catalogBlobUrl ? `<button onclick="downloadCatalog()">Zapisz katalog PDF</button>` : ''}
      </div>
    ` : ''}

    ${viewMode === 'pdf' ? `
      <div class="pdf-preview">
        <iframe src="${pdfPreviewUrl + encodeURIComponent(getActivePdfUrl())}&t=${Date.now()}" title="Podgląd PDF"></iframe>
      </div>
    ` : viewMode === 'catalog' ? `
      <div class="actions">
        <button onclick="exportCSV()">Zapisz CSV</button>
        <button onclick="exportXLS()">Zapisz XLS</button>
        <button onclick="createCatalog()">Utwórz katalog</button>
        <button onclick="clearCatalogData()" ${selectedProducts.size || catalogBlobUrl ? '' : 'disabled'}>Wyczyść dane</button>
        ${catalogBlobUrl ? `<button onclick="downloadCatalog()">Zapisz katalog PDF</button>` : ''}
      </div>
      <div class="pdf-preview">
        ${catalogLoading ? `<div class="catalog-loading">
          <div class="catalog-spinner"></div>
          <div>Tworzę katalog...</div>
        </div>` : ''}
        ${catalogBlobUrl ? `<iframe src="${catalogBlobUrl}#view=FitH&zoom=page-fit" title="Katalog PDF"></iframe>` : ''}
      </div>
    ` : `
      <div class="filters">
        ${producerKey ? `
          <div class="filter">
            <label>Producent</label>
            <select onchange="setFilter('producer', this.value)">
              <option value="">Wszyscy</option>
              ${uniqueValues(fullData, producerKey).map(v =>
                `<option value="${escapeHtml(v)}" ${filters.producer === String(v) ? 'selected' : ''}>${escapeHtml(v)}</option>`
              ).join('')}
            </select>
          </div>
        ` : ''}
        ${groupKey ? `
          <div class="filter">
            <label>Grupa</label>
            <select onchange="setFilter('group', this.value)">
              <option value="">Wszystkie</option>
              ${uniqueValues(fullData, groupKey).map(v =>
                `<option value="${escapeHtml(v)}" ${filters.group === String(v) ? 'selected' : ''}>${escapeHtml(v)}</option>`
              ).join('')}
            </select>
          </div>
        ` : ''}
        ${nameKey ? `
          <div class="filter">
            <label>Nazwa</label>
            <input type="text" value="${escapeAttr(filters.name)}" oninput="setFilter('name', this.value, true)" placeholder="Szukaj po nazwie">
          </div>
        ` : ''}
        ${eanKey ? `
          <div class="filter">
            <label>Kod EAN</label>
            <input type="text" value="${escapeAttr(filters.ean)}" oninput="setFilter('ean', this.value, true)" placeholder="Np. 590...">
          </div>
        ` : ''}
        ${indexKey ? `
          <div class="filter">
            <label>Indeks</label>
            <input type="text" value="${escapeAttr(filters.index)}" oninput="setFilter('index', this.value, true)" placeholder="Np. 12345">
          </div>
        ` : ''}
        <div class="filter">
          <label>Ilość pozycji</label>
          <input type="number" min="1" value="${escapeAttr(filters.limit)}" oninput="setFilter('limit', this.value, true)" placeholder="Np. 50">
        </div>
        <div class="filter">
          <label>&nbsp;</label>
          <button class="btn-outline" onclick="resetFilters(); render();">Wyczyść filtry</button>
        </div>
      </div>

      <div class="filters-actions">
        <button class="btn-outline" onclick="selectAllVisible()">Zaznacz wszystkie</button>
        <button class="btn-outline" onclick="clearSelected()">Odznacz zaznaczone</button>
      </div>

      <div class="table-wrap">
        <table>
          <thead>
            <tr>${(cols.length ? cols : skeletonCols).map(c => renderHeaderCell(c, indexKey)).join('')}</tr>
          </thead>
          <tbody id="data-tbody">
            ${isLoading ? skeletonRows : ''}
          </tbody>
        </table>
      </div>
      <div id="expand-container"></div>
    `}
  `;

  activeContainer.innerHTML = html;

  if(viewMode === 'table' && !isLoading){
    updateTable();
  }

  document.dispatchEvent(new CustomEvent('wf:quick-order-refresh'));
}

function expandTable(){
  expanded = true;
  viewMode = 'table';
  updateTable();
}

function clearIndexImport(){
  importedIndexSet = null;
  importedIndexCount = 0;
  importedIndexFile = '';
  render();
}

async function importIndexExcel(){
  const input = document.getElementById('index-import-input');
  const file = input?.files?.[0];
  if(!file) return;
  try{
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const set = new Set();
    rows.slice(1).forEach(row => {
      const idx = normalizeIndexValue(row[0]);
      if(idx) set.add(idx);
    });
    importedIndexSet = set.size ? set : null;
    importedIndexCount = set.size;
    importedIndexFile = file.name;
  }catch(e){
    console.error(e);
    importedIndexSet = null;
    importedIndexCount = 0;
    importedIndexFile = '';
  }
  render();
}

function setFilter(key, value, soft){
  filters[key] = value;
  expanded = false;
  if(soft){
    updateTable();
  }else{
    render();
  }
}

function normalizeIndexValue(value){
  return String(value ?? '')
    .trim()
    .replace(/\.0+$/, '')
    .replace(/\s+/g, '');
}

function normalizeCustomerCodeValue(value){
  return String(value ?? '')
    .trim()
    .replace(/\.0+$/, '')
    .replace(/\s+/g, '');
}

function formatDiscountPercent(value){
  const cleaned = String(value ?? '').trim();
  if(!cleaned) return '';
  return cleaned.includes('%') ? cleaned : `${cleaned}%`;
}

function normalizeProducerValue(value){
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function dedupeRowsByIndex(rows, indexKey){
  if(!indexKey || !Array.isArray(rows) || !rows.length) return rows;
  const seen = new Set();
  return rows.filter(row => {
    const idx = normalizeIndexValue(row?.[indexKey]);
    if(!idx) return true;
    if(seen.has(idx)) return false;
    seen.add(idx);
    return true;
  });
}

const listingCodeInput = document.getElementById('listing-code');
const listingStartBtn = document.getElementById('listing-start-btn');
const listingStopBtn = document.getElementById('listing-stop-btn');
const listingSearchBtn = document.getElementById('listing-search-btn');
const listingExportBtn = document.getElementById('listing-export-btn');
const listingExcelInput = document.getElementById('listing-excel');
const listingImportBtn = document.getElementById('listing-import-btn');
const listingExportPdfBtn = document.getElementById('listing-export-pdf-btn');
const listingHead = document.getElementById('listing-head');
const listingBody = document.getElementById('listing-body');
const listingConfirmModal = document.getElementById('listing-confirm-modal');
const listingConfirmYes = document.getElementById('listing-confirm-yes');
const listingConfirmNo = document.getElementById('listing-confirm-no');
let pendingListingAdd = null;
const listingFoundModal = document.getElementById('listing-found-modal');
const listingFoundText = document.getElementById('listing-found-text');
const listingFoundOk = document.getElementById('listing-found-ok');
const listingAddModal = document.getElementById('listing-add-modal');
const listingAddYes = document.getElementById('listing-add-yes');
const listingAddNo = document.getElementById('listing-add-no');
let pendingMaspoAdd = null;
const listingAddImage = document.getElementById('listing-add-image');
const listingAddName = document.getElementById('listing-add-name');
const listingMarionModal = document.getElementById('listing-marion-modal');
const listingMarionYes = document.getElementById('listing-marion-yes');
const listingMarionNo = document.getElementById('listing-marion-no');
const listingMarionImage = document.getElementById('listing-marion-image');
const listingMarionName = document.getElementById('listing-marion-name');
const reportsMessageModal = document.getElementById('reports-message-modal');
const reportsMessageText = document.getElementById('reports-message-text');
const reportsMessageCopy = document.getElementById('reports-message-copy');
const reportsMessageClose = document.getElementById('reports-message-close');
let pendingMarionAdd = null;

if(listingStartBtn){
  listingStartBtn.addEventListener('click', startListingScanner);
}
if(listingStopBtn){
  listingStopBtn.addEventListener('click', stopListingScanner);
}
if(listingSearchBtn){
  listingSearchBtn.addEventListener('click', () => {
    const code = listingCodeInput?.value.trim();
    if(code) searchListingByCode(code);
  });
}
if(listingExportBtn){
  listingExportBtn.addEventListener('click', exportListingXlsx);
}
if(listingExportPdfBtn){
  listingExportPdfBtn.addEventListener('click', exportListingPdf);
}
if(listingImportBtn){
  listingImportBtn.addEventListener('click', importListingExcel);
}
if(listingConfirmYes){
  listingConfirmYes.addEventListener('click', () => {
    if(pendingListingAdd){
      applyListingAdd(pendingListingAdd.code, pendingListingAdd.matches, true);
      pendingListingAdd = null;
    }
    if(listingConfirmModal) listingConfirmModal.classList.add('hidden');
  });
}
if(listingConfirmNo){
  listingConfirmNo.addEventListener('click', () => {
    pendingListingAdd = null;
    if(listingConfirmModal) listingConfirmModal.classList.add('hidden');
  });
}
if(listingFoundOk){
  listingFoundOk.addEventListener('click', () => {
    if(listingFoundModal) listingFoundModal.classList.add('hidden');
  });
}
if(listingAddYes){
  listingAddYes.addEventListener('click', () => {
    if(pendingMaspoAdd){
      applyListingAdd(pendingMaspoAdd.code, pendingMaspoAdd.matches, false);
      pendingMaspoAdd = null;
    }
    if(listingAddModal) listingAddModal.classList.add('hidden');
  });
}
if(listingAddNo){
  listingAddNo.addEventListener('click', () => {
    pendingMaspoAdd = null;
    if(listingAddModal) listingAddModal.classList.add('hidden');
  });
}
if(listingMarionYes){
  listingMarionYes.addEventListener('click', () => {
    if(pendingMarionAdd){
      applyListingAdd(pendingMarionAdd.code, pendingMarionAdd.matches, false);
      pendingMarionAdd = null;
    }
    if(listingMarionModal) listingMarionModal.classList.add('hidden');
  });
}
if(listingMarionNo){
  listingMarionNo.addEventListener('click', () => {
    pendingMarionAdd = null;
    if(listingMarionModal) listingMarionModal.classList.add('hidden');
  });
}
if(reportsMessageClose){
  reportsMessageClose.addEventListener('click', () => {
    if(reportsMessageModal) reportsMessageModal.classList.add('hidden');
  });
}
if(reportsMessageCopy){
  reportsMessageCopy.addEventListener('click', async () => {
    const text = getClientRecommendationMessageText();
    if(!text) return;
    try{
      await navigator.clipboard.writeText(text);
      reportsMessageCopy.textContent = 'Skopiowano';
      setTimeout(() => {
        if(reportsMessageCopy) reportsMessageCopy.textContent = 'Kopiuj';
      }, 1500);
    }catch(e){
      console.error('Clipboard copy failed', e);
    }
  });
}
if(listingCodeInput){
  listingCodeInput.addEventListener('keydown', (e) => {
    if(e.key === 'Enter'){
      e.preventDefault();
      const code = listingCodeInput.value.trim();
      if(code) searchListingByCode(code);
    }
  });
}

async function importListingExcel(){
  const file = listingExcelInput?.files?.[0];
  if(!file) return;
  try{
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const eans = rows.slice(1)
      .map(r => String(r[0] ?? '').replace(/\D/g, ''))
      .filter(Boolean);
    for(const code of eans){
      const matches = await searchMarionDatabase(code);
      if(matches.length){
        applyListingAdd(code, matches, false);
      }
    }
  }catch(e){
    console.error(e);
  }
}

function startListingScanner(){
  const target = document.getElementById('listing-camera');
  if(!target || typeof Quagga === 'undefined') return;
  target.classList.add('is-active');
  Quagga.init({
    inputStream: {
      type: 'LiveStream',
      target,
      constraints: {
        facingMode: { exact: 'environment' },
        width: { min: 640 },
        height: { min: 480 }
      }
    },
    decoder: {
      readers: ['ean_reader', 'ean_8_reader', 'upc_reader', 'upc_e_reader']
    },
    locate: true
  }, err => {
    if(err){
      // fallback bez exact
      Quagga.init({
        inputStream: {
          type: 'LiveStream',
          target,
          constraints: {
            facingMode: 'environment'
          }
        },
        decoder: {
          readers: ['ean_reader', 'ean_8_reader', 'upc_reader', 'upc_e_reader']
        },
        locate: true
      }, err2 => {
        if(err2){
          console.error(err2);
          return;
        }
        Quagga.start();
      });
      return;
    }
    Quagga.start();
    setTimeout(() => {
      const video = target.querySelector('video');
      if(video){
        video.setAttribute('playsinline', 'true');
        video.setAttribute('webkit-playsinline', 'true');
        video.muted = true;
      }
    }, 100);
  });

  Quagga.offDetected();
  Quagga.onDetected(result => {
    const code = result?.codeResult?.code;
    if(code){
      if(listingCooldown) return;
      listingCooldown = true;
      setTimeout(() => { listingCooldown = false; }, 1200);
      listingCodeInput.value = code;
      searchListingByCode(code, true);
    }
  });
}

function stopListingScanner(){
  if(typeof Quagga === 'undefined') return;
  const target = document.getElementById('listing-camera');
  if(target) target.classList.remove('is-active');
  try{
    Quagga.stop();
  }catch(e){
    // ignore
  }
}

async function loadListingData(){
  if(listingAllDataCache) return listingAllDataCache;
  const sources = [
    { name: 'Rumunia - Słodycze', type: 'json', url: jsonUrl },
    { name: 'Rumunia - Mięso i wędliny', type: 'json', url: jsonUrlMieso },
    { name: 'Rumunia - Nabiał', type: 'json', url: jsonUrlNabial },
    { name: 'Rumunia - Napoje', type: 'json', url: jsonUrlNapoje },
    { name: 'Rumunia - Przyprawy', type: 'json', url: jsonUrlPrzyprawyProszek },
    { name: 'Rumunia - Puszki i słoiki', type: 'json', url: jsonUrlPuszkiSloiki },
    { name: 'Rumunia - Produkty podstawowe', type: 'json', url: jsonUrlProduktyPodstawowe },
    { name: 'Rumunia - Kawy i herbaty', type: 'json', url: jsonUrlKawyHerbaty },
    { name: 'Rumunia - Top Rumunia', type: 'json', url: jsonUrlTopRumunia },
    { name: 'Ukraina - Słodycze', type: 'json', url: jsonUrlSlodyczeUkraina },
    { name: 'Ukraina - Mięso i wędliny', type: 'json', url: jsonUrlMiesoUkraina },
    { name: 'Ukraina - Kawy i herbaty', type: 'json', url: jsonUrlKawyUkraina },
    { name: 'Ukraina - Puszki i słoiki', type: 'xlsx', url: xlsxUrlPuszkiUkraina },
    { name: 'Ukraina - Napoje', type: 'json', url: jsonUrlNapojeUkraina },
    { name: 'Ukraina - Przyprawy', type: 'json', url: jsonUrlPrzyprawyUkraina },
    { name: 'Ukraina - Top Ukraina', type: 'json', url: jsonUrlTopUkraina }
  ];
  if(isConfiguredDataSource(ukrainaDodatkiDoPotrawSource)) sources.push(ukrainaDodatkiDoPotrawSource);

  const results = [];
  for(const src of sources){
    try{
      const data = await loadDataSourceRows(src);
      results.push({ ...src, data });
    }catch(e){
      console.error('Listing load error', src.name, e);
    }
  }
  listingAllDataCache = results;
  return results;
}

async function searchListingByCode(code, fromScan){
  const normalized = normalizeEanForBarcode(code);
  const searchCode = normalized ? normalized.code : code.replace(/\D/g, '');
  if(!searchCode){
    listingResults = [];
    listingResultsMap.clear();
    renderListingTable();
    return;
  }
  const datasets = await loadListingData();
  const matches = [];
  for(const ds of datasets){
    const cols = Object.keys(ds.data[0] || {});
    const eanKey = findColumn(cols, ['kod ean', 'ean']);
    const indexKey = findColumn(cols, ['indeks', 'index', 'id', 'numer katalogowy']);
    const nameKey = findColumn(cols, ['nazwa', 'name']);
    const producerKey = findColumn(cols, ['producent', 'producer']);
    const groupKey = findColumn(cols, ['grupa', 'group']);
    const rankingKey = findColumn(cols, ['ranking']);
    for(const row of ds.data){
      const eanVal = eanKey ? String(row[eanKey] ?? '').replace(/\D/g, '') : '';
      const idxVal = indexKey ? String(row[indexKey] ?? '').replace(/\D/g, '') : '';
      if(eanVal === searchCode || idxVal === searchCode){
        const imageUrl = await buildListingImageUrl(ds.name, idxVal || row[indexKey], row);
        matches.push({
          _source: ds.name,
          Indeks: indexKey ? row[indexKey] ?? '' : '',
          Nazwa: nameKey ? row[nameKey] ?? '' : '',
          Producent: producerKey ? row[producerKey] ?? '' : '',
          Grupa: groupKey ? row[groupKey] ?? '' : '',
          Ranking: rankingKey ? row[rankingKey] ?? '' : '',
          'Kod EAN': eanKey ? row[eanKey] ?? '' : '',
          Zdjęcie: imageUrl || ''
        });
      }
    }
  }
  if(matches.length){
    if(fromScan){
      if(listingScannedCodes.has(searchCode)){
        pendingListingAdd = { code: searchCode, matches };
        if(listingConfirmModal) listingConfirmModal.classList.remove('hidden');
      }else{
        showListingAddPrompt(matches[0]);
        pendingMaspoAdd = { code: searchCode, matches };
        if(listingAddModal) listingAddModal.classList.remove('hidden');
      }
    }else{
      if(listingResultsMap.size){
        applyListingAdd(searchCode, matches, false);
      }else{
        listingResults = matches;
        renderListingTable();
      }
    }
    return;
  }

  const marionMatches = await searchMarionDatabase(searchCode);
  if(fromScan){
    if(marionMatches.length){
      if(listingScannedCodes.has(searchCode)){
        pendingListingAdd = { code: searchCode, matches: marionMatches };
        if(listingConfirmModal) listingConfirmModal.classList.remove('hidden');
      }else{
        showListingMarionPrompt(marionMatches[0]);
        pendingMarionAdd = { code: searchCode, matches: marionMatches };
        if(listingMarionModal) listingMarionModal.classList.remove('hidden');
      }
    }
  }else{
    listingResults = marionMatches;
    renderListingTable();
  }
}

function maybeAddListingResult(code, matches){
  if(!matches.length){
    return;
  }
  showListingFoundInfo(matches[0]);
  if(listingScannedCodes.has(code)){
    pendingListingAdd = { code, matches };
    if(listingConfirmModal) listingConfirmModal.classList.remove('hidden');
    return;
  }
  applyListingAdd(code, matches, false);
}

function showListingAddPrompt(item){
  if(!listingAddModal) return;
  if(listingAddName) listingAddName.textContent = item?.Nazwa || 'Produkt';
  if(listingAddImage){
    listingAddImage.src = item?.Zdjęcie || '';
    listingAddImage.alt = item?.Nazwa || '';
  }
}

function showListingMarionPrompt(item){
  if(!listingMarionModal) return;
  if(listingMarionName) listingMarionName.textContent = item?.Nazwa || 'Produkt';
  if(listingMarionImage){
    listingMarionImage.src = item?.Zdjęcie || '';
    listingMarionImage.alt = item?.Nazwa || '';
  }
}

function applyListingAdd(code, matches, force){
  if(!force && listingScannedCodes.has(code)) return;
  listingScannedCodes.add(code);
  matches.forEach(item => {
    const key = `${item._source || ''}|${item.Indeks || ''}|${item['Kod EAN'] || ''}|${item.Nazwa || ''}`;
    if(!listingResultsMap.has(key) || force){
      listingResultsMap.set(key, item);
    }
  });
  listingResults = Array.from(listingResultsMap.values());
  renderListingTable();
  showListingFoundToast(matches[0]);
}

function showListingFoundToast(item){
  const name = item?.Nazwa || 'Produkt';
  if(listingFoundText) listingFoundText.textContent = `Znaleziono: ${name}`;
  if(listingFoundModal) listingFoundModal.classList.remove('hidden');
}

function renderListingTable(){
  if(!listingHead || !listingBody){
    return;
  }
  if(!listingResults.length){
    listingHead.innerHTML = '<th>Brak wyników</th>';
    listingBody.innerHTML = '';
    return;
  }
  const preferred = ['Zdjęcie','Nazwa','Indeks','SKU Number','Price','Producent','Grupa','Ranking','Kod EAN'];
  const cols = preferred.filter(c => Object.prototype.hasOwnProperty.call(listingResults[0], c));
  listingHead.innerHTML = cols.map(c => `<th class="${getListingColumnClass(c)}">${escapeHtml(c)}</th>`).join('');
  listingBody.innerHTML = listingResults.map(r => `
    <tr>
      ${cols.map(c => c === 'Zdjęcie'
        ? `<td class="${getListingColumnClass(c)}">${r[c] ? buildDeferredImageTag(String(r[c]), 'listing-img') : ''}</td>`
        : `<td class="${getListingColumnClass(c)}">${escapeHtml(r[c] ?? '')}</td>`).join('')}
    </tr>
  `).join('');
  registerLazyImages(listingBody);
}

function getListingColumnClass(columnName){
  const normalized = String(columnName || '').trim().toLowerCase();
  if(normalized === 'zdjęcie') return 'listing-col listing-col--image';
  if(normalized === 'indeks' || normalized === 'sku number') return 'listing-col listing-col--index';
  if(normalized === 'nazwa') return 'listing-col listing-col--name';
  if(normalized === 'ranking') return 'listing-col listing-col--ranking';
  if(normalized === 'grupa') return 'listing-col listing-col--group';
  if(normalized === 'producent') return 'listing-col listing-col--producer';
  if(normalized === 'kod ean') return 'listing-col listing-col--ean';
  if(normalized === 'price') return 'listing-col listing-col--price';
  return 'listing-col';
}

function exportListingXlsx(){
  if(!listingResults.length || typeof XLSX === 'undefined') return;
  const ws = XLSX.utils.json_to_sheet(listingResults);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Listing');
  XLSX.writeFile(wb, 'listing.xlsx');
}

async function exportListingPdf(){
  if(!listingResults.length || !window.jspdf) return;
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({ orientation: 'p', unit: 'pt', format: 'a4' });
  const margin = 36;
  const pageW = pdf.internal.pageSize.getWidth();
  const pageH = pdf.internal.pageSize.getHeight();
  let y = margin;

  pdf.setFont('helvetica', 'bold');
  pdf.setFontSize(14);
  pdf.text('Listing - produkty', margin, y);
  y += 18;

  const rows = listingResults;
  for(const item of rows){
    if(y + 70 > pageH - margin){
      pdf.addPage();
      y = margin;
    }

    let imgData = null;
    if(item['Zdjęcie']){
      imgData = await loadImageAsDataUrl(item['Zdjęcie']);
    }
    if(imgData?.dataUrl){
      pdf.addImage(imgData.dataUrl, imgData.format, margin, y, 48, 48);
    }

    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(11);
    pdf.text(String(item['Nazwa'] || item['Name'] || '').slice(0, 80), margin + 60, y + 14);

    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(9);
    const ean = item['Kod EAN'] || '';
    const idx = item['Indeks'] || item['SKU Number'] || '';
    const price = item['Price'] || '';
    pdf.text(`Indeks/SKU: ${idx}`, margin + 60, y + 30);
    pdf.text(`EAN: ${ean}`, margin + 60, y + 44);
    if(price) pdf.text(`Cena: ${price}`, margin + 60, y + 58);

    y += 70;
  }

  pdf.save('listing.pdf');
}

async function loadImageAsDataUrl(url){
  try{
    const res = await fetch(url);
    if(!res.ok) return null;
    const blob = await res.blob();
    const objectUrl = URL.createObjectURL(blob);
    const img = new Image();
    img.crossOrigin = 'anonymous';
    try{
      await new Promise((resolve, reject) => {
        img.onload = resolve;
        img.onerror = reject;
        img.src = objectUrl;
      });
      const canvas = document.createElement('canvas');
      canvas.width = img.naturalWidth || img.width || 1;
      canvas.height = img.naturalHeight || img.height || 1;
      const ctx = canvas.getContext('2d');
      const type = String(blob.type || '').toLowerCase();
      const useJpeg = type.includes('jpg') || type.includes('jpeg');
      if(useJpeg){
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
      }else{
        ctx.clearRect(0, 0, canvas.width, canvas.height);
      }
      ctx.drawImage(img, 0, 0);
      return {
        dataUrl: useJpeg ? canvas.toDataURL('image/jpeg', 0.92) : canvas.toDataURL('image/png'),
        format: useJpeg ? 'JPEG' : 'PNG'
      };
    }finally{
      URL.revokeObjectURL(objectUrl);
    }
  }catch(e){
    return null;
  }
}


async function buildListingImageUrl(source, indexValue, row){
  const idx = String(indexValue ?? '').replace(/\D/g, '');
  if(!idx) return '';
  const isUkraina = String(source).toLowerCase().includes('ukraina');
  const base = isUkraina ? imageBaseUrlUkraina : imageBaseUrlRumunia;
  try{
    return await resolveImageAssetUrl(idx, base);
  }catch(error){
    console.error('Listing image resolve error', error);
    return '';
  }
}

async function loadMarionDatabase(){
  if(listingMarionCache) return listingMarionCache;
  try{
    const res = await fetch(jsonUrlMarion);
    const data = await res.json();
    listingMarionCache = Array.isArray(data) ? data : [];
  }catch(e){
    console.error('Marion load error', e);
    listingMarionCache = [];
  }
  return listingMarionCache;
}

async function searchMarionDatabase(code){
  const data = await loadMarionDatabase();
  const matches = [];
  data.forEach(item => {
    const ean1 = String(item['EAN 1'] ?? '').replace(/\D/g, '');
    const ean2 = String(item['EAN 2'] ?? '').replace(/\D/g, '');
    const sku = String(item['SKU Number'] ?? '').replace(/\D/g, '');
    if(ean1 === code || ean2 === code || sku === code){
      const imgUrl = item[''] || item['Image'] || item['image'] || item['Photo'] || '';
      matches.push({
        _source: 'Marion',
        Nazwa: item['Name'] ?? '',
        'SKU Number': item['SKU Number'] ?? '',
        Price: item['Price'] ?? '',
        'Kod EAN': ean1 === code ? item['EAN 1'] ?? '' : item['EAN 2'] ?? item['EAN 1'] ?? '',
        Zdjęcie: imgUrl
      });
    }
  });
  return matches;
}

function resetFilters(){
  filters = { producer:'', group:'', name:'', ean:'', index:'', limit:'' };
}

function setFullData(data, columns){
  fullData = Array.isArray(data) ? data : [];
  fullDataColumns = Array.isArray(columns) && columns.length ? columns : null;
}

function getCurrentColumns(){
  const cols = fullDataColumns && fullDataColumns.length
    ? fullDataColumns
    : (fullData && fullData.length ? Object.keys(fullData[0]) : []);
  if(!cols.length) return [];
  const seen = new Set();
  const out = [];
  cols.forEach(col => {
    const key = String(col || '')
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .trim();
    if(!key || seen.has(key)) return;
    seen.add(key);
    out.push(col);
  });
  return out;
}

function getRowImageBase(row){
  if(currentCategorySlug && currentCategorySlug.includes('ukraina')) return imageBaseUrlUkraina;
  if(currentCategorySlug && currentCategorySlug.includes('marka_wlasna')) return imageBaseUrlMarkaWlasna;
  if(!row) return currentImageBaseUrl;
  const country = String(row.__country || row.__source || '').toLowerCase();
  if(country.includes('ukraina')) return imageBaseUrlUkraina;
  if(country.includes('marka wlasna') || country.includes('marka_wlasna')) return imageBaseUrlMarkaWlasna;
  return currentImageBaseUrl || imageBaseUrlRumunia;
}

function normalizeRowIndex(row){
  if(!row || typeof row !== 'object') return;
  if(row.INDEKS !== undefined && row.INDEKS !== null && String(row.INDEKS).trim() !== '') return;
  const keys = Object.keys(row);
  const key = findColumn(keys, ['indeks', 'index', 'id', 'numer katalogowy', 'sku number']);
  if(key){
    row.INDEKS = row[key];
  }
}

function getValueByVariants(row, variants){
  if(!row) return '';
  const keys = Object.keys(row);
  const key = findColumn(keys, variants);
  if(key) return row[key];
  for(const v of variants){
    if(row[v] !== undefined) return row[v];
  }
  return '';
}

async function loadImportDane(){
  if(importDaneCache) return importDaneCache;
  const sources = [
    { name: 'Rumunia - Słodycze', type: 'json', url: jsonUrl },
    { name: 'Rumunia - Mięso i wędliny', type: 'json', url: jsonUrlMieso },
    { name: 'Rumunia - Nabiał', type: 'json', url: jsonUrlNabial },
    { name: 'Rumunia - Napoje', type: 'json', url: jsonUrlNapoje },
    { name: 'Rumunia - Przyprawy', type: 'json', url: jsonUrlPrzyprawyProszek },
    { name: 'Rumunia - Puszki i słoiki', type: 'json', url: jsonUrlPuszkiSloiki },
    { name: 'Rumunia - Produkty podstawowe', type: 'json', url: jsonUrlProduktyPodstawowe },
    { name: 'Rumunia - Kawy i herbaty', type: 'json', url: jsonUrlKawyHerbaty },
    { name: 'Rumunia - Top Rumunia', type: 'json', url: jsonUrlTopRumunia },
    { name: 'Ukraina - Słodycze', type: 'json', url: jsonUrlSlodyczeUkraina },
    { name: 'Ukraina - Mięso i wędliny', type: 'json', url: jsonUrlMiesoUkraina },
    { name: 'Ukraina - Kawy i herbaty', type: 'json', url: jsonUrlKawyUkraina },
    { name: 'Ukraina - Puszki i słoiki', type: 'xlsx', url: xlsxUrlPuszkiUkraina },
    { name: 'Ukraina - Napoje', type: 'json', url: jsonUrlNapojeUkraina },
    { name: 'Ukraina - Przyprawy', type: 'json', url: jsonUrlPrzyprawyUkraina },
    { name: 'Ukraina - Top Ukraina', type: 'json', url: jsonUrlTopUkraina }
  ];
  if(isConfiguredDataSource(ukrainaDodatkiDoPotrawSource)) sources.push(ukrainaDodatkiDoPotrawSource);

  const allRows = [];
  const columns = ['INDEKS', 'NAZWA', 'RANKING', 'GRUPA', 'PRODUCENT', 'KOD EAN'];

  for(const src of sources){
    try{
      const data = await loadDataSourceRows(src);
      (data || []).forEach(row => {
        const item = {
          INDEKS: String(getValueByVariants(row, ['indeks', 'index', 'id', 'numer katalogowy', 'sku number']) || '').trim(),
          NAZWA: String(getValueByVariants(row, ['nazwa', 'name']) || '').trim(),
          RANKING: getValueByVariants(row, ['ranking']),
          GRUPA: getValueByVariants(row, ['grupa', 'group']),
          PRODUCENT: getValueByVariants(row, ['producent', 'producer']),
          'KOD EAN': getValueByVariants(row, ['kod ean', 'ean'])
        };
        item.__source = src.name;
        item.__country = src.name.toLowerCase().includes('ukraina') ? 'Ukraina' : 'Rumunia';
        allRows.push(item);
      });
    }catch(e){
      console.error('Import dane load error', src.name, e);
    }
  }

  importDaneCache = { data: allRows, columns };
  return importDaneCache;
}

function exportCSV(){
  const cols = getCurrentColumns();
  let csv = cols.join(',') + '\n';

  const selected = selectedProducts.size ? Array.from(selectedProducts.values()) : getFilteredData();
  selected.forEach(r => {
    csv += cols
      .map(c => `"${(r[c] ?? '').toString().replace(/"/g, '""')}"`)
      .join(',') + '\n';
  });

  download(csv, `${currentCategorySlug}_rumunia.csv`, 'text/csv');
}

function exportXLS(){
  const cols = getCurrentColumns();
  const selected = selectedProducts.size ? Array.from(selectedProducts.values()) : getFilteredData();
  const rows = [cols, ...selected.map(r => cols.map(c => r[c] ?? ''))];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Dane');
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  downloadBlob(blob, `${currentCategorySlug}_rumunia.xlsx`);
}

function clearCatalogPreview(){
  if(!catalogBlobUrl) return;
  try{
    URL.revokeObjectURL(catalogBlobUrl);
  }catch(error){
    console.warn('Catalog blob revoke failed', error);
  }
  catalogBlobUrl = null;
}

function clearCatalogData(){
  selectedProducts.clear();
  catalogLoading = false;
  clearCatalogPreview();
  catalogCoverDataUrl = null;
  catalogPriceMap = null;
  catalogPriceRows = null;
  catalogOptionsOverride = null;

  if(catalogNameInput) catalogNameInput.value = '';
  if(catalogCoverInput) catalogCoverInput.value = '';
  if(catalogCurrencySelect) catalogCurrencySelect.value = '';
  if(catalogPriceColorInput) catalogPriceColorInput.value = '#000000';
  if(catalogExcelInput) catalogExcelInput.value = '';
  if(catalogIndexColSelect){
    catalogIndexColSelect.innerHTML = '<option value="auto">Auto</option>';
    catalogIndexColSelect.value = 'auto';
  }
  if(catalogPriceColSelect){
    catalogPriceColSelect.innerHTML = '<option value="auto">Auto</option>';
    catalogPriceColSelect.value = 'auto';
  }
  if(catalogError) catalogError.textContent = '';

  closeCatalogModal();
  viewMode = 'table';
  render();
}

async function createCatalog(){
  viewMode = 'catalog';
  catalogLoading = true;
  clearCatalogPreview();
  render();
  try{
    const products = Array.from(selectedProducts.values());
    if(!products.length){
      catalogLoading = false;
      render();
      return;
    }
    if(typeof window.createCatalogPdf !== 'function'){
      catalogLoading = false;
      render();
      return;
    }
    window.imageBaseUrl = currentImageBaseUrl;
    const blob = await window.createCatalogPdf(products, catalogOptionsOverride || {});
    catalogBlobUrl = URL.createObjectURL(blob);
  }catch(e){
    console.error(e);
  }finally{
    catalogLoading = false;
    render();
  }
}

function downloadCatalog(){
  if(!catalogBlobUrl) return;
  openCatalogModal();
}

function clearSelected(){
  selectedProducts.clear();
  updateTable();
}

function selectAllVisible(){
  const filteredData = getFilteredData();
  const limitValue = getLimitValue();
  const data = expanded ? filteredData : filteredData.slice(0, limitValue);
  const cols = getCurrentColumns();
  const indexKey = findColumn(cols, ['indeks', 'index', 'id', 'numer katalogowy', 'sku number']);
  if(!indexKey) return;
  data.forEach(r => {
    const idx = String(r[indexKey] ?? '').trim();
    if(idx) selectedProducts.set(idx, r);
  });
  updateTable();
}

function renderHeaderCell(col, indexKey){
  if(indexKey && col === indexKey){
    const allSelected = areAllVisibleSelected();
    const checked = allSelected ? 'checked' : '';
    return `<th class="catalog-col catalog-col--select">
              <label class="select-box select-box--header">
                <input type="checkbox" onclick="toggleSelectAll(this)" ${checked}>
                <span></span>
              </label>
            </th>
            <th class="catalog-col catalog-col--lp">Lp.</th>
            <th class="catalog-col catalog-col--image">Zdjęcie</th>
            <th class="catalog-col catalog-col--index">${escapeHtml(col)}</th>`;
  }
  return `<th class="${getCatalogColumnClass(col)}">${escapeHtml(col)}</th>`;
}

function toggleSelectAll(checkbox){
  if(checkbox.checked){
    selectAllVisible();
  }else{
    clearSelected();
  }
}

function areAllVisibleSelected(){
  const filteredData = getFilteredData();
  const limitValue = getLimitValue();
  const data = expanded ? filteredData : filteredData.slice(0, limitValue);
  const cols = getCurrentColumns();
  const indexKey = findColumn(cols, ['indeks', 'index', 'id', 'numer katalogowy', 'sku number']);
  if(!indexKey || data.length === 0) return false;
  return data.every(r => {
    const idx = String(r[indexKey] ?? '').trim();
    return idx && selectedProducts.has(idx);
  });
}

function openCatalogModal(){
  if(!catalogModal) return;
  if(catalogNameInput) catalogNameInput.value = `katalog_${currentCategorySlug}`;
  if(catalogCoverInput) catalogCoverInput.value = '';
  if(catalogExcelInput) catalogExcelInput.value = '';
  resetCatalogColumnOptions();
  applyAdminCatalogPriceMapToCatalogState();
  if(catalogError) catalogError.textContent = '';
  catalogModal.classList.remove('hidden');
  if(isCurrentUserAdmin()){
    void ensureAdminCatalogPriceCache().then(() => {
      if(catalogModal && !catalogModal.classList.contains('hidden')){
        applyAdminCatalogPriceMapToCatalogState();
      }
    });
  }
}

function closeCatalogModal(){
  if(catalogModal) catalogModal.classList.add('hidden');
}

async function buildAndSaveCatalog(){
  const name = (catalogNameInput?.value || '').trim() || `katalog_${currentCategorySlug}`;
  if(catalogError) catalogError.textContent = '';
  if(catalogSaveBtn) catalogSaveBtn.disabled = true;
  try{
    const products = Array.from(selectedProducts.values());
    if(!products.length){
      if(catalogError) catalogError.textContent = 'Najpierw zaznacz produkty.';
      return;
    }
    await ensurePasswordIfNeeded(products.length);
    if(isCurrentUserAdmin() && !catalogPriceMap && !catalogExcelInput?.files?.length){
      await ensureAdminCatalogPriceCache();
      applyAdminCatalogPriceMapToCatalogState();
    }
    const currency = catalogCurrencySelect?.value || '';
    if(catalogPriceMap && !currency){
      if(catalogError) catalogError.textContent = 'Wybierz walutę przed zapisem katalogu.';
      return;
    }
    const priceColor = catalogPriceColorInput?.value || '#000000';
    const options = {
      ...(catalogOptionsOverride || {}),
      coverDataUrl: catalogCoverDataUrl || catalogOptionsOverride?.coverDataUrl || null,
      priceMap: catalogPriceMap,
      currency,
      priceColor
    };
    window.imageBaseUrl = currentImageBaseUrl;
    const blob = await window.createCatalogPdf(products, options);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${name}.pdf`;
    a.click();
    closeCatalogModal();
  }catch(e){
    console.error(e);
    if(catalogError) catalogError.textContent = 'Nie udało się zapisać katalogu.';
  }finally{
    if(catalogSaveBtn) catalogSaveBtn.disabled = false;
  }
}

function fileToDataUrl(file){
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function normalizeCatalogHeader(value){
  return String(value ?? '')
    .toLowerCase()
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function findCatalogColumn(header, candidates, fallback = -1){
  const normalizedCandidates = (Array.isArray(candidates) ? candidates : [])
    .map(normalizeCatalogHeader)
    .filter(Boolean);

  const exactIndex = header.findIndex(label => normalizedCandidates.includes(label));
  if(exactIndex !== -1) return exactIndex;

  const includesIndex = header.findIndex(label => (
    normalizedCandidates.some(candidate => label.includes(candidate))
  ));
  return includesIndex !== -1 ? includesIndex : fallback;
}

function parseCatalogNumericValue(value){
  if(typeof value === 'number' && Number.isFinite(value)) return value;
  const normalized = String(value ?? '')
    .replace(/\s+/g, '')
    .replace(/[^0-9,.-]/g, '')
    .replace(',', '.');
  const numeric = Number(normalized);
  return Number.isFinite(numeric) ? numeric : null;
}

function formatCatalogNumericValue(value, maxFractionDigits = 3){
  const numeric = typeof value === 'number' && Number.isFinite(value)
    ? value
    : parseCatalogNumericValue(value);
  if(!Number.isFinite(numeric)) return String(value ?? '').trim();

  return numeric.toLocaleString('en-GB', {
    minimumFractionDigits: 0,
    maximumFractionDigits: maxFractionDigits,
    useGrouping: false
  });
}

function getCatalogCurrencySymbol(value){
  const text = String(value ?? '');
  if(text.includes('£')) return '£';
  if(text.includes('€')) return '€';
  if(text.includes('$')) return '$';
  return '';
}

function formatCatalogPriceValue(value){
  const numeric = parseCatalogNumericValue(value);
  const currencySymbol = getCatalogCurrencySymbol(value);
  if(!Number.isFinite(numeric)){
    return String(value ?? '').trim();
  }

  const price = numeric.toFixed(2);
  return currencySymbol ? `${currencySymbol} ${price}` : price;
}

function buildPriceMap(rows, config){
  if(!rows || !rows.length) return null;
  const headerRow = rows[0] || [];
  const header = headerRow.map(normalizeCatalogHeader);
  const hasHeader =
    header.some(h => h.includes('indeks')) ||
    header.some(h => h.includes('index')) ||
    header.some(h => h.includes('cena')) ||
    header.some(h => h.includes('price')) ||
    header.some(h => h.includes('netto-cell'));

  const idxCol = typeof config?.indexCol === 'number'
    ? config.indexCol
    : (hasHeader ? findCatalogColumn(header, ['indeks', 'index', 'index-cell'], 0) : 0);
  const priceCol = typeof config?.priceCol === 'number'
    ? config.priceCol
    : (hasHeader ? findCatalogColumn(header, ['cena', 'price', 'netto', 'netto-cell'], 1) : 1);
  const unitCol = hasHeader
    ? findCatalogColumn(header, ['jednostka', 'unit', 'package-cell 2'], -1)
    : -1;
  const minimumQtyCol = hasHeader
    ? findCatalogColumn(header, ['minimum', 'min zamowienie', 'minimalne zamowienie', 'ilosc minimalna', 'ilość minimalna', 'package-cell'], -1)
    : -1;
  const nameCol = hasHeader
    ? findCatalogColumn(header, ['nazwa', 'name', 'product', 'text-decoration-none'], -1)
    : -1;
  const eanCol = hasHeader
    ? findCatalogColumn(header, ['ean', 'kod ean', 'barcode', 'text-right'], -1)
    : -1;
  const producerCol = hasHeader
    ? findCatalogColumn(header, ['producent', 'producer', 'text-left'], -1)
    : -1;

  if(idxCol === -1 || priceCol === -1) return null;
  const map = new Map();
  const startRow = hasHeader ? 1 : 0;

  for(let i = startRow; i < rows.length; i++){
    const r = rows[i];
    const indexRaw = r[idxCol];
    const index = normalizeIndexValue(indexRaw);
    if(!index) continue;
    const priceRaw = r[priceCol];
    const price = formatCatalogPriceValue(priceRaw);
    const priceNumeric = parseCatalogNumericValue(priceRaw);
    const unit = unitCol !== -1 ? String(r[unitCol] || '').trim() : '';
    const minimumOrderQuantityRaw = minimumQtyCol !== -1 ? r[minimumQtyCol] : '';
    const minimumOrderQuantityNumeric = parseCatalogNumericValue(minimumOrderQuantityRaw);
    const minimumOrderQuantity = Number.isFinite(minimumOrderQuantityNumeric)
      ? formatCatalogNumericValue(minimumOrderQuantityNumeric)
      : String(minimumOrderQuantityRaw ?? '').trim();
    const minimumOrderUnit = unit;
    const minimumOrderValue = Number.isFinite(priceNumeric) && Number.isFinite(minimumOrderQuantityNumeric)
      ? (priceNumeric * minimumOrderQuantityNumeric).toFixed(2)
      : '';
    const currencySymbol = getCatalogCurrencySymbol(priceRaw);

    map.set(index, {
      price,
      unit,
      priceNumeric: Number.isFinite(priceNumeric) ? priceNumeric.toFixed(4) : '',
      currencySymbol,
      minimumOrderQuantity,
      minimumOrderUnit,
      minimumOrderValue,
      name: nameCol !== -1 ? String(r[nameCol] || '').trim() : '',
      ean: eanCol !== -1 ? String(r[eanCol] || '').trim() : '',
      producer: producerCol !== -1 ? String(r[producerCol] || '').trim() : ''
    });
  }
  return map;
}

function getCatalogColumnOptionItems(rows){
  const normalizedRows = Array.isArray(rows) ? rows : [];
  const colCount = normalizedRows.reduce((maxColumns, row) => (
    Math.max(maxColumns, Array.isArray(row) ? row.length : 0)
  ), 0);
  const header = Array.isArray(normalizedRows[0]) ? normalizedRows[0] : [];

  return Array.from({ length: colCount }, (_, index) => {
    const headerLabel = header[index] ? ` (${header[index]})` : '';
    return {
      value: String(index),
      label: `${columnLabel(index)}${headerLabel}`
    };
  });
}

function resetCatalogColumnOptions(){
  if(catalogIndexColSelect){
    catalogIndexColSelect.innerHTML = '<option value="auto">Auto</option>';
    catalogIndexColSelect.value = 'auto';
  }
  if(catalogPriceColSelect){
    catalogPriceColSelect.innerHTML = '<option value="auto">Auto</option>';
    catalogPriceColSelect.value = 'auto';
  }
}

function setCatalogColumnSelectValue(select, value = 'auto'){
  if(!select) return;
  const normalizedValue = String(value ?? 'auto');
  const hasMatchingOption = Array.from(select.options).some(option => option.value === normalizedValue);
  select.value = hasMatchingOption ? normalizedValue : 'auto';
}

function getCatalogPriceImportState(){
  const options = Array.isArray(catalogPriceRows) && catalogPriceRows.length
    ? getCatalogColumnOptionItems(catalogPriceRows)
    : [];

  let source = 'none';
  if(Array.isArray(catalogPriceRows) && catalogPriceRows.length){
    source = catalogPriceRows === adminCatalogPriceRowsCache ? 'saved-admin' : 'file';
  }else if(adminCatalogPriceMapCache instanceof Map && adminCatalogPriceMapCache.size){
    source = 'saved-admin';
  }

  return {
    available: typeof XLSX !== 'undefined',
    hasPriceMap: catalogPriceMap instanceof Map && catalogPriceMap.size > 0,
    itemsCount: catalogPriceMap instanceof Map ? catalogPriceMap.size : 0,
    options,
    indexCol: String(catalogIndexColSelect?.value || 'auto'),
    priceCol: String(catalogPriceColSelect?.value || 'auto'),
    source
  };
}

function applyCatalogPriceRows(rows, options = {}){
  const normalizedRows = Array.isArray(rows) ? rows : [];
  catalogPriceRows = normalizedRows;

  if(!normalizedRows.length){
    catalogPriceMap = null;
    resetCatalogColumnOptions();
    return getCatalogPriceImportState();
  }

  populateCatalogColumnOptions(normalizedRows);
  setCatalogColumnSelectValue(catalogIndexColSelect, options.indexCol ?? 'auto');
  setCatalogColumnSelectValue(catalogPriceColSelect, options.priceCol ?? 'auto');
  catalogPriceMap = buildPriceMap(normalizedRows, getCatalogColumnConfig());
  return getCatalogPriceImportState();
}

async function importCatalogPriceFileForGenerator(file){
  if(!file){
    throw new Error('Wybierz plik Excel z cennikiem.');
  }
  if(typeof XLSX === 'undefined'){
    throw new Error('Moduł Excel nie jest jeszcze gotowy. Odśwież stronę i spróbuj ponownie.');
  }

  const buf = await file.arrayBuffer();
  const workbook = XLSX.read(buf, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const sheet = sheetName ? workbook.Sheets[sheetName] : null;
  const rows = sheet
    ? XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' })
    : [];

  return applyCatalogPriceRows(rows, {
    indexCol: 'auto',
    priceCol: 'auto'
  });
}

function updateCatalogPriceColumnsForGenerator(indexCol = 'auto', priceCol = 'auto'){
  setCatalogColumnSelectValue(catalogIndexColSelect, indexCol);
  setCatalogColumnSelectValue(catalogPriceColSelect, priceCol);

  if(Array.isArray(catalogPriceRows) && catalogPriceRows.length){
    catalogPriceMap = buildPriceMap(catalogPriceRows, getCatalogColumnConfig());
  }

  return getCatalogPriceImportState();
}

function applyAdminCatalogPriceMapToCatalogState(){
  if(!isCurrentUserAdmin()){
    return false;
  }

  if(
    !adminCatalogPriceMapCache
    || !(adminCatalogPriceMapCache instanceof Map)
    || !Array.isArray(adminCatalogPriceRowsCache)
    || !adminCatalogPriceRowsCache.length
  ){
    catalogPriceMap = null;
    catalogPriceRows = null;
    return false;
  }

  catalogPriceRows = adminCatalogPriceRowsCache;
  catalogPriceMap = adminCatalogPriceMapCache;
  populateCatalogColumnOptions(adminCatalogPriceRowsCache);
  if(catalogIndexColSelect) catalogIndexColSelect.value = 'auto';
  if(catalogPriceColSelect) catalogPriceColSelect.value = 'auto';
  return true;
}

async function ensureAdminCatalogPriceCache(options = {}){
  if(!isCurrentUserAdmin()){
    adminCatalogPriceMapCache = null;
    adminCatalogPriceRowsCache = null;
    adminCatalogPriceLoadPromise = null;
    return null;
  }

  if(adminCatalogPriceMapCache && adminCatalogPriceRowsCache && !options.forceRefresh){
    return adminCatalogPriceMapCache;
  }

  if(adminCatalogPriceLoadPromise && !options.forceRefresh){
    return adminCatalogPriceLoadPromise;
  }

  adminCatalogPriceLoadPromise = (async () => {
    if(!storage || typeof storage.ref !== 'function'){
      return null;
    }

    try{
      const buffer = await readStorageSourceArrayBuffer(ADMIN_PRICE_LIST_STORAGE_PATH);
      const workbook = XLSX.read(buffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = sheetName ? workbook.Sheets[sheetName] : null;
      const rows = sheet
        ? XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' })
        : [];
      const priceMap = buildPriceMap(rows, {
        indexCol: null,
        priceCol: null
      });

      if(priceMap && priceMap.size){
        adminCatalogPriceRowsCache = rows;
        adminCatalogPriceMapCache = priceMap;
        return priceMap;
      }
    }catch(error){
      const code = String(error?.code || '').trim();
      const message = String(error?.message || '').trim();
      if(code !== 'storage/object-not-found' && !message.includes('404')){
        console.error('Admin price list load error', error);
      }
    }

    adminCatalogPriceRowsCache = null;
    adminCatalogPriceMapCache = null;
    return null;
  })();

  try{
    return await adminCatalogPriceLoadPromise;
  }finally{
    adminCatalogPriceLoadPromise = null;
  }
}

function populateCatalogColumnOptions(rows){
  if(!catalogIndexColSelect || !catalogPriceColSelect) return;
  const options = getCatalogColumnOptionItems(rows)
    .map(option => `<option value="${option.value}">${option.label}</option>`)
    .join('');
  catalogIndexColSelect.innerHTML = `<option value="auto">Auto</option>${options}`;
  catalogPriceColSelect.innerHTML = `<option value="auto">Auto</option>${options}`;
}

window.catalogPriceAdminTools = {
  isAvailable: () => typeof XLSX !== 'undefined',
  getState: () => getCatalogPriceImportState(),
  loadSaved: async () => {
    if(isCurrentUserAdmin()){
      await ensureAdminCatalogPriceCache();
      applyAdminCatalogPriceMapToCatalogState();
    }
    return getCatalogPriceImportState();
  },
  importFile: importCatalogPriceFileForGenerator,
  updateColumns: updateCatalogPriceColumnsForGenerator
};

function getCatalogColumnConfig(){
  return {
    indexCol: parseColumnValue(catalogIndexColSelect?.value),
    priceCol: parseColumnValue(catalogPriceColSelect?.value)
  };
}

function parseColumnValue(value){
  if(value === undefined || value === null || value === '' || value === 'auto') return null;
  const num = Number(value);
  return Number.isFinite(num) ? num : null;
}

function columnLabel(index){
  let n = index + 1;
  let label = '';
  while(n > 0){
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

function ensurePasswordIfNeeded(count){
  return new Promise((resolve, reject) => {
    if(count <= 50) return resolve();
    if(!passwordModal) return reject();
    if(passwordInput) passwordInput.value = '';
    if(passwordError) passwordError.textContent = '';
    passwordModal.classList.remove('hidden');

    const onConfirm = () => {
      const pass = passwordInput?.value || '';
      if(pass !== 'Maspo2026'){
        if(passwordError) passwordError.textContent = 'Nieprawidłowe hasło.';
        return;
      }
      passwordModal.classList.add('hidden');
      cleanup();
      resolve();
    };

    const onCancel = () => {
      passwordModal.classList.add('hidden');
      cleanup();
      reject();
    };

    const cleanup = () => {
      passwordConfirm?.removeEventListener('click', onConfirm);
      passwordCancel?.removeEventListener('click', onCancel);
    };

    passwordConfirm?.addEventListener('click', onConfirm);
    passwordCancel?.addEventListener('click', onCancel);
  });
}

async function downloadPDF(){
  const res = await fetch(getActivePdfUrl());
  const blob = await res.blob();
  const suffix = filters.producer ? `_${slugify(filters.producer)}` : '';
  downloadBlob(blob, `${currentCategorySlug}_rumunia${suffix}.pdf`);
}

function download(content, filename, type){
  const blob = new Blob([content], { type });
  downloadBlob(blob, filename);
}

function downloadBlob(blob, filename){
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
}

async function loadXlsxAsJson(url){
  const res = await fetch(url);
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { defval: '' });
}

async function readStorageSourceArrayBuffer(path){
  const normalizedPath = String(path || '').trim().replace(/^\/+/, '');
  if(!normalizedPath){
    throw new Error('Brak ścieżki pliku źródłowego.');
  }
  if(!storage || typeof storage.ref !== 'function'){
    throw new Error('Firebase Storage nie jest gotowy.');
  }
  const reference = storage.ref().child(normalizedPath);
  return readReportEntryArrayBuffer(buildStorageReportEntry(reference, {
    fullPath: normalizedPath
  }));
}

async function loadStorageJson(path){
  const buffer = await readStorageSourceArrayBuffer(path);
  const rows = JSON.parse(new TextDecoder('utf-8').decode(new Uint8Array(buffer)));
  return Array.isArray(rows) ? rows : [];
}

async function loadStorageXlsxAsJson(path){
  const buffer = await readStorageSourceArrayBuffer(path);
  const wb = XLSX.read(buffer, { type: 'array' });
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { defval: '' });
}

function getFilteredData(){
  const cols = getCurrentColumns();
  const producerKey = findColumn(cols, ['producent', 'producer']);
  const groupKey = findColumn(cols, ['grupa', 'group']);
  const nameKey = findColumn(cols, ['nazwa', 'name']);
  const eanKey = findColumn(cols, ['kod ean', 'ean']);
  const indexKey = findColumn(cols, ['indeks', 'index', 'id', 'numer katalogowy', 'sku number']);

  const filtered = fullData.filter(r => {
    if(importedIndexSet && indexKey){
      const idx = normalizeIndexValue(r[indexKey]);
      if(!importedIndexSet.has(idx)) return false;
    }
    if(producerKey && filters.producer && String(r[producerKey] ?? '') !== filters.producer) return false;
    if(groupKey && filters.group && String(r[groupKey] ?? '') !== filters.group) return false;
    if(nameKey && filters.name && !includesText(r[nameKey], filters.name)) return false;
    if(eanKey && filters.ean && !includesText(r[eanKey], filters.ean)) return false;
    if(indexKey && filters.index && !includesText(r[indexKey], filters.index)) return false;
    return true;
  });

  if(importedIndexSet && indexKey){
    return dedupeRowsByIndex(filtered, indexKey);
  }

  return filtered;
}

function updateTable(){
  const cols = getCurrentColumns();
  const filteredData = getFilteredData();
  const limitValue = getLimitValue();
  const data = expanded ? filteredData : filteredData.slice(0, limitValue);
  const indexKey = findColumn(cols, ['indeks', 'index', 'id', 'numer katalogowy', 'sku number']);
  const eanKey = findColumn(cols, ['kod ean', 'ean']);

  const tbody = document.getElementById('data-tbody');
  if(tbody){
    tbody.innerHTML = data.map((r, rowIndex) => `
      <tr>
        ${cols.map(c => renderCell(c, r, indexKey, eanKey, rowIndex + 1)).join('')}
      </tr>
    `).join('');
    registerLazyImages(tbody);
  }

  const expandContainer = document.getElementById('expand-container');
  if(expandContainer){
    if(!expanded && !filters.limit && filteredData.length > LIMIT){
      expandContainer.innerHTML = `
        <div class="actions">
          <button onclick="expandTable()">Rozwiń dalej</button>
        </div>
      `;
    }else{
      expandContainer.innerHTML = '';
    }
  }
}

function findColumn(cols, variants){
  const lower = cols.map(c => c.toLowerCase());
  for(const v of variants){
    const i = lower.indexOf(v.toLowerCase());
    if(i !== -1) return cols[i];
  }
  return null;
}

function includesText(value, query){
  const v = String(value ?? '').toLowerCase();
  const q = String(query ?? '').toLowerCase();
  return v.includes(q);
}

function uniqueValues(data, key){
  const set = new Set();
  data.forEach(r => {
    const v = r[key];
    if(v !== undefined && v !== null && String(v).trim() !== '') set.add(String(v));
  });
  return Array.from(set).sort((a,b) => a.localeCompare(b));
}

function escapeHtml(str){
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function escapeAttr(str){
  return escapeHtml(str).replace(/"/g, '&quot;');
}

function getLimitValue(){
  const v = parseInt(filters.limit, 10);
  if(Number.isFinite(v) && v > 0) return v;
  return LIMIT;
}

function getActivePdfUrl(){
  if(filters.producer){
    const key = String(filters.producer).toLowerCase().trim();
    if(producerPdfMap[key]) return producerPdfMap[key];
  }
  return pdfUrl;
}

function slugify(value){
  return String(value)
    .toLowerCase()
    .trim()
    .replace(/\s+/g, '_')
    .replace(/[^a-z0-9_]/g, '');
}

function setActiveCard(id){
  document.querySelectorAll('.grid .card').forEach(c => c.classList.remove('active'));
  const el = document.getElementById(id);
  if(el) el.classList.add('active');
}

function renderCategoryLoadError(container, message){
  if(!container) return;
  container.innerHTML = `
    <div class="sales-insight-empty">
      ${escapeHtml(message || 'Nie udało się wczytać danych dla tej kategorii.')}
    </div>
  `;
}

function isCurrentUserAdmin(){
  return Boolean(currentUserIsAdmin);
}

function syncCatalogAdminPriceTools(isAdmin = isCurrentUserAdmin()){
  if(!catalogAdminPriceTools) return;
  catalogAdminPriceTools.classList.remove('hidden');
  if(catalogExcelInput) catalogExcelInput.disabled = false;
  if(catalogIndexColSelect) catalogIndexColSelect.disabled = false;
  if(catalogPriceColSelect) catalogPriceColSelect.disabled = false;
}

function toggleAdminPanel(email, isAdmin = false){
  const adminEnabled = Boolean(isAdmin);
  currentUserIsAdmin = adminEnabled;
  if(adminPanelBtn) adminPanelBtn.classList.toggle('hidden', !adminEnabled);
  if(reportsCard) reportsCard.classList.toggle('hidden', !adminEnabled);
  syncCatalogAdminPriceTools(adminEnabled);
  if(!adminEnabled){
    adminCatalogPriceMapCache = null;
    adminCatalogPriceRowsCache = null;
    adminCatalogPriceLoadPromise = null;
    if(reportsContainer) reportsContainer.innerHTML = '';
  }
}

function toggleProductSelection(checkbox){
  const index = checkbox.getAttribute('data-index');
  if(!index) return;
  const cols = getCurrentColumns();
  const indexKey = findColumn(cols, ['indeks', 'index', 'id', 'numer katalogowy', 'sku number']);
  if(!indexKey) return;
  const row = fullData.find(r => String(r[indexKey]).trim() === index);
  if(checkbox.checked && row){
    selectedProducts.set(index, row);
  }else{
    selectedProducts.delete(index);
  }
}

function renderCell(col, row, indexKey, eanKey, rowNumber){
  const value = row ? row[col] : '';
  if(indexKey && col === indexKey){
    const idx = String(value ?? '').trim();
    const checked = selectedProducts.has(idx) ? 'checked' : '';
    const imgBase = getRowImageBase(row);
    const img = idx
      ? `<span class="img-hover" onmouseenter="positionPopup(this)">
           ${buildProductImageTag(idx, imgBase, 'index-img')}
           <span class="img-pop">
             ${buildProductImageTag(idx, imgBase, 'index-img-large', { popup: true })}
           </span>
         </span>`
      : '';
    return `<td class="catalog-col catalog-col--select">
              <label class="select-box">
                <input type="checkbox" data-index="${escapeAttr(idx)}" onchange="toggleProductSelection(this)" ${checked}>
                <span></span>
              </label>
            </td>
            <td class="catalog-col catalog-col--lp"><span class="row-number">${rowNumber || ''}.</span></td>
            <td class="catalog-col catalog-col--image">${img}</td>
            <td class="catalog-col catalog-col--index">${escapeHtml(idx)}</td>`;
  }
  if(eanKey && col === eanKey){
    const raw = String(value ?? '').trim();
    const normalized = normalizeEanForBarcode(raw);
    if(normalized){
      const barcode = getBarcodeDataUrl(normalized);
      return `<td class="catalog-col catalog-col--ean ean-cell">
                ${barcode ? `<img class="ean-barcode" src="${barcode}" alt="${escapeAttr(normalized.code)}">` : ''}
                <div class="ean-text">${escapeHtml(normalized.code)}</div>
              </td>`;
    }
  }
  return `<td class="${getCatalogColumnClass(col)}">${escapeHtml(value ?? '')}</td>`;
}

function getCatalogColumnClass(columnName){
  const normalized = String(columnName || '').trim().toLowerCase();
  if(normalized === 'nazwa') return 'catalog-col catalog-col--name';
  if(normalized === 'ranking') return 'catalog-col catalog-col--ranking';
  if(normalized === 'grupa' || normalized === 'grupa produktowa') return 'catalog-col catalog-col--group';
  if(normalized === 'producent' || normalized === 'skrot_producenta') return 'catalog-col catalog-col--producer';
  if(normalized === 'kod ean' || normalized === 'ean') return 'catalog-col catalog-col--ean';
  return 'catalog-col';
}

function calculateEAN13Checksum(code12){
  const digits = code12.split('').map(Number);
  let sum = 0;
  for(let i = 0; i < 12; i++){
    sum += digits[i] * (i % 2 === 0 ? 1 : 3);
  }
  return (10 - (sum % 10)) % 10;
}

function calculateEAN8Checksum(code7){
  const digits = code7.split('').map(Number);
  let sum = 0;
  for(let i = 0; i < 7; i++){
    sum += digits[i] * (i % 2 === 0 ? 3 : 1);
  }
  return (10 - (sum % 10)) % 10;
}

function normalizeEanForBarcode(raw){
  const digits = String(raw ?? '').replace(/\D/g, '');
  if(!digits) return null;

  if(digits.length <= 8){
    let code = digits;
    if(code.length < 7) code = code.padStart(7, '0');
    if(code.length >= 7){
      const base = code.slice(0, 7);
      code = base + calculateEAN8Checksum(base);
    }
    if(code.length !== 8) return null;
    return { code, format: 'EAN8' };
  }

  let code = digits;
  if(code.length < 12) code = code.padStart(12, '0');
  const base = code.slice(0, 12);
  code = base + calculateEAN13Checksum(base);
  if(code.length !== 13) return null;
  return { code, format: 'EAN13' };
}

function getBarcodeDataUrl(normalized){
  const key = `${normalized.format}:${normalized.code}`;
  if(barcodeCache.has(key)) return barcodeCache.get(key);
  if(typeof JsBarcode === 'undefined') return '';
  const canvas = document.createElement('canvas');
  try{
    JsBarcode(canvas, normalized.code, {
      format: normalized.format,
      width: 1.6,
      height: 44,
      displayValue: false,
      margin: 0
    });
    const url = canvas.toDataURL('image/png');
    barcodeCache.set(key, url);
    return url;
  }catch(e){
    console.error('Barcode error', e);
    return '';
  }
}

function buildImageUrl(index, ext, baseUrl){
  const base = baseUrl || currentImageBaseUrl;
  const normalizedIndex = normalizeIndexValue(index) || String(index ?? '').trim();
  const normalizedExt = getNormalizedImageExt(ext) || getDefaultImageExtForBase(base);
  if(isPrivateMediaBase(base)){
    const normalizedBase = String(base || '').replace(/^gs:\/\/[^/]+\//, '');
    return `${normalizedBase}${normalizedIndex}.${normalizedExt}`;
  }
  return `${base}${encodeURIComponent(normalizedIndex)}.${normalizedExt}?alt=media`;
}

function isPrivateMediaBase(baseUrl){
  return String(baseUrl || '').startsWith(MEDIA_STORAGE_GS_PREFIX);
}

function isUkrainaImageBase(baseUrl){
  const normalized = String(baseUrl || currentImageBaseUrl || '').toLowerCase();
  return normalized.includes('ukraina');
}

function getPrivateImageCacheKey(path){
  return String(path || '').trim();
}

function isRumuniaPrivateBase(baseUrl){
  const normalizedBase = String(baseUrl || '').replace(/^gs:\/\/[^/]+\//, '');
  return normalizedBase === 'World Food/Rumunia/';
}

async function loadRumuniaImageManifest(){
  if(rumuniaImageManifestPromise){
    return rumuniaImageManifestPromise;
  }

  rumuniaImageManifestPromise = fetch(rumuniaImageManifestUrl, { cache: 'no-store' })
    .then(response => {
      if(!response.ok){
        throw new Error(`Nie udało się pobrać mapy zdjęć Rumunii (HTTP ${response.status}).`);
      }
      return response.json();
    })
    .then(data => {
      if(!data || typeof data !== 'object'){
        throw new Error('Mapa zdjęć Rumunii ma nieprawidłowy format.');
      }
      return data;
    })
    .catch(error => {
      rumuniaImageManifestPromise = null;
      throw error;
    });

  return rumuniaImageManifestPromise;
}

async function getManifestImageUrl(storagePath, baseUrl){
  if(!isRumuniaPrivateBase(baseUrl)){
    return '';
  }

  const manifest = await loadRumuniaImageManifest();
  return String(manifest?.[storagePath] || '').trim();
}

async function getSignedMediaUrl(storagePath){
  const normalizedPath = String(storagePath || '').trim().replace(/^\/+/, '');
  if(!normalizedPath){
    throw new Error('Brak ścieżki obrazu.');
  }
  if(!functionsService || typeof functionsService.httpsCallable !== 'function'){
    throw new Error('Firebase Functions nie jest dostępne.');
  }

  const callable = functionsService.httpsCallable('getMediaSignedUrl');
  const result = await callable({ path: normalizedPath });
  const url = String(result?.data?.url || '').trim();
  if(!url){
    throw new Error('Nie udało się pobrać signed URL obrazu.');
  }
  return url;
}

function getStorageBucketName(reference, metadata){
  const referenceBucket = String(reference?.bucket || '').trim();
  if(referenceBucket) return referenceBucket;
  const metadataBucket = String(metadata?.bucket || '').trim();
  if(metadataBucket) return metadataBucket;
  return String(MEDIA_STORAGE_BUCKET || '').replace(/^gs:\/\//, '').replace(/\/+$/, '');
}

function extractStorageDownloadToken(metadata){
  const rawToken =
    metadata?.downloadTokens ||
    metadata?.downloadToken ||
    metadata?.customMetadata?.firebaseStorageDownloadTokens ||
    metadata?.customMetadata?.downloadTokens ||
    '';

  return String(rawToken || '')
    .split(',')
    .map(part => part.trim())
    .find(Boolean) || '';
}

function buildStorageTokenUrl(reference, metadata){
  const bucketName = getStorageBucketName(reference, metadata);
  const fullPath = String(metadata?.fullPath || reference?.fullPath || '').trim();
  const token = extractStorageDownloadToken(metadata);
  if(!bucketName || !fullPath || !token) return '';
  return `https://firebasestorage.googleapis.com/v0/b/${encodeURIComponent(bucketName)}/o/${encodeURIComponent(fullPath)}?alt=media&token=${encodeURIComponent(token)}`;
}

function getSharedMediaDownloadToken(storagePath){
  const normalizedPath = String(storagePath || '').trim().replace(/^\/+/, '');
  if(!normalizedPath) return '';
  for(const [prefix, token] of Object.entries(MEDIA_SHARED_DOWNLOAD_TOKEN_BY_PREFIX)){
    if(normalizedPath.startsWith(prefix)){
      return String(token || '').trim();
    }
  }
  return '';
}

function buildSharedMediaTokenUrl(storagePath){
  const normalizedPath = String(storagePath || '').trim().replace(/^\/+/, '');
  const token = getSharedMediaDownloadToken(normalizedPath);
  const bucketName = String(MEDIA_STORAGE_BUCKET || '').replace(/^gs:\/\//, '').replace(/\/+$/, '');
  if(!bucketName || !normalizedPath || !token) return '';
  return `https://firebasestorage.googleapis.com/v0/b/${encodeURIComponent(bucketName)}/o/${encodeURIComponent(normalizedPath)}?alt=media&token=${encodeURIComponent(token)}`;
}

async function resolveStorageReferenceUrl(reference){
  if(!reference || typeof reference.getMetadata !== 'function') return '';
  const metadata = await reference.getMetadata();
  return buildStorageTokenUrl(reference, metadata);
}

async function fetchBlobFromDownloadUrl(downloadUrl){
  const response = await fetch(downloadUrl, { cache: 'force-cache' });
  if(!response.ok){
    throw new Error(`Nie udało się pobrać obrazu (HTTP ${response.status}).`);
  }
  return response.blob();
}

async function readStorageReferenceBlob(reference){
  const errors = [];

  if(typeof reference.getBlob === 'function'){
    try{
      return await reference.getBlob();
    }catch(error){
      errors.push(error);
    }
  }

  try{
    const downloadUrl = await reference.getDownloadURL();
    try{
      return await fetchBlobFromDownloadUrl(downloadUrl);
    }catch(fetchError){
      errors.push(fetchError);
    }
  }catch(error){
    errors.push(error);
  }

  throw errors[errors.length - 1] || new Error('Nie udało się odczytać obrazu ze Storage.');
}

async function resolvePrivateImageSource(index, baseUrl, preferredExtOverride = ''){
  const extCandidates = getImageCandidateExts(index, baseUrl, preferredExtOverride);
  const normalizedBase = String(baseUrl || '').replace(/^gs:\/\/[^/]+\//, '');
  const errors = [];

  for(const ext of extCandidates){
    const storagePath = `${normalizedBase}${normalizeIndexValue(index) || String(index ?? '').trim()}.${ext}`;
    const cacheKey = getPrivateImageCacheKey(storagePath);
    if(privateImageUrlCache.has(cacheKey)){
      return {
        objectUrl: privateImageUrlCache.get(cacheKey),
        ext
      };
    }

    try{
      try{
        const manifestUrl = await getManifestImageUrl(storagePath, baseUrl);
        if(manifestUrl){
          privateImageUrlCache.set(cacheKey, manifestUrl);
          return {
            objectUrl: manifestUrl,
            ext
          };
        }
      }catch(manifestError){
        errors.push(manifestError);
      }

      const reference = mediaStorage ? mediaStorage.ref().child(storagePath) : null;

      try{
        const directUrl = reference ? await resolveStorageReferenceUrl(reference) : '';
        if(directUrl){
          privateImageUrlCache.set(cacheKey, directUrl);
          return {
            objectUrl: directUrl,
            ext
          };
        }
      }catch(metadataError){
        errors.push(metadataError);
      }

      try{
        const sharedUrl = buildSharedMediaTokenUrl(storagePath);
        if(sharedUrl){
          privateImageUrlCache.set(cacheKey, sharedUrl);
          return {
            objectUrl: sharedUrl,
            ext
          };
        }
      }catch(sharedUrlError){
        errors.push(sharedUrlError);
      }

      try{
        const signedUrl = await getSignedMediaUrl(storagePath);
        if(signedUrl){
          privateImageUrlCache.set(cacheKey, signedUrl);
          return {
            objectUrl: signedUrl,
            ext
          };
        }
      }catch(signedUrlError){
        errors.push(signedUrlError);
      }
      if(reference){
        const blob = await readStorageReferenceBlob(reference);
        const objectUrl = URL.createObjectURL(blob);
        privateImageUrlCache.set(cacheKey, objectUrl);
        return {
          objectUrl,
          ext
        };
      }
    }catch(error){
      errors.push(error);
    }
  }

  throw errors[errors.length - 1] || new Error('Nie znaleziono zdjęcia dla wybranego indeksu.');
}

async function resolveImageAssetUrl(index, baseUrl){
  const normalizedIndex = normalizeIndexValue(index) || String(index ?? '').trim();
  const base = String(baseUrl || currentImageBaseUrl || '').trim();
  if(!normalizedIndex || !base){
    return '';
  }

  if(isPrivateMediaBase(base)){
    const result = await resolvePrivateImageSource(normalizedIndex, base);
    rememberImageExt(normalizedIndex, base, result.ext);
    return result.objectUrl;
  }

  const ext = getPreferredImageExt(normalizedIndex, base);
  return buildImageUrl(normalizedIndex, ext, base);
}

function createImageExtensionCache(){
  if(typeof localStorage === 'undefined') return new Map();
  try{
    const raw = localStorage.getItem(IMAGE_EXTENSION_CACHE_KEY);
    if(!raw) return new Map();
    const parsed = JSON.parse(raw);
    if(!parsed || typeof parsed !== 'object') return new Map();
    return new Map(
      Object.entries(parsed).filter(([key, value]) => key && SUPPORTED_IMAGE_EXTENSIONS.includes(value))
    );
  }catch(e){
    console.warn('Image cache restore failed', e);
    return new Map();
  }
}

function persistImageExtensionCache(){
  if(typeof localStorage === 'undefined') return;
  try{
    const trimmedEntries = Array.from(imageExtensionCache.entries()).slice(-MAX_IMAGE_EXTENSION_CACHE_ENTRIES);
    localStorage.setItem(IMAGE_EXTENSION_CACHE_KEY, JSON.stringify(Object.fromEntries(trimmedEntries)));
  }catch(e){
    console.warn('Image cache save failed', e);
  }
}

function getImageCacheKey(index, baseUrl){
  const normalizedIndex = normalizeIndexValue(index);
  const base = String(baseUrl || currentImageBaseUrl || '').trim();
  if(!normalizedIndex || !base) return '';
  return `${base}|${normalizedIndex}`;
}

function getNormalizedImageExt(value){
  const normalized = String(value || '').trim().toLowerCase();
  return SUPPORTED_IMAGE_EXTENSIONS.includes(normalized) ? normalized : '';
}

function getDefaultImageExtForBase(baseUrl){
  if(isPrivateMediaBase(baseUrl)){
    return 'webp';
  }
  return isUkrainaImageBase(baseUrl) ? 'png' : 'webp';
}

function getImageCandidateExts(index, baseUrl, preferredExtOverride = ''){
  const preferred = getNormalizedImageExt(preferredExtOverride) || getPreferredImageExt(index, baseUrl);
  return [preferred, ...SUPPORTED_IMAGE_EXTENSIONS.filter(ext => ext !== preferred)];
}

function getPreferredImageExt(index, baseUrl){
  if(isPrivateMediaBase(baseUrl)){
    return getDefaultImageExtForBase(baseUrl);
  }
  const key = getImageCacheKey(index, baseUrl);
  return (key && imageExtensionCache.get(key)) || getDefaultImageExtForBase(baseUrl);
}

function rememberImageExt(index, baseUrl, ext){
  const normalizedExt = getNormalizedImageExt(ext) || getDefaultImageExtForBase(baseUrl);
  const key = getImageCacheKey(index, baseUrl);
  if(!key) return;
  if(imageExtensionCache.has(key)) imageExtensionCache.delete(key);
  imageExtensionCache.set(key, normalizedExt);
  while(imageExtensionCache.size > MAX_IMAGE_EXTENSION_CACHE_ENTRIES){
    const oldestKey = imageExtensionCache.keys().next().value;
    if(!oldestKey) break;
    imageExtensionCache.delete(oldestKey);
  }
  persistImageExtensionCache();
}

window.getPreferredImageExt = getPreferredImageExt;
window.rememberImageExt = rememberImageExt;
window.resolveImageAssetUrl = resolveImageAssetUrl;
window.resolvePrivateImageSource = resolvePrivateImageSource;

function buildDeferredImageTag(src, className, options = {}){
  const {
    alt = '',
    attrs = '',
    onload = '',
    onerror = '',
    popup = false
  } = options;
  const sourceAttr = popup ? '' : `src="${IMAGE_PLACEHOLDER_SRC}"`;
  const loadAttr = popup ? 'eager' : 'lazy';
  const lazyFlag = popup ? 'data-popup-image="1"' : 'data-lazy-image="1"';
  const fetchPriority = popup ? 'high' : 'low';
  return `<img class="${className}" ${sourceAttr} data-src="${escapeAttr(src)}" data-state="idle" ${lazyFlag} loading="${loadAttr}" decoding="async" fetchpriority="${fetchPriority}"${onload ? ` onload="${onload}"` : ''}${onerror ? ` onerror="${onerror}"` : ''}${attrs ? ` ${attrs}` : ''} alt="${escapeAttr(alt)}">`;
}

function buildProductImageTag(index, baseUrl, className, options = {}){
  if(!index) return '';
  const base = baseUrl || currentImageBaseUrl;
  const ext = getPreferredImageExt(index, base);
  const src = buildImageUrl(index, ext, base);
  const extOrder = getImageCandidateExts(index, base).join(',');
  const attrs = `data-index="${escapeAttr(index)}" data-base="${escapeAttr(base)}" data-tried="${ext}" data-ext-order="${escapeAttr(extOrder)}"${options.attrs ? ` ${options.attrs}` : ''}`;
  return buildDeferredImageTag(src, className, {
    ...options,
    attrs,
    onload: 'handleDeferredImageLoad(this)',
    onerror: 'imageFallback(this)'
  });
}

function ensureLazyImageObserver(){
  if(lazyImageObserver || typeof IntersectionObserver === 'undefined') return;
  lazyImageObserver = new IntersectionObserver(entries => {
    entries.forEach(entry => {
      if(!entry.isIntersecting) return;
      lazyImageObserver.unobserve(entry.target);
      loadDeferredImage(entry.target);
    });
  }, { rootMargin: '300px 0px' });
}

async function loadDeferredImage(img){
  if(!img) return;
  const state = img.getAttribute('data-state');
  if(state === 'loading' || state === 'loaded') return;

  const index = img.getAttribute('data-index');
  const base = img.getAttribute('data-base') || currentImageBaseUrl;

  if(index && isPrivateMediaBase(base)){
    img.setAttribute('data-state', 'loading');
    try{
      const result = await resolvePrivateImageSource(index, base);
      img.setAttribute('data-tried', result.ext);
      img.setAttribute('data-src', result.objectUrl);
      img.src = result.objectUrl;
    }catch(error){
      console.error('Private image load error', error);
      img.setAttribute('data-state', 'error');
      img.classList.add('img-missing');
      const pop = img.closest('.img-hover')?.querySelector('.index-img-large');
      if(pop){
        pop.setAttribute('data-state', 'error');
        pop.classList.add('img-missing');
      }
    }
    return;
  }

  const src = img.getAttribute('data-src');
  if(!src) return;
  img.setAttribute('data-state', 'loading');
  img.src = src;
}

function registerLazyImages(root = document){
  const images = root.querySelectorAll('img[data-lazy-image="1"]');
  if(!images.length) return;
  if(typeof IntersectionObserver === 'undefined'){
    images.forEach(loadDeferredImage);
    return;
  }
  ensureLazyImageObserver();
  images.forEach(img => {
    const state = img.getAttribute('data-state');
    if(state === 'loading' || state === 'loaded') return;
    lazyImageObserver.observe(img);
  });
}

function handleDeferredImageLoad(img){
  const state = img.getAttribute('data-state');
  if(state !== 'loading' && state !== 'fallback') return;
  img.setAttribute('data-state', 'loaded');
  img.classList.remove('img-missing');

  const index = img.getAttribute('data-index');
  const base = img.getAttribute('data-base') || currentImageBaseUrl;
  const tried = img.getAttribute('data-tried') || getDefaultImageExtForBase(base);
  if(!index) return;

  rememberImageExt(index, base, tried);
  const popupImage = img.closest('.img-hover')?.querySelector('.index-img-large');
  if(popupImage && popupImage !== img && !popupImage.getAttribute('src')){
    popupImage.setAttribute('data-tried', tried);
    popupImage.setAttribute('data-src', buildImageUrl(index, tried, base));
  }
}

function preparePopupImage(img){
  if(!img) return;
  const index = img.getAttribute('data-index');
  const base = img.getAttribute('data-base') || currentImageBaseUrl;
  const ext = getPreferredImageExt(index, base);
  const extOrder = getImageCandidateExts(index, base).join(',');
  img.setAttribute('data-tried', ext);
  img.setAttribute('data-ext-order', extOrder);
  img.setAttribute('data-src', buildImageUrl(index, ext, base));
  loadDeferredImage(img);
}

function imageFallback(img){
  const index = img.getAttribute('data-index');
  const base = img.getAttribute('data-base') || currentImageBaseUrl;
  const tried = img.getAttribute('data-tried');
  const pop = img.closest('.img-hover')?.querySelector('.index-img-large');
  const extOrder = String(img.getAttribute('data-ext-order') || getImageCandidateExts(index, base).join(','))
    .split(',')
    .map(getNormalizedImageExt)
    .filter(Boolean);
  const triedIndex = extOrder.indexOf(tried);
  const nextExt = triedIndex === -1 ? (extOrder[0] || '') : (extOrder[triedIndex + 1] || '');
  if(nextExt){
    if(index && isPrivateMediaBase(base)){
      img.setAttribute('data-tried', nextExt);
      img.setAttribute('data-ext-order', extOrder.join(','));
      img.setAttribute('data-state', 'loading');
      resolvePrivateImageSource(index, base, nextExt)
        .then(result => {
          img.setAttribute('data-tried', result.ext);
          img.setAttribute('data-src', result.objectUrl);
          img.src = result.objectUrl;
          if(pop){
            pop.setAttribute('data-tried', result.ext);
            pop.setAttribute('data-src', result.objectUrl);
            if(pop.getAttribute('src')){
              pop.setAttribute('data-state', 'loading');
              pop.src = result.objectUrl;
            }
          }
        })
        .catch(error => {
          console.error('Private image fallback error', error);
          img.setAttribute('data-state', 'error');
          img.classList.add('img-missing');
          if(pop){
            pop.setAttribute('data-state', 'error');
            pop.classList.add('img-missing');
          }
        });
      return;
    }
    const nextUrl = buildImageUrl(index, nextExt, base);
    img.setAttribute('data-tried', nextExt);
    img.setAttribute('data-ext-order', extOrder.join(','));
    img.setAttribute('data-src', nextUrl);
    img.setAttribute('data-state', 'fallback');
    img.src = nextUrl;
    if(pop){
      pop.setAttribute('data-tried', nextExt);
      pop.setAttribute('data-ext-order', extOrder.join(','));
      pop.setAttribute('data-src', nextUrl);
      if(pop.getAttribute('src')){
        pop.setAttribute('data-state', 'fallback');
        pop.src = nextUrl;
      }
    }
    return;
  }
  img.setAttribute('data-state', 'error');
  img.classList.add('img-missing');
  if(pop){
    pop.setAttribute('data-state', 'error');
    pop.classList.add('img-missing');
  }
}

function positionPopup(el){
  const pop = el.querySelector('.img-pop');
  if(!pop) return;
  preparePopupImage(pop.querySelector('.index-img-large'));
  pop.style.left = '';
  pop.style.right = '';
  const rect = el.getBoundingClientRect();
  const popWidth = pop.offsetWidth || 260;
  const popHeight = pop.offsetHeight || 260;
  const padding = 16;

  // Horizontal position: prefer right, else left
  const spaceRight = window.innerWidth - rect.right;
  const spaceLeft = rect.left;
  if(spaceRight < popWidth + padding && spaceLeft >= popWidth + padding){
    pop.style.left = `${Math.max(padding, rect.left - popWidth - 12)}px`;
  }else{
    pop.style.left = `${Math.min(window.innerWidth - popWidth - padding, rect.right + 12)}px`;
  }

  // Vertical position: clamp within viewport
  const centerY = rect.top + rect.height / 2;
  let top = centerY - popHeight / 2;
  if(top < padding) top = padding;
  if(top + popHeight > window.innerHeight - padding){
    top = window.innerHeight - popHeight - padding;
  }
  pop.style.top = `${top}px`;
}

function parseSalesValue(value){
  if(typeof value === 'number'){
    return Number.isFinite(value) ? value : 0;
  }

  const normalized = String(value ?? '')
    .replace(/\s+/g, '')
    .replace(',', '.');
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function stripFileExtension(value){
  return String(value || '').replace(/\.[^.]+$/u, '');
}

function normalizeUserSalesConfigData(data){
  const source = data && typeof data === 'object' ? data : {};
  const storageFolder = String(source.salesReportFolder || source.storageFolder || '').trim();
  const displayName = String(source.salesRepName || source.displayName || '').trim() || storageFolder;
  return storageFolder ? { storageFolder, displayName } : null;
}

function getReportEntryUpdatedAt(reportEntry){
  return String(
    reportEntry?.metadata?.updated
    || reportEntry?.metadata?.timeCreated
    || ''
  ).trim();
}

function getReportEntryCacheKey(reportEntry){
  const fullPath = String(
    reportEntry?.item?.fullPath
    || reportEntry?.fullPath
    || reportEntry?.item?.name
    || reportEntry?.name
    || ''
  ).trim();
  const updatedAt = getReportEntryUpdatedAt(reportEntry);
  if(!fullPath) return '';
  return `${fullPath}::${updatedAt}`;
}

function decodeBase64ToUint8Array(base64Value){
  const normalized = String(base64Value || '').trim();
  if(!normalized) return new Uint8Array();
  const binary = atob(normalized);
  const bytes = new Uint8Array(binary.length);
  for(let index = 0; index < binary.length; index += 1){
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function buildCallableReportEntries(reports){
  const entries = Array.isArray(reports) ? reports : [];
  return entries.map(report => {
    const payload = report && typeof report === 'object' ? report : {};
    const metadata = payload.metadata && typeof payload.metadata === 'object'
      ? payload.metadata
      : {};
    const name = String(payload.name || payload.fileName || payload.fullPath || '').trim().split('/').pop() || '';
    const fullPath = String(payload.fullPath || payload.path || name).trim();
    const contentBase64 = String(payload.contentBase64 || '').trim();
    const downloadUrl = String(payload.downloadUrl || '').trim();
    const bytes = contentBase64 ? decodeBase64ToUint8Array(contentBase64) : null;
    const item = {
      name,
      fullPath,
      async getBytes(){
        if(!bytes) throw new Error('Brak inline content dla raportu.');
        return bytes;
      },
      async getBlob(){
        if(!bytes) throw new Error('Brak inline content dla raportu.');
        return new Blob([bytes], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
      },
      async getDownloadURL(){
        if(downloadUrl) return downloadUrl;
        throw new Error('Brak download URL dla raportu.');
      }
    };

    return {
      item,
      fullPath,
      contentBase64,
      downloadUrl,
      metadata: {
        updated: String(metadata.updated || '').trim(),
        timeCreated: String(metadata.timeCreated || '').trim(),
        customMetadata: metadata.customMetadata && typeof metadata.customMetadata === 'object'
          ? metadata.customMetadata
          : {}
      },
      updatedAt: Date.parse(metadata.updated || metadata.timeCreated || '') || 0
    };
  });
}

function buildStorageReportEntry(reference, metadata = {}){
  return {
    item: reference,
    fullPath: String(metadata?.fullPath || reference?.fullPath || '').trim(),
    downloadUrl: String(metadata?.downloadUrl || '').trim(),
    metadata: {
      updated: String(metadata.updated || '').trim(),
      timeCreated: String(metadata.timeCreated || '').trim(),
      customMetadata: metadata.customMetadata && typeof metadata.customMetadata === 'object'
        ? metadata.customMetadata
        : {}
    },
    updatedAt: Date.parse(metadata.updated || metadata.timeCreated || '') || 0
  };
}

function withRejectingTimeout(promise, timeoutMs, message){
  return Promise.race([
    promise,
    new Promise((_, reject) => {
      window.setTimeout(() => {
        reject(new Error(message || 'Przekroczono czas oczekiwania na odpowiedź.'));
      }, timeoutMs);
    })
  ]);
}

function padStorageDatePart(value){
  return String(value).padStart(2, '0');
}

function buildAdminWeeklyReportCandidatePaths(daysBack = 45){
  const paths = [];
  const seen = new Set();
  const now = new Date();

  for(let offset = 0; offset <= daysBack; offset += 1){
    const date = new Date(now);
    date.setDate(now.getDate() - offset);
    const ddmm = `${padStorageDatePart(date.getDate())}${padStorageDatePart(date.getMonth() + 1)}`;
    const path = `${SALES_REPORTS_WEEKLY_ADMIN_ROOT}/kh_tygodnie_${ddmm}.xlsx`;
    if(!seen.has(path)){
      seen.add(path);
      paths.push(path);
    }
  }

  return paths;
}

async function getDirectAdminWeeklySalesReportFromCandidates(){
  if(!storage || typeof storage.ref !== 'function') return null;

  const candidatePaths = buildAdminWeeklyReportCandidatePaths();
  for(const path of candidatePaths){
    const reference = storage.ref().child(path);
    try{
      const metadata = await withRejectingTimeout(
        reference.getMetadata(),
        STORAGE_LIST_TIMEOUT_MS,
        `Timeout metadata dla ${path}`
      );
      return buildStorageReportEntry(reference, metadata);
    }catch(error){
      const code = String(error?.code || '').trim();
      try{
        const downloadUrl = typeof reference.getDownloadURL === 'function'
          ? await withRejectingTimeout(
            reference.getDownloadURL(),
            STORAGE_LIST_TIMEOUT_MS,
            `Timeout download URL dla ${path}`
          )
          : '';
        if(downloadUrl){
          return buildStorageReportEntry(reference, {
            fullPath: path,
            downloadUrl,
            updated: '',
            timeCreated: '',
            customMetadata: {
              reportType: 'weekly-sales'
            }
          });
        }
      }catch(downloadError){
        const downloadCode = String(downloadError?.code || '').trim();
        if((code && code !== 'storage/object-not-found') || (downloadCode && downloadCode !== 'storage/object-not-found')){
          console.error('Admin weekly direct report candidate error', path, error, downloadError);
        }
      }
    }
  }

  return null;
}

async function getPersonalSalesConfig(uid, options = {}){
  const cacheKey = String(uid || '').trim();
  if(!cacheKey) return null;
  if(!options.forceRefresh && personalSalesConfigCache.has(cacheKey)){
    return personalSalesConfigCache.get(cacheKey);
  }

  let config = null;

  if(functionsService && typeof functionsService.httpsCallable === 'function'){
    try{
      const callable = functionsService.httpsCallable('getPersonalSalesConfig');
      const result = await callable({});
      config = normalizeUserSalesConfigData(result?.data?.config);
    }catch(error){
      console.error('Personal sales config callable error', error);
    }
  }

  if(!config && db){
    try{
      const userDoc = await db.collection(USERS_COLLECTION).doc(cacheKey).get();
      if(userDoc.exists){
        config = normalizeUserSalesConfigData(userDoc.data());
      }
    }catch(error){
      console.error('Personal sales config Firestore read error', error);
    }
  }

  if(!config){
    const fallbackConfig = PERSONAL_SALES_USERS[cacheKey];
    config = normalizeUserSalesConfigData(fallbackConfig);
  }

  personalSalesConfigCache.set(cacheKey, config || null);
  return config || null;
}

async function getAllAssignedSalesConfigs(options = {}){
  if(!options.forceRefresh && Array.isArray(adminAssignedSalesConfigsCache)){
    return adminAssignedSalesConfigsCache;
  }
  if(!db){
    const fallbackConfigs = Object.values(PERSONAL_SALES_USERS || {})
      .map(entry => normalizeUserSalesConfigData(entry))
      .filter(Boolean)
      .filter((config, index, array) => array.findIndex(item => item.storageFolder === config.storageFolder) === index)
      .sort((a, b) => String(a.displayName || '').localeCompare(String(b.displayName || ''), 'pl'));
    adminAssignedSalesConfigsCache = fallbackConfigs;
    return fallbackConfigs;
  }

  try{
    const snapshot = await db.collection(USERS_COLLECTION).get();
    const firestoreConfigs = snapshot.docs
      .map(doc => normalizeUserSalesConfigData(doc.data()))
      .filter(Boolean);
    const fallbackConfigs = Object.values(PERSONAL_SALES_USERS || {})
      .map(entry => normalizeUserSalesConfigData(entry))
      .filter(Boolean);
    const configs = [...firestoreConfigs, ...fallbackConfigs]
      .filter((config, index, array) => array.findIndex(entry => entry.storageFolder === config.storageFolder) === index)
      .sort((a, b) => String(a.displayName || '').localeCompare(String(b.displayName || ''), 'pl'));
    adminAssignedSalesConfigsCache = configs;
    return configs;
  }catch(error){
    console.error('Assigned sales configs read error', error);
    const fallbackConfigs = Object.values(PERSONAL_SALES_USERS || {})
      .map(entry => normalizeUserSalesConfigData(entry))
      .filter(Boolean)
      .filter((config, index, array) => array.findIndex(item => item.storageFolder === config.storageFolder) === index)
      .sort((a, b) => String(a.displayName || '').localeCompare(String(b.displayName || ''), 'pl'));
    adminAssignedSalesConfigsCache = fallbackConfigs;
    return fallbackConfigs;
  }
}

async function listStorageSpreadsheetEntries(prefixes){
  if(!storage) return [];
  const entries = [];

  for(const prefix of prefixes){
    if(!prefix) continue;
    try{
      const folderRef = storage.ref().child(prefix);
      const listResult = await folderRef.listAll();
      const topLevelItems = listResult.items.filter(item => {
        const name = String(item.name || '');
        return /\.xlsx$/i.test(name);
      });
      const folderEntries = await Promise.all(topLevelItems.map(async item => {
        const metadata = await item.getMetadata();
        const updatedAt = Date.parse(metadata.updated || metadata.timeCreated || '') || 0;
        return {
          item,
          metadata,
          updatedAt
        };
      }));
      entries.push(...folderEntries);
    }catch(error){
      console.error('Storage report list error', prefix, error);
    }
  }

  entries.sort((a, b) => {
    if(b.updatedAt !== a.updatedAt) return b.updatedAt - a.updatedAt;
    return String(b.item?.name || '').localeCompare(String(a.item?.name || ''), 'pl');
  });

  return entries;
}

async function listPersonalSalesReportEntries(config, options = {}){
  const storageFolder = String(config?.storageFolder || '').trim();
  if(!storageFolder){
    return {
      weeklyEntries: [],
      detailedEntries: []
    };
  }

  const cacheKey = `${storageFolder}::${options.forceRefresh ? 'force' : 'cache'}`;
  if(!options.forceRefresh && personalSalesFolderEntriesCache.has(cacheKey)){
    return personalSalesFolderEntriesCache.get(cacheKey);
  }

  const weeklyPrefixes = [
    `${SALES_REPORTS_WEEKLY_REP_ROOT}/${storageFolder}`,
    `${LEGACY_SALES_REPORTS_ROOT}/${storageFolder}`
  ];
  const detailedPrefixes = [
    `${SALES_REPORTS_INDEX_REP_ROOT}/${storageFolder}`,
    `${LEGACY_SALES_REPORTS_ROOT}/${storageFolder}`
  ];

  const result = {
    weeklyEntries: await listStorageSpreadsheetEntries(weeklyPrefixes),
    detailedEntries: await listStorageSpreadsheetEntries(detailedPrefixes)
  };

  personalSalesFolderEntriesCache.set(cacheKey, result);
  return result;
}

function convertBytesToArrayBuffer(bytes){
  if(bytes instanceof ArrayBuffer){
    return bytes;
  }
  if(ArrayBuffer.isView(bytes)){
    return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  }
  throw new Error('Nieobsługiwany format pobranych danych raportu.');
}

function buildReportDownloadErrorMessage(errorList){
  const hasFetchError = Array.isArray(errorList) && errorList.some(error =>
    String(error?.message || '').toLowerCase().includes('failed to fetch')
  );
  if(hasFetchError){
    return 'Nie udało się pobrać raportu sprzedaży. Odśwież stronę i spróbuj ponownie.';
  }
  return 'Nie udało się pobrać pliku raportu z Firebase Storage.';
}

async function fetchArrayBufferFromDownloadUrl(downloadUrl){
  const response = await fetch(downloadUrl, { cache: 'no-store' });
  if(!response.ok){
    throw new Error(`Pobieranie raportu nie powiodło się (HTTP ${response.status}).`);
  }
  return response.arrayBuffer();
}

function fetchArrayBufferViaXhr(downloadUrl){
  return new Promise((resolve, reject) => {
    const request = new XMLHttpRequest();
    request.open('GET', downloadUrl, true);
    request.responseType = 'arraybuffer';
    request.onload = () => {
      if(request.status >= 200 && request.status < 300){
        resolve(request.response);
      }else{
        reject(new Error(`Pobieranie raportu nie powiodło się (HTTP ${request.status}).`));
      }
    };
    request.onerror = () => reject(new Error('Błąd sieci podczas pobierania raportu.'));
    request.send();
  });
}

async function readReportEntryArrayBuffer(reportEntry){
  if(String(reportEntry?.contentBase64 || '').trim()){
    return convertBytesToArrayBuffer(decodeBase64ToUint8Array(reportEntry.contentBase64));
  }

  const reference = reportEntry?.item;
  if(!reference){
    throw new Error('Brak pliku raportu do odczytu.');
  }

  const errors = [];

  if(typeof reference.getBlob === 'function'){
    try{
      const blob = await reference.getBlob();
      return await blob.arrayBuffer();
    }catch(error){
      errors.push(error);
    }
  }

  if(typeof reference.getBytes === 'function'){
    try{
      const bytes = await reference.getBytes(25 * 1024 * 1024);
      return convertBytesToArrayBuffer(bytes);
    }catch(error){
      errors.push(error);
    }
  }

  try{
    const downloadUrl = typeof reference.getDownloadURL === 'function'
      ? await reference.getDownloadURL()
      : String(reportEntry?.downloadUrl || '').trim();
    if(downloadUrl){
      try{
        return await fetchArrayBufferFromDownloadUrl(downloadUrl);
      }catch(fetchError){
        errors.push(fetchError);
        try{
          return await fetchArrayBufferViaXhr(downloadUrl);
        }catch(xhrError){
          errors.push(xhrError);
        }
      }
    }
  }catch(error){
    errors.push(error);
  }

  console.error('Report download attempts failed', errors);
  throw new Error(buildReportDownloadErrorMessage(errors));
}

async function getLatestPersonalSalesReport(config, options = {}){
  if(!config) return null;

  if(!options.skipCallable && !isCurrentUserAdmin() && functionsService && typeof functionsService.httpsCallable === 'function'){
    try{
      const callable = functionsService.httpsCallable('getPersonalWeeklySalesReport');
      const result = await callable({});
      const reportEntries = buildCallableReportEntries(result?.data?.report ? [result.data.report] : []);
      if(reportEntries.length){
        return reportEntries[0];
      }
    }catch(error){
      console.error('Personal weekly sales callable error', error);
    }
  }

  const { weeklyEntries } = await listPersonalSalesReportEntries(config, options);
  return weeklyEntries[0] || null;
}

async function getCachedWeeklySalesReportData(reportEntry){
  const cacheKey = getReportEntryCacheKey(reportEntry);
  if(cacheKey && weeklySalesParsedReportCache.has(cacheKey)){
    return weeklySalesParsedReportCache.get(cacheKey);
  }
  if(cacheKey && weeklySalesParsedReportPromiseCache.has(cacheKey)){
    return weeklySalesParsedReportPromiseCache.get(cacheKey);
  }

  const loadPromise = (async () => {
    const buffer = await readReportEntryArrayBuffer(reportEntry);
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheetName = workbook.SheetNames?.[0];
    if(!sheetName){
      throw new Error('Raport tygodniowy nie zawiera żadnego arkusza.');
    }
    const sheet = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const rows = normalizeWeeklySalesRowsFromMatrix(matrix);
    return {
      rows,
      fileName: reportEntry?.item?.name || '',
      updatedAt: getReportEntryUpdatedAt(reportEntry)
    };
  })();

  if(cacheKey){
    weeklySalesParsedReportPromiseCache.set(cacheKey, loadPromise);
  }

  try{
    const data = await loadPromise;
    if(cacheKey){
      weeklySalesParsedReportCache.set(cacheKey, data);
    }
    return data;
  }finally{
    if(cacheKey){
      weeklySalesParsedReportPromiseCache.delete(cacheKey);
    }
  }
}

function renderPersonalSalesInsightLoading(config){
  setPersonalSalesInsightMarkup(`
    <div class="sales-insight-layout">
      <div class="sales-insight-shell">
        <div class="sales-insight-main">
          <span class="sales-insight-eyebrow">Twoja sprzedaż</span>
          <h2>${escapeHtml(config.displayName)}</h2>
          <p>Ładuję porównanie dwóch ostatnich tygodni sprzedaży z najnowszego raportu.</p>
        </div>
        <div class="sales-insight-side">
          <div class="admin-user-empty">Pobieram raport z folderu ${escapeHtml(config.storageFolder)}...</div>
        </div>
      </div>
      ${renderPersonalSalesPotentialCard({ state: 'loading' })}
    </div>
  `);
}

function renderPersonalSalesInsightEmpty(config, message){
  setPersonalSalesInsightMarkup(`
    <div class="sales-insight-empty">
      <strong>${escapeHtml(config.displayName)}</strong><br>
      ${escapeHtml(message)}
    </div>
  `);
}

function renderPersonalSalesPotentialCard(options = {}){
  const {
    state = 'ready',
    summary = null,
    message = ''
  } = options;

  if(state === 'loading'){
    return `
      <div class="sales-potential-shell is-loading">
        <div class="sales-potential-main">
          <div class="sales-potential-header">
            <span class="sales-potential-eyebrow">Potencjał klienta</span>
            <h3 class="sales-potential-title">Analizuję ofertę Top Rumunia</h3>
            <p class="sales-potential-subtitle">Pobieram raport per indeks i buduję skrót kategorii dla Twoich klientów.</p>
          </div>
        </div>
        <div class="sales-potential-side">
          <div class="sales-potential-empty">Ładuję dane raportowe...</div>
        </div>
      </div>
    `;
  }

  if(state === 'error' || state === 'empty' || !summary){
    return `
      <div class="sales-potential-shell">
        <div class="sales-potential-main">
          <div class="sales-potential-header">
            <span class="sales-potential-eyebrow">Potencjał klienta</span>
            <h3 class="sales-potential-title">Brak skrótu potencjału klienta</h3>
            <p class="sales-potential-subtitle">${escapeHtml(message || 'Nie udało się jeszcze przygotować danych do tego widoku.')}</p>
          </div>
        </div>
        <div class="sales-potential-side">
          <div class="sales-potential-empty">Sekcja wróci automatycznie po poprawnym wczytaniu raportu per indeks.</div>
        </div>
      </div>
    `;
  }

  const groupsHtml = summary.groupBreakdown.map(group => `
    <div class="sales-potential-group">
      <div class="sales-potential-group-head">
        <span class="sales-potential-group-name">${escapeHtml(group.shortLabel || group.groupLabel || 'Bez grupy')}</span>
        <span class="sales-potential-group-value">${escapeHtml(formatPercent(group.coveragePercent, { minimumFractionDigits: 0, maximumFractionDigits: 0 }))}</span>
      </div>
      <div class="sales-potential-group-track">
        <span class="sales-potential-group-fill" style="width:${Math.max(0, Math.min(group.coveragePercent, 100)).toFixed(2)}%"></span>
      </div>
    </div>
  `).join('');
  const strongestGroupLabel = summary.strongestGroup?.shortLabel || summary.strongestGroup?.groupLabel || '';
  const strongestGroupPercent = Number(summary.strongestGroup?.coveragePercent || 0);
  const opportunityGroupLabel = summary.opportunityGroup?.shortLabel || summary.opportunityGroup?.groupLabel || '';
  const opportunityGroupPercent = Number(summary.opportunityGroup?.coveragePercent || 0);

  return `
    <div class="sales-potential-shell">
      <div class="sales-potential-main">
        <div class="sales-potential-header">
          <span class="sales-potential-eyebrow">Potencjał klienta</span>
          <h3 class="sales-potential-title">Twoi klienci kupują średnio ${escapeHtml(formatPercent(summary.averageCoverage, { minimumFractionDigits: 0, maximumFractionDigits: 0 }))} oferty Top Rumunia</h3>
          <p class="sales-potential-subtitle">${escapeHtml(formatNumber(summary.activeCustomers, { maximumFractionDigits: 0 }))} z ${escapeHtml(formatNumber(summary.totalCustomers, { maximumFractionDigits: 0 }))} klientów ma już aktywność w ofercie rankingowej.</p>
        </div>
        <div class="sales-potential-kpis">
          <div class="sales-potential-kpi">
            <span class="sales-potential-kpi-label">Klienci w analizie</span>
            <strong class="sales-potential-kpi-value">${escapeHtml(formatNumber(summary.totalCustomers, { maximumFractionDigits: 0 }))}</strong>
          </div>
          <div class="sales-potential-kpi">
            <span class="sales-potential-kpi-label">Kupujący Top</span>
            <strong class="sales-potential-kpi-value">${escapeHtml(formatNumber(summary.activeCustomers, { maximumFractionDigits: 0 }))}</strong>
          </div>
          <div class="sales-potential-kpi">
            <span class="sales-potential-kpi-label">Średni udział oferty</span>
            <strong class="sales-potential-kpi-value">${escapeHtml(formatPercent(summary.averageCoverage, { minimumFractionDigits: 0, maximumFractionDigits: 0 }))}</strong>
          </div>
        </div>
        <div class="sales-potential-insight">
          <span class="sales-potential-insight-label">Zakres analizy</span>
          <strong class="sales-potential-insight-value">${escapeHtml(summary.scopeLabel || 'Ostatni raport per indeks')}</strong>
          <span class="sales-potential-insight-copy">
            ${escapeHtml(
              strongestGroupLabel && opportunityGroupLabel
                ? `Najwyższy udział ma ${strongestGroupLabel} (${formatPercent(strongestGroupPercent, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}), a największa luka zostaje w ${opportunityGroupLabel} (${formatPercent(opportunityGroupPercent, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}).`
                : 'Analiza pokazuje aktualne pokrycie oferty Top Rumunia w kategoriach Twoich klientów.'
            )}
          </span>
        </div>
      </div>
      <div class="sales-potential-side">
        <div class="sales-potential-groups">
          <div class="sales-potential-groups-title">Udział według kategorii</div>
          ${groupsHtml || '<div class="sales-potential-empty">Brak kategorii do wyświetlenia.</div>'}
        </div>
        <div class="sales-potential-footnote">
          ${summary.updatedAt ? `Aktualizacja: ${escapeHtml(formatInsightDate(summary.updatedAt))}` : 'Aktualizacja: brak daty'}
          ${summary.fileName ? `<br>Raport: ${escapeHtml(summary.fileName)}` : ''}
        </div>
      </div>
    </div>
  `;
}

function buildPersonalSalesPotentialSummary(detailRows, topRows){
  const sourceRows = Array.isArray(detailRows) ? detailRows : [];
  const offerRows = Array.isArray(topRows) ? topRows.filter(Boolean) : [];
  if(!sourceRows.length || !offerRows.length) return null;

  const groupOrder = new Map(REPORT_GROUP_CONFIGS.map((group, index) => [group.dashboardLabel, index]));
  const customers = new Map();
  sourceRows.forEach(row => {
    const customerKey = getDetailedSalesCustomerKey(row.customerCode, row.customerName);
    if(!customerKey || customerKey === '|||') return;
    const current = customers.get(customerKey) || {
      customerKey,
      customerCode: row.customerCode,
      customerName: row.customerName,
      purchasedIndexSet: new Set(),
      purchasedNameSet: new Set(),
      purchasedCompositeSet: new Set()
    };
    const normalizedIndex = normalizeIndexValue(row.index);
    const normalizedName = normalizeDetailedSalesName(row.name);
    if(normalizedIndex){
      current.purchasedIndexSet.add(normalizedIndex);
    }
    if(normalizedName){
      current.purchasedNameSet.add(normalizedName);
    }
    if(normalizedIndex && normalizedName){
      current.purchasedCompositeSet.add(`${normalizedIndex}|||${normalizedName}`);
    }
    customers.set(customerKey, current);
  });

  const customerRows = Array.from(customers.values());
  if(!customerRows.length) return null;

  const groupStats = new Map();
  offerRows.forEach(topRow => {
    const groupLabel = getDetailedSalesDashboardGroupLabel(topRow.group);
    const groupKey = getDetailedSalesTopGroupKey(groupLabel) || normalizeHeaderKey(groupLabel) || 'bez-grupy';
    const current = groupStats.get(groupKey) || {
      groupKey,
      groupLabel,
      shortLabel: getDetailedSalesShortGroupLabel(groupLabel),
      offerCount: 0,
      purchasedCount: 0
    };
    current.offerCount += 1;
    groupStats.set(groupKey, current);
  });

  const coverageRows = customerRows.map(customer => {
    let purchasedCount = 0;
    const groupPurchased = new Map();

    offerRows.forEach(topRow => {
      if(!isDetailedSalesTopMatch(topRow, customer.purchasedIndexSet, customer.purchasedNameSet, customer.purchasedCompositeSet)){
        return;
      }
      purchasedCount += 1;
      const groupLabel = getDetailedSalesDashboardGroupLabel(topRow.group);
      const groupKey = getDetailedSalesTopGroupKey(groupLabel) || normalizeHeaderKey(groupLabel) || 'bez-grupy';
      groupPurchased.set(groupKey, (groupPurchased.get(groupKey) || 0) + 1);
    });

    groupPurchased.forEach((count, groupKey) => {
      const current = groupStats.get(groupKey);
      if(current){
        current.purchasedCount += count;
      }
    });

    return {
      purchasedCount,
      coveragePercent: getDetailedSalesTopCoveragePercent(purchasedCount, offerRows.length)
    };
  });

  const totalCustomers = coverageRows.length;
  const activeCustomers = coverageRows.filter(row => row.purchasedCount > 0).length;
  const averageCoverage = coverageRows.reduce((sum, row) => sum + Number(row.coveragePercent || 0), 0) / (totalCustomers || 1);
  const groupBreakdown = Array.from(groupStats.values())
    .map(group => {
      const possiblePurchases = group.offerCount * totalCustomers;
      return {
        ...group,
        coveragePercent: getDetailedSalesTopCoveragePercent(group.purchasedCount, possiblePurchases)
      };
    })
    .sort((a, b) => {
      const aOrder = groupOrder.has(a.groupLabel) ? groupOrder.get(a.groupLabel) : Number.POSITIVE_INFINITY;
      const bOrder = groupOrder.has(b.groupLabel) ? groupOrder.get(b.groupLabel) : Number.POSITIVE_INFINITY;
      if(aOrder !== bOrder) return aOrder - bOrder;
      if(b.coveragePercent !== a.coveragePercent) return b.coveragePercent - a.coveragePercent;
      return String(a.groupLabel || '').localeCompare(String(b.groupLabel || ''), 'pl');
    });
  const strongestGroup = groupBreakdown.length
    ? [...groupBreakdown].sort((a, b) => Number(b.coveragePercent || 0) - Number(a.coveragePercent || 0))[0]
    : null;
  const opportunityGroup = groupBreakdown.length
    ? [...groupBreakdown].sort((a, b) => {
      const gapA = 100 - Number(a.coveragePercent || 0);
      const gapB = 100 - Number(b.coveragePercent || 0);
      if(gapB !== gapA) return gapB - gapA;
      return Number(b.offerCount || 0) - Number(a.offerCount || 0);
    })[0]
    : null;

  return {
    totalCustomers,
    activeCustomers,
    averageCoverage,
    groupBreakdown,
    strongestGroup,
    opportunityGroup
  };
}

async function readPersonalSalesPotentialSummary(config){
  if(typeof getLatestPersonalDetailedSalesReports !== 'function' || typeof getCachedDetailedSalesReportData !== 'function'){
    return null;
  }

  const [{ entries }, topRows] = await Promise.all([
    getLatestPersonalDetailedSalesReports(config, 1),
    loadTopRumuniaOfferRows()
  ]);

  const latestDetailedReport = entries?.[0];
  if(!latestDetailedReport){
    return null;
  }

  const reportData = await getCachedDetailedSalesReportData(latestDetailedReport);
  const summary = buildPersonalSalesPotentialSummary(reportData.rows, topRows);
  if(!summary){
    return null;
  }
  return {
    ...summary,
    scopeLabel: reportData.reportLabel || 'Ostatni raport per indeks',
    updatedAt: reportData.updatedAt || '',
    fileName: reportData.fileName || ''
  };
}

async function loadPersonalSalesPotentialForUser(config, insight, version){
  try{
    const summary = await readPersonalSalesPotentialSummary(config);
    if(version !== authStateVersion) return;
    renderPersonalSalesInsight(config, insight, summary
      ? { state: 'ready', summary }
      : { state: 'empty', message: 'Brak danych z raportu potencjału klienta dla tego przedstawiciela.' });
  }catch(error){
    console.error('Personal sales potential error', error);
    if(version !== authStateVersion) return;
    renderPersonalSalesInsight(config, insight, {
      state: 'error',
      message: error?.message || 'Nie udało się wczytać skrótu potencjału klienta.'
    });
  }
}

function renderPersonalSalesInsight(config, insight, potential = { state: 'loading' }){
  const statusMeta = getSalesInsightStatusMeta(insight.direction);
  const maxValue = Math.max(insight.previousTotal, insight.lastTotal, 1);
  const previousWidth = Math.max((insight.previousTotal / maxValue) * 100, insight.previousTotal > 0 ? 8 : 0);
  const lastWidth = Math.max((insight.lastTotal / maxValue) * 100, insight.lastTotal > 0 ? 8 : 0);

  setPersonalSalesInsightMarkup(`
    <div class="sales-insight-layout">
      <div class="sales-insight-shell ${statusMeta.directionClass}">
        <div class="sales-insight-main">
          <span class="sales-insight-eyebrow">Twoja sprzedaż</span>
          <h2>${escapeHtml(config.displayName)}</h2>
          <p>Porównanie dwóch ostatnich tygodni z najnowszego raportu sprzedaży przypisanego do Twojego konta.</p>
          <div class="sales-insight-status ${statusMeta.statusClass}">
            <span>${statusMeta.statusIcon}</span>
            <strong>${escapeHtml(statusMeta.statusText)}</strong>
          </div>
          <div class="sales-insight-kpis">
            <div class="sales-insight-kpi">
              <span class="sales-insight-kpi-label">Aktualny tydzień</span>
              <strong class="sales-insight-kpi-value">${escapeHtml(formatSalesNumber(insight.lastTotal))}</strong>
            </div>
            <div class="sales-insight-kpi">
              <span class="sales-insight-kpi-label">Poprzedni tydzień</span>
              <strong class="sales-insight-kpi-value">${escapeHtml(formatSalesNumber(insight.previousTotal))}</strong>
            </div>
            <div class="sales-insight-kpi">
              <span class="sales-insight-kpi-label">Trend</span>
              <strong class="sales-insight-kpi-value">${escapeHtml(statusMeta.trendKpiLabel)}</strong>
            </div>
          </div>
          <div class="sales-insight-actions">
            <button
              type="button"
              class="sales-insight-details-btn"
              onclick="openPersonalSalesReportDetails()"
              data-personal-sales-details-btn="1"
            >Pokaż szczegóły</button>
            <button
              type="button"
              class="sales-insight-details-btn"
              onclick="openPersonalSalesTopComparison()"
              data-personal-sales-top-btn="1"
            >Top Rumunia</button>
          </div>
          <div class="sales-insight-change">
            <div class="sales-insight-delta">${escapeHtml(formatSalesDelta(insight.difference))}</div>
            <div class="sales-insight-percent">${escapeHtml(formatPercentageDelta(insight.percentChange))}</div>
          </div>
        </div>
        <div class="sales-insight-side">
          <div class="sales-insight-bars">
            <div class="sales-insight-bar-card">
              <span class="sales-insight-bar-label">${escapeHtml(insight.previousWeek)}</span>
              <div class="sales-insight-bar-track">
                <span class="sales-insight-bar-fill sales-insight-bar-fill--prev" style="width:${previousWidth.toFixed(2)}%"></span>
              </div>
              <span class="sales-insight-bar-value">${escapeHtml(formatSalesNumber(insight.previousTotal))}</span>
            </div>
            <div class="sales-insight-bar-card">
              <span class="sales-insight-bar-label">${escapeHtml(insight.lastWeek)}</span>
              <div class="sales-insight-bar-track">
                <span class="sales-insight-bar-fill sales-insight-bar-fill--last" style="width:${lastWidth.toFixed(2)}%"></span>
              </div>
              <span class="sales-insight-bar-value">${escapeHtml(formatSalesNumber(insight.lastTotal))}</span>
            </div>
          </div>
          <div class="sales-insight-footnote">
            Raport: ${escapeHtml(insight.fileName)}<br>
            Ostatnia aktualizacja: ${escapeHtml(formatInsightDate(insight.updatedAt))}
          </div>
        </div>
      </div>
      ${renderPersonalSalesPotentialCard(potential)}
    </div>
  `);
}

function readPersonalSalesInsightFromMetadata(reportEntry){
  const customMetadata = reportEntry?.metadata?.customMetadata || {};
  const previousWeek = String(customMetadata.insightPreviousWeek || '').trim();
  const lastWeek = String(customMetadata.insightLastWeek || '').trim();

  if(!previousWeek || !lastWeek){
    return null;
  }

  const previousTotal = parseSalesValue(customMetadata.insightPreviousTotal);
  const lastTotal = parseSalesValue(customMetadata.insightLastTotal);
  const difference = lastTotal - previousTotal;
  const percentChange = previousTotal !== 0
    ? (difference / previousTotal) * 100
    : (lastTotal === 0 ? 0 : 100);
  const direction = difference > 0.005 ? 'up' : difference < -0.005 ? 'down' : 'flat';

  return {
    previousWeek,
    lastWeek,
    previousTotal,
    lastTotal,
    difference,
    percentChange,
    direction,
    fileName: reportEntry?.item?.name || '',
    updatedAt: getReportEntryUpdatedAt(reportEntry)
  };
}

function isWeeklySalesReportEntry(reportEntry){
  const customMetadata = reportEntry?.metadata?.customMetadata || {};
  const reportType = String(customMetadata.reportType || '').trim().toLowerCase();

  if(reportType === 'weekly-sales'){
    return true;
  }

  if(reportType === 'sales-by-index'){
    return false;
  }

  const hasWeeklyInsightMetadata =
    String(customMetadata.insightPreviousWeek || '').trim()
    && String(customMetadata.insightLastWeek || '').trim();

  if(hasWeeklyInsightMetadata){
    return true;
  }

  const fileName = String(reportEntry?.item?.name || '').trim().toLowerCase();
  return fileName.includes('per tydzien') || fileName.includes('per tydzień');
}

async function getLatestAdminWeeklySalesReport(){
  if(functionsService && typeof functionsService.httpsCallable === 'function'){
    try{
      const callable = functionsService.httpsCallable('getAdminWeeklySalesReport');
      const result = await callable({});
      const reportEntries = buildCallableReportEntries(result?.data?.report ? [result.data.report] : []);
      if(reportEntries.length){
        return reportEntries[0];
      }
    }catch(error){
      console.error('Admin weekly sales callable error', error);
    }
  }

  const directEntry = await getDirectAdminWeeklySalesReportFromCandidates();
  if(directEntry){
    return directEntry;
  }

  const entries = await withRejectingTimeout(
    listStorageSpreadsheetEntries([
      SALES_REPORTS_WEEKLY_ADMIN_ROOT,
      LEGACY_SALES_REPORTS_ROOT
    ]),
    STORAGE_LIST_TIMEOUT_MS,
    'Timeout listowania raportów administratora.'
  );
  const weeklyEntries = entries.filter(isWeeklySalesReportEntry);
  if(weeklyEntries[0] || entries[0]){
    return weeklyEntries[0] || entries[0] || null;
  }
  return null;
}

async function readPersonalSalesInsight(reportEntry){
  const metadataInsight = readPersonalSalesInsightFromMetadata(reportEntry);
  if(metadataInsight){
    return metadataInsight;
  }

  const buffer = await readReportEntryArrayBuffer(reportEntry);
  const workbook = XLSX.read(buffer, { type: 'array', raw: true, cellDates: true });
  const sheetName = workbook.SheetNames?.[0];

  if(!sheetName){
    throw new Error('Raport sprzedaży nie zawiera arkusza.');
  }

  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: true,
    defval: ''
  });

  if(rows.length < 3){
    throw new Error('Raport sprzedaży ma zbyt mało danych do porównania.');
  }

  const weekLabels = [];
  const headerRow = Array.isArray(rows[0]) ? rows[0] : [];
  for(let columnIndex = 3; columnIndex < headerRow.length; columnIndex += 1){
    const label = String(headerRow[columnIndex] ?? '').trim();
    if(label){
      weekLabels.push({ columnIndex, label });
    }
  }

  if(weekLabels.length < 2){
    throw new Error('Raport sprzedaży nie zawiera dwóch tygodni do porównania.');
  }

  const previousWeek = weekLabels[weekLabels.length - 2];
  const lastWeek = weekLabels[weekLabels.length - 1];
  let previousTotal = 0;
  let lastTotal = 0;

  rows.slice(2).forEach(row => {
    if(!Array.isArray(row)) return;
    previousTotal += parseSalesValue(row[previousWeek.columnIndex]);
    lastTotal += parseSalesValue(row[lastWeek.columnIndex]);
  });

  const difference = lastTotal - previousTotal;
  const percentChange = previousTotal !== 0
    ? (difference / previousTotal) * 100
    : (lastTotal === 0 ? 0 : 100);
  const direction = difference > 0.005 ? 'up' : difference < -0.005 ? 'down' : 'flat';

  return {
    previousWeek: previousWeek.label,
    lastWeek: lastWeek.label,
    previousTotal,
    lastTotal,
    difference,
    percentChange,
    direction,
    fileName: reportEntry?.item?.name || '',
    updatedAt: getReportEntryUpdatedAt(reportEntry)
  };
}

async function loadPersonalSalesInsightForUser(user, version){
  const config = await getPersonalSalesConfig(user?.uid);
  if(!config){
    hidePersonalSalesInsight();
    return;
  }

  renderPersonalSalesInsightLoading(config);

  try{
    const latestReport = await getLatestPersonalSalesReport(config);
    if(version !== authStateVersion) return;

    if(!latestReport){
      renderPersonalSalesInsightEmpty(config, 'Brak podzielonego raportu sprzedaży w folderze tego przedstawiciela.');
      return;
    }

    const insight = await readPersonalSalesInsight(latestReport);
    if(version !== authStateVersion) return;
    renderPersonalSalesInsight(config, insight, { state: 'loading' });
    void loadPersonalSalesPotentialForUser(config, insight, version);
  }catch(error){
    console.error('Personal sales insight error', error);
    if(version !== authStateVersion) return;
    renderPersonalSalesInsightEmpty(config, error?.message || 'Nie udało się wczytać danych sprzedaży.');
  }
}

async function loadAdminRepresentativeWeeklySalesData(options = {}){
  const configs = await getAllAssignedSalesConfigs(options);
  const items = [];

  await Promise.all(configs.map(async config => {
    try{
      const latestReport = await getLatestPersonalSalesReport(config, { skipCallable: true });
      if(!latestReport) return;
      const reportData = await getCachedWeeklySalesReportData(latestReport);
      items.push({
        config,
        rows: reportData.rows,
        fileName: reportData.fileName,
        updatedAt: reportData.updatedAt
      });
    }catch(error){
      console.error('Admin weekly report load error', config, error);
    }
  }));

  return {
    items,
    configsCount: configs.length
  };
}

async function loadAdminSalesInsight(version){
  renderAdminSalesInsightLoading();

  try{
    let summary = null;
    const adminReportEntry = await getLatestAdminWeeklySalesReport();
    if(adminReportEntry){
      const adminReportData = await getCachedWeeklySalesReportData(adminReportEntry);
      const rows = Array.isArray(adminReportData?.rows) ? adminReportData.rows : [];
      if(rows.length){
        summary = buildAdminWeeklySalesInsightSummaryFromRows(rows, {
          updatedAt: adminReportData.updatedAt || getReportEntryUpdatedAt(adminReportEntry),
          fileName: adminReportData.fileName || adminReportEntry?.item?.name || 'Raport administratora'
        });
      }
    }

    if(!summary){
      const teamData = await loadAdminRepresentativeWeeklySalesData({ forceRefresh: true });
      if(version !== authStateVersion) return;

      const rows = teamData.items.flatMap(item => Array.isArray(item.rows) ? item.rows : []);
      if(!rows.length){
        renderAdminSalesInsightEmpty('Brak raportu tygodniowego administratora i brak raportów przypisanych do przedstawicieli handlowych.');
        return;
      }

      const newestUpdatedAt = teamData.items
        .map(item => item.updatedAt)
        .filter(Boolean)
        .sort((a, b) => String(b).localeCompare(String(a), 'pl'))[0] || '';

      summary = buildAdminWeeklySalesInsightSummaryFromRows(rows, {
        updatedAt: newestUpdatedAt,
        fileName: `Raporty PH: ${teamData.items.length}/${teamData.configsCount}`
      });
    }

    if(version !== authStateVersion) return;

    renderAdminSalesInsight(summary);
  }catch(error){
    console.error('Admin sales insight error', error);
    if(version !== authStateVersion) return;
    renderAdminSalesInsightEmpty(error?.message || 'Nie udało się wczytać sprzedaży zespołu.');
  }
}

async function getUserBlockState(uid){
  if(!db || !uid) return false;

  try{
    const doc = await db.collection(USERS_COLLECTION).doc(uid).get();
    return Boolean(doc.exists && doc.data()?.blocked);
  }catch(error){
    console.error('User block check error', error);
    return false;
  }
}

function bindCategoryCard(id, handler){
  const element = document.getElementById(id);
  if(!element) return;
  element.addEventListener('click', handler);
}

async function openMarkaWlasnaCard(slug){
  const normalizedSlug = String(slug || '').trim();
  if(!normalizedSlug) return;

  await ensureMarkaWlasnaCardsLoaded();

  const card = getMarkaWlasnaCardConfig(normalizedSlug);
  const container = getMarkaWlasnaCardContainer(normalizedSlug);
  const source = getMarkaWlasnaCardSource(normalizedSlug);
  const cardId = `marka-wlasna-card-${normalizedSlug}`;
  const categoryName = String(card?.title || normalizedSlug)
    .trim()
    .replace(/\s+/g, '_');

  if(!container){
    console.error('Marka Wlasna container missing', normalizedSlug);
    return;
  }

  setActiveCard(cardId);
  setImageBase(imageBaseUrlMarkaWlasna);
  prepareCategoryView(container, categoryName || 'Marka_Wlasna', `${normalizedSlug}_marka_wlasna`);

  if(!isConfiguredDataSource(source)){
    isLoading = false;
    if(window.MARKA_WLASNA_LAYOUT?.renderPendingState){
      window.MARKA_WLASNA_LAYOUT.renderPendingState(container, card);
    }
    return;
  }

  try{
    const rows = await loadDataSourceRows(source);
    const normalizedRows = (Array.isArray(rows) ? rows : []).map(row => (
      row && typeof row === 'object' ? { ...row } : {}
    ));
    setFullData(normalizedRows);
    isLoading = false;
    render();
  }catch(error){
    console.error(`Marka Wlasna - ${normalizedSlug} load error`, error);
    isLoading = false;
    renderCategoryLoadError(container, error?.message || `Nie udało się wczytać danych dla kafelka ${card?.title || normalizedSlug}.`);
  }
}

function initializeUiBindings(){
  if(window.__wfBindingsReady) return;
  window.__wfBindingsReady = true;
  if(window.MARKA_WLASNA_LAYOUT?.ensureLayout){
    window.MARKA_WLASNA_LAYOUT.ensureLayout();
  }

  if(loginBtn){
    loginBtn.addEventListener('click', async () => {
      if(!authReady || !auth){
        setLoginError('Firebase Auth nie jest gotowy.');
        return;
      }
      setLoginError('');
      lastLoginEmail = loginEmail.value.trim();
      lastLoginPassword = loginPassword.value;
      try{
        await auth.signInWithEmailAndPassword(lastLoginEmail, lastLoginPassword);
      }catch(error){
        console.error('Login error', error);
        setLoginError(getFriendlyAuthErrorMessage(error));
      }
    });
  }

  if(logoutBtn){
    logoutBtn.addEventListener('click', async () => {
      try{
        if(auth){
          await auth.signOut();
        }
      }catch(error){
        console.error('Logout error', error);
      }finally{
        hidePersonalSalesInsight();
        resetAppState();
        showLogin();
      }
    });
  }

  if(loginPassword){
    loginPassword.addEventListener('keydown', event => {
      if(event.key === 'Enter' && loginBtn){
        loginBtn.click();
      }
    });
  }

  if(appBrandLink){
    appBrandLink.addEventListener('click', event => {
      if(auth && auth.currentUser){
        event.preventDefault();
        showApp();
        resetAppState();
        window.scrollTo({ top: 0, behavior: 'smooth' });
      }
    });
  }

  if(modalClose){
    modalClose.addEventListener('click', () => {
      if(securityModal) securityModal.classList.add('hidden');
    });
  }

  if(catalogCoverInput){
    catalogCoverInput.addEventListener('change', async () => {
      const file = catalogCoverInput.files?.[0];
      catalogCoverDataUrl = file ? await fileToDataUrl(file) : null;
    });
  }

  if(catalogExcelInput){
    catalogExcelInput.addEventListener('change', async () => {
      const file = catalogExcelInput.files?.[0];
      if(!file){
        if(!applyAdminCatalogPriceMapToCatalogState()){
          resetCatalogColumnOptions();
        }
        return;
      }
      if(priceWarningModal) priceWarningModal.classList.remove('hidden');
      try{
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        catalogPriceRows = rows;
        populateCatalogColumnOptions(rows);
        catalogPriceMap = buildPriceMap(rows, getCatalogColumnConfig());
      }catch(error){
        console.error('Catalog price import error', error);
        catalogPriceMap = null;
        catalogPriceRows = null;
      }
    });
  }

  if(catalogIndexColSelect){
    catalogIndexColSelect.addEventListener('change', () => {
      if(catalogPriceRows){
        catalogPriceMap = buildPriceMap(catalogPriceRows, getCatalogColumnConfig());
      }
    });
  }

  if(catalogPriceColSelect){
    catalogPriceColSelect.addEventListener('change', () => {
      if(catalogPriceRows){
        catalogPriceMap = buildPriceMap(catalogPriceRows, getCatalogColumnConfig());
      }
    });
  }

  if(catalogSaveBtn){
    catalogSaveBtn.addEventListener('click', buildAndSaveCatalog);
  }

  if(catalogCancelBtn){
    catalogCancelBtn.addEventListener('click', closeCatalogModal);
  }

  if(priceWarningClose){
    priceWarningClose.addEventListener('click', () => {
      if(priceWarningModal) priceWarningModal.classList.add('hidden');
    });
  }

  if(passwordCancel){
    passwordCancel.addEventListener('click', () => {
      if(passwordModal) passwordModal.classList.add('hidden');
    });
  }

  if(worldFoodBtn){
    worldFoodBtn.addEventListener('click', () => {
      const activeTab = document.querySelector('.tab.active')?.dataset.tab || 'rumunia';
      activateTab(activeTab);
    });
  }

  if(markaWlasnaBtn){
    markaWlasnaBtn.addEventListener('click', () => {
      void showMarkaWlasnaView();
    });
  }

  if(salesReportsBtn){
    salesReportsBtn.addEventListener('click', () => {
      openReportsView();
    });
  }

  if(tasksBtn){
    tasksBtn.addEventListener('click', () => {
      showTasksView();
    });
  }

  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      activateTab(tab.dataset.tab);
    });
  });

  bindCategoryCard('rumunia-slodycze', async () => {
    setActiveCard('rumunia-slodycze');
    setImageBase(imageBaseUrlRumunia);
    prepareCategoryView(slodyczeContainer, 'Slodycze', 'slodycze');
    setFullData(await loadDataSourceRows({ type: 'json', url: jsonUrl }));
    isLoading = false;
    render();
  });

  bindCategoryCard('rumunia-mieso-wedliny', async () => {
    setActiveCard('rumunia-mieso-wedliny');
    setImageBase(imageBaseUrlRumunia);
    prepareCategoryView(miesoContainer, 'Mieso_i_wedliny', 'mieso_i_wedliny');
    setFullData(await loadDataSourceRows({ type: 'json', url: jsonUrlMieso }));
    isLoading = false;
    render();
  });

  bindCategoryCard('rumunia-nabial', async () => {
    setActiveCard('rumunia-nabial');
    setImageBase(imageBaseUrlRumunia);
    prepareCategoryView(nabialContainer, 'Nabial', 'nabial');
    setFullData(await loadDataSourceRows({ type: 'json', url: jsonUrlNabial }));
    isLoading = false;
    render();
  });

  bindCategoryCard('rumunia-napoje', async () => {
    setActiveCard('rumunia-napoje');
    setImageBase(imageBaseUrlRumunia);
    prepareCategoryView(napojeContainer, 'Napoje', 'napoje');
    setFullData(await loadDataSourceRows({ type: 'json', url: jsonUrlNapoje }));
    isLoading = false;
    render();
  });

  bindCategoryCard('rumunia-przyprawy-proszek', async () => {
    setActiveCard('rumunia-przyprawy-proszek');
    setImageBase(imageBaseUrlRumunia);
    prepareCategoryView(przyprawyProszekContainer, 'Przyprawy_i_dodatki_w_proszku', 'przyprawy_i_dodatki_w_proszku');
    setFullData(await loadDataSourceRows({ type: 'json', url: jsonUrlPrzyprawyProszek }));
    isLoading = false;
    render();
  });

  bindCategoryCard('rumunia-puszki-sloiki', async () => {
    setActiveCard('rumunia-puszki-sloiki');
    setImageBase(imageBaseUrlRumunia);
    prepareCategoryView(puszkiSloikiContainer, 'Puszki_i_sloiki', 'puszki_i_sloiki');
    setFullData(await loadDataSourceRows({ type: 'json', url: jsonUrlPuszkiSloiki }));
    isLoading = false;
    render();
  });

  bindCategoryCard('rumunia-produkty-podstawowe', async () => {
    setActiveCard('rumunia-produkty-podstawowe');
    setImageBase(imageBaseUrlRumunia);
    prepareCategoryView(produktyPodstawoweContainer, 'Produkty_podstawowe', 'produkty_podstawowe');
    setFullData(await loadDataSourceRows({ type: 'json', url: jsonUrlProduktyPodstawowe }));
    isLoading = false;
    render();
  });

  bindCategoryCard('rumunia-kawy-herbaty', async () => {
    setActiveCard('rumunia-kawy-herbaty');
    setImageBase(imageBaseUrlRumunia);
    prepareCategoryView(kawyHerbatyContainer, 'Kawy_i_herbaty', 'kawy_i_herbaty');
    setFullData(await loadDataSourceRows({ type: 'json', url: jsonUrlKawyHerbaty }));
    isLoading = false;
    render();
  });

  bindCategoryCard('rumunia-top-rumunia', async () => {
    setActiveCard('rumunia-top-rumunia');
    setImageBase(imageBaseUrlRumunia);
    prepareCategoryView(topRumuniaContainer, 'Top_Rumunia', 'top_rumunia');
    setFullData(await loadDataSourceRows({ type: 'json', url: jsonUrlTopRumunia }));
    isLoading = false;
    render();
  });

  bindCategoryCard('ukraina-slodycze', async () => {
    setActiveCard('ukraina-slodycze');
    setImageBase(imageBaseUrlUkraina);
    prepareCategoryView(ukrainaSlodyczeContainer, 'Slodycze_Ukraina', 'slodycze_ukraina');
    setFullData(await loadDataSourceRows({ type: 'json', url: jsonUrlSlodyczeUkraina }));
    isLoading = false;
    render();
  });

  bindCategoryCard('ukraina-mieso-wedliny', async () => {
    setActiveCard('ukraina-mieso-wedliny');
    setImageBase(imageBaseUrlUkraina);
    prepareCategoryView(ukrainaMiesoContainer, 'Mieso_i_wedliny_Ukraina', 'mieso_i_wedliny_ukraina');
    setFullData(await loadDataSourceRows({ type: 'json', url: jsonUrlMiesoUkraina }));
    isLoading = false;
    render();
  });

  bindCategoryCard('ukraina-kawy-herbaty', async () => {
    setActiveCard('ukraina-kawy-herbaty');
    setImageBase(imageBaseUrlUkraina);
    prepareCategoryView(ukrainaKawyContainer, 'Kawy_i_herbaty_Ukraina', 'kawy_i_herbaty_ukraina');
    setFullData(await loadDataSourceRows({ type: 'json', url: jsonUrlKawyUkraina }));
    isLoading = false;
    render();
  });

  bindCategoryCard('ukraina-puszki-sloiki', async () => {
    setActiveCard('ukraina-puszki-sloiki');
    setImageBase(imageBaseUrlUkraina);
    prepareCategoryView(ukrainaPuszkiContainer, 'Puszki_i_sloiki_Ukraina', 'puszki_i_sloiki_ukraina');
    setFullData(await loadXlsxAsJson(xlsxUrlPuszkiUkraina));
    isLoading = false;
    render();
  });

  bindCategoryCard('ukraina-napoje', async () => {
    setActiveCard('ukraina-napoje');
    setImageBase(imageBaseUrlUkraina);
    prepareCategoryView(ukrainaNapojeContainer, 'Napoje_Ukraina', 'napoje_ukraina');
    setFullData(await loadDataSourceRows({ type: 'json', url: jsonUrlNapojeUkraina }));
    isLoading = false;
    render();
  });

  bindCategoryCard('ukraina-przyprawy-proszek', async () => {
    setActiveCard('ukraina-przyprawy-proszek');
    setImageBase(imageBaseUrlUkraina);
    prepareCategoryView(ukrainaPrzyprawyContainer, 'Przyprawy_Ukraina', 'przyprawy_ukraina');
    setFullData(await loadDataSourceRows({ type: 'json', url: jsonUrlPrzyprawyUkraina }));
    isLoading = false;
    render();
  });

  bindCategoryCard('ukraina-dodatki-do-potraw', async () => {
    if(!isConfiguredDataSource(ukrainaDodatkiDoPotrawSource)) return;
    setActiveCard('ukraina-dodatki-do-potraw');
    setImageBase(imageBaseUrlUkraina);
    prepareCategoryView(ukrainaDodatkiContainer, 'Dodatki_do_potraw_Ukraina', 'dodatki_do_potraw_ukraina');
    setFullData(await loadDataSourceRows(ukrainaDodatkiDoPotrawSource));
    isLoading = false;
    render();
  });

  bindCategoryCard('ukraina-top-ukraina', async () => {
    setActiveCard('ukraina-top-ukraina');
    setImageBase(imageBaseUrlUkraina);
    prepareCategoryView(ukrainaTopContainer, 'Top_Ukraina', 'top_ukraina');
    setFullData(await loadDataSourceRows({ type: 'json', url: jsonUrlTopUkraina }));
    isLoading = false;
    render();
  });

  if(importDaneCard){
    importDaneCard.addEventListener('click', async () => {
      setActiveCard('import-dane');
      setImageBase(imageBaseUrlRumunia);
      prepareCategoryView(importDaneContainer, 'Import_dane', 'import_dane');
      const result = await loadImportDane();
      setFullData(result.data, result.columns);
      isLoading = false;
      render();
    });
  }

  if(reportsCard){
    reportsCard.addEventListener('click', () => {
      openReportsView();
    });
  }
}

if(typeof firebase !== 'undefined'){
  if(!firebase.apps.length){
    firebase.initializeApp(firebaseConfig);
  }
  auth = firebase.auth();
  db = typeof firebase.firestore === 'function' ? firebase.firestore() : null;
  storage = typeof firebase.storage === 'function'
    ? firebase.app().storage(REPORTS_STORAGE_BUCKET)
    : null;
  mediaStorage = typeof firebase.storage === 'function'
    ? firebase.app().storage(MEDIA_STORAGE_BUCKET)
    : null;
  functionsService = typeof firebase.functions === 'function'
    ? firebase.app().functions(FUNCTIONS_REGION)
    : null;
  authReady = true;

  auth.onAuthStateChanged(async user => {
    const currentVersion = ++authStateVersion;

    if(isQuickOrderPublicMode()){
      setAuthBootState(false);
      if(loginView) loginView.classList.add('hidden');
      if(appView) appView.classList.add('hidden');
      return;
    }

    if(user){
      const isBlocked = await getUserBlockState(user.uid);
      if(currentVersion !== authStateVersion) return;

      if(isBlocked){
        setLoginError('To konto jest zablokowane.');
        try{
          await auth.signOut();
        }catch(error){
          console.error('Blocked user sign out error', error);
        }
        return;
      }

      const isAdmin = await hasAdminAccess(user, { forceRefresh: false });
      if(currentVersion !== authStateVersion) return;

      toggleAdminPanel(user.email || '', isAdmin);
      resetAppState();
      setLoginError('');
      showApp();

      if(isAdmin){
        void ensureAdminCatalogPriceCache({ forceRefresh: true });
        await loadAdminSalesInsight(currentVersion);
      }else{
        await loadPersonalSalesInsightForUser(user, currentVersion);
      }
    }else{
      toggleAdminPanel('', false);
      hidePersonalSalesInsight();
      resetAppState();
      showLogin();
    }
  });
}else{
  setLoginError('Firebase nie załadował się. Sprawdź połączenie.');
}

syncCatalogAdminPriceTools(false);
initializeUiBindings();
