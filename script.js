const jsonUrl =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Slodycze%20Ranking%20Rumunia.json';
if(window.__wfScriptLoaded){
  console.warn('script.js loaded more than once; duplicate bindings will be skipped.');
}
window.__wfScriptLoaded = true;
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
const imageBaseUrlRumunia =
  'https://firebasestorage.googleapis.com/v0/b/pdf-creator-f7a8b.firebasestorage.app/o/zdjecia%20-%20World%20food%2F';
const imageBaseUrlUkraina =
  'https://firebasestorage.googleapis.com/v0/b/pdf-creator-f7a8b.firebasestorage.app/o/zdjecia%20-%20World%20food%2FUkraina%2F';
const maspoLogoUrl =
  'https://firebasestorage.googleapis.com/v0/b/pdf-creator-f7a8b.firebasestorage.app/o/szablony%20maspo%2Fmaspo%20logo.png?alt=media&token=a5786f0b-19a7-4c01-a5bb-4b8d0cc9f583';
let currentImageBaseUrl = imageBaseUrlRumunia;
window.imageBaseUrl = currentImageBaseUrl;
window.imageBaseUrlRumunia = imageBaseUrlRumunia;
window.imageBaseUrlUkraina = imageBaseUrlUkraina;

const firebaseConfig = {
  apiKey: 'AIzaSyDwzVRS5W2lklGMLcZJn-YPCK9OtBQZ7bI',
  authDomain: 'pdf-creator-f7a8b.firebaseapp.com',
  projectId: 'pdf-creator-f7a8b',
  appId: '1:606744201676:web:6f8c1b2c323fbaf6f3b569'
};

const ADMIN_EMAILS = ['admin@admin.info', 'admin@admin.com'];
const USERS_COLLECTION = 'app_users';
const SALES_REPORTS_ROOT = 'raporty - maspo';
const STORAGE_BUCKET = 'gs://pdf-creator-f7a8b.firebasestorage.app';
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
const salesReportsBtn = document.getElementById('sales-reports-btn');
const worldFoodSubnav = document.getElementById('world-food-subnav');

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
const reportsBannerActions = document.getElementById('reports-banner-actions');
const ukrainaSlodyczeContainer = document.getElementById('ukraina-slodycze-content');
const ukrainaMiesoContainer = document.getElementById('ukraina-mieso-wedliny-content');
const ukrainaKawyContainer = document.getElementById('ukraina-kawy-herbaty-content');
const ukrainaPuszkiContainer = document.getElementById('ukraina-puszki-sloiki-content');
const ukrainaNapojeContainer = document.getElementById('ukraina-napoje-content');
const ukrainaPrzyprawyContainer = document.getElementById('ukraina-przyprawy-proszek-content');
const reportsCard = document.getElementById('reports-card');

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
let catalogPriceMap = null;
let catalogPriceRows = null;
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
let weeklySalesSourceRows = [];
let weeklySalesGeneratedRows = [];
let weeklySalesImportFile = '';
let weeklySalesSelectedRepresentative = '';
let weeklySalesSelectedCustomer = '';
let weeklySalesSelectedProducer = '';
let weeklySalesSelectedWeek = '';
let weeklySalesOnlyLastWeek250 = false;
let weeklySalesRepComparison = false;
let weeklySalesSortOrder = 'sales-desc';
let detailedSalesSourceRows = [];
let detailedSalesGeneratedRows = [];
let detailedSalesImportFile = '';
let detailedSalesSelectedCustomer = '';
let detailedSalesComparisonCustomer = '';
let detailedSalesSelectedGroups = [];
let detailedSalesSortOrder = 'sales-desc';
let detailedSalesWeeksLimit = 1;
let detailedSalesLoadedWeeksLimit = 0;
let detailedSalesLoadedReportsCount = 0;
let detailedSalesAvailableReportsCount = 0;
let detailedSalesUpdatedAt = '';
let detailedSalesLoading = false;
let detailedSalesErrorMessage = '';
let detailedSalesTopPurchasedRows = [];
let detailedSalesTopMissingRows = [];
let detailedSalesTopSummary = null;
let detailedSalesTopComparisonLoading = false;
let detailedSalesTopComparisonError = '';
let detailedSalesTopComparisonRequestId = 0;
let detailedSalesTopPurchasedFilters = { producer:'' };
let detailedSalesTopMissingFilters = { producer:'' };
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
let customerDiscountMapCache = null;
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
const IMAGE_PLACEHOLDER_SRC = 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==';
const IMAGE_EXTENSION_CACHE_KEY = 'wf-image-extension-cache-v1';
const MAX_IMAGE_EXTENSION_CACHE_ENTRIES = 3000;
const imageExtensionCache = createImageExtensionCache();
let lazyImageObserver = null;

let auth = null;
let db = null;
let storage = null;
let authReady = false;
let lastLoginEmail = '';
let lastLoginPassword = '';
let authStateVersion = 0;
let personalSalesDetailsLoading = false;
let phReportsAutoLoading = false;
let currentUserIsAdmin = false;
let personalSalesConfigCache = new Map();
let personalSalesFolderEntriesCache = new Map();
let detailedSalesParsedReportCache = new Map();
let detailedSalesParsedReportPromiseCache = new Map();
let detailedSalesGroupsDropdownOpen = false;

function clearDetailedSalesTopComparisonResultState(options = {}){
  const preserveCustomerSummaryRows = Boolean(options.preserveCustomerSummaryRows);
  detailedSalesTopPurchasedRows = [];
  detailedSalesTopMissingRows = [];
  detailedSalesTopSummary = null;
  detailedSalesTopComparisonError = '';
  detailedSalesTopPurchasedFilters = { producer:'' };
  detailedSalesTopMissingFilters = { producer:'' };
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
  const isWorldFood = mode !== 'reports';
  setWorldFoodExpanded(isWorldFood);
  if(salesReportsBtn){
    salesReportsBtn.classList.toggle('active', mode === 'reports');
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

function showApp(){
  setAuthBootState(false);
  if(loginView) loginView.classList.add('hidden');
  if(appView) appView.classList.remove('hidden');
  showSecurityModalOnce();
}

function showLogin(){
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
  catalogOptionsOverride = null;
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
  weeklySalesOnlyLastWeek250 = false;
  weeklySalesRepComparison = false;
  weeklySalesSortOrder = 'sales-desc';
  detailedSalesSourceRows = [];
  detailedSalesGeneratedRows = [];
  detailedSalesImportFile = '';
  detailedSalesSelectedCustomer = '';
  detailedSalesComparisonCustomer = '';
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
  personalSalesConfigCache = new Map();
  personalSalesFolderEntriesCache = new Map();
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
}

function setImageBase(base){
  currentImageBaseUrl = base;
  window.imageBaseUrl = currentImageBaseUrl;
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

  return Promise.race([
    auth.signInWithEmailAndPassword(email, password),
    new Promise((_, reject) => {
      setTimeout(() => {
        reject({ code: 'auth/network-timeout' });
      }, timeoutMs);
    })
  ]);
}

function hidePersonalSalesInsight(){
  if(personalSalesInsight) personalSalesInsight.classList.add('hidden');
  if(personalSalesInsightContent) personalSalesInsightContent.innerHTML = '';
}

function setPersonalSalesInsightMarkup(markup){
  if(!personalSalesInsight || !personalSalesInsightContent) return;
  personalSalesInsightContent.innerHTML = markup;
  personalSalesInsight.classList.remove('hidden');
}

function formatSalesNumber(value){
  return new Intl.NumberFormat('pl-PL', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(Number.isFinite(value) ? value : 0);
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

function normalizeDetailedSalesWeeksLimit(value){
  const parsed = Number.parseInt(String(value ?? '').trim(), 10);
  return Number.isFinite(parsed) && parsed >= 1 ? parsed : 1;
}

function stripFileExtension(filename){
  return String(filename || '').replace(/\.[^.]+$/, '');
}

function extractDetailedSalesReportLabel(reportEntry){
  const sourceName = String(
    reportEntry?.metadata?.customMetadata?.sourceFile
    || reportEntry?.item?.name
    || ''
  ).trim();
  const dateMatch = sourceName.match(/\b(\d{2}\.\d{2}\.\d{4})\b/);
  if(dateMatch){
    return dateMatch[1];
  }
  return stripFileExtension(sourceName) || 'Bez daty';
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

function renderPersonalSalesInsightLoading(config){
  setPersonalSalesInsightMarkup(`
    <div class="sales-insight-shell">
      <div class="sales-insight-main">
        <span class="sales-insight-eyebrow">Twoja sprzedaż</span>
        <h2>${escapeHtml(config.displayName)}</h2>
        <p>Ładuję porównanie dwóch ostatnich tygodni sprzedaży z najnowszego raportu.</p>
      </div>
      <div class="sales-insight-side">
        <div class="admin-user-empty">Pobieram dane raportu...</div>
      </div>
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

function renderPersonalSalesInsight(config, insight){
  const directionClass = insight.direction === 'up'
    ? 'is-up'
    : insight.direction === 'down'
      ? 'is-down'
      : 'is-flat';
  const statusClass = insight.direction === 'up'
    ? 'sales-insight-status--up'
    : insight.direction === 'down'
      ? 'sales-insight-status--down'
      : 'sales-insight-status--flat';
  const statusText = insight.direction === 'up'
    ? 'Sprzedaż rośnie względem poprzedniego tygodnia'
    : insight.direction === 'down'
      ? 'Sprzedaż spada względem poprzedniego tygodnia'
      : 'Sprzedaż utrzymała ten sam poziom';
  const statusIcon = insight.direction === 'up' ? '↗' : insight.direction === 'down' ? '↘' : '→';
  const maxValue = Math.max(insight.previousTotal, insight.lastTotal, 1);
  const previousWidth = Math.max((insight.previousTotal / maxValue) * 100, insight.previousTotal > 0 ? 8 : 0);
  const lastWidth = Math.max((insight.lastTotal / maxValue) * 100, insight.lastTotal > 0 ? 8 : 0);

  setPersonalSalesInsightMarkup(`
    <div class="sales-insight-shell ${directionClass}">
      <div class="sales-insight-main">
        <span class="sales-insight-eyebrow">Twoja sprzedaż</span>
        <h2>${escapeHtml(config.displayName)}</h2>
        <p>Porównanie dwóch ostatnich tygodni z najnowszego raportu sprzedaży przypisanego do Twojego konta.</p>
        <div class="sales-insight-status ${statusClass}">
          <span>${statusIcon}</span>
          <strong>${escapeHtml(statusText)}</strong>
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
          Ostatnia aktualizacja: ${escapeHtml(formatInsightDate(insight.updatedAt))}
        </div>
      </div>
    </div>
  `);
}

function findBestWeeklyRepresentativeMatch(representatives, config){
  const options = Array.isArray(representatives) ? representatives.filter(Boolean) : [];
  if(!options.length) return '';

  const normalizedCandidates = [
    String(config?.displayName || '').trim(),
    String(config?.storageFolder || '').trim()
  ]
    .map(value => normalizeProducerValue(value))
    .filter(Boolean);

  for(const candidate of normalizedCandidates){
    const exactMatch = options.find(rep => normalizeProducerValue(rep) === candidate);
    if(exactMatch) return exactMatch;
  }

  const candidateTokens = normalizedCandidates
    .flatMap(value => value.split(' '))
    .filter(token => token.length >= 3);

  let bestMatch = '';
  let bestScore = 0;
  options.forEach(rep => {
    const normalizedRep = normalizeProducerValue(rep);
    let score = 0;
    candidateTokens.forEach(token => {
      if(normalizedRep.includes(token)){
        score += 1;
      }
    });
    if(score > bestScore){
      bestScore = score;
      bestMatch = rep;
    }
  });

  if(bestScore > 0){
    return bestMatch;
  }

  return options.length === 1 ? options[0] : '';
}

async function loadWeeklySalesRowsFromFirebaseReport(config){
  const latestReport = await getLatestPersonalSalesReport(config);
  if(!latestReport){
    throw new Error('Brak pliku raportu sprzedaży.');
  }

  const buffer = await readReportEntryArrayBuffer(latestReport);
  const workbook = XLSX.read(buffer, { type: 'array' });
  const sheetName = workbook.SheetNames.includes('Export') ? 'Export' : workbook.SheetNames[0];
  if(!sheetName){
    throw new Error('Raport sprzedaży nie zawiera żadnego arkusza.');
  }

  const sheet = workbook.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  const rows = normalizeWeeklySalesRowsFromMatrix(matrix);
  if(!rows.length){
    throw new Error('Raport sprzedaży nie zawiera danych do wyświetlenia.');
  }

  return {
    rows,
    fileName: latestReport.item.name
  };
}

async function loadDetailedSalesRowsFromFirebaseReport(config, options = {}){
  const requestedWeeksLimit = normalizeDetailedSalesWeeksLimit(detailedSalesWeeksLimit);
  const { entries, totalCount, allEntries } = await getLatestPersonalDetailedSalesReports(config, requestedWeeksLimit, options);
  if(!entries.length){
    throw new Error('Brak szczegółowego raportu sprzedaży.');
  }

  const reportResults = await Promise.all(
    entries.map(async (reportEntry, reportIndex) => {
      const reportData = await getCachedDetailedSalesReportData(reportEntry);
      const rows = reportData.rows.map(row => ({
        ...row,
        reportLabel: reportData.reportLabel,
        reportOrder: reportIndex
      }));

      return {
        rows,
        fileName: reportData.fileName,
        updatedAt: reportData.updatedAt
      };
    })
  );

  warmDetailedSalesReportsCache(allEntries, Math.max(3, requestedWeeksLimit));

  const rows = reportResults.flatMap(report => report.rows);
  if(!rows.length){
    throw new Error('Szczegółowy raport sprzedaży nie zawiera danych do wyświetlenia.');
  }

  return {
    rows,
    fileName: reportResults[0].fileName,
    updatedAt: reportResults[0].updatedAt,
    loadedWeeksLimit: requestedWeeksLimit,
    loadedReportsCount: reportResults.length,
    availableReportsCount: totalCount
  };
}

async function applyPersonalWeeklySalesReport(config){
  const reportData = await loadWeeklySalesRowsFromFirebaseReport(config);
  weeklySalesSourceRows = reportData.rows;
  weeklySalesImportFile = reportData.fileName;

  const representatives = getWeeklySalesRepresentatives();
  weeklySalesSelectedRepresentative = findBestWeeklyRepresentativeMatch(representatives, config)
    || representatives[0]
    || '';
  weeklySalesSelectedWeek = '';
  weeklySalesOnlyLastWeek250 = false;
  weeklySalesRepComparison = false;
  weeklySalesSortOrder = 'sales-desc';

  await loadCustomerDiscountMap();
  generateWeeklySalesReport();
}

async function applyPersonalDetailedSalesReport(config, options = {}){
  const previousCustomer = detailedSalesSelectedCustomer;
  const previousComparisonCustomer = detailedSalesComparisonCustomer;
  const previousGroups = [...detailedSalesSelectedGroups];
  const reportData = await loadDetailedSalesRowsFromFirebaseReport(config, options);
  detailedSalesSourceRows = reportData.rows;
  detailedSalesImportFile = reportData.fileName;
  detailedSalesUpdatedAt = reportData.updatedAt;
  detailedSalesLoadedWeeksLimit = reportData.loadedWeeksLimit;
  detailedSalesLoadedReportsCount = reportData.loadedReportsCount;
  detailedSalesAvailableReportsCount = reportData.availableReportsCount;
  detailedSalesSelectedCustomer = reportData.rows.some(row => `${row.customerCode}|||${row.customerName}` === previousCustomer)
    ? previousCustomer
    : '';
  const customerOptions = getDetailedSalesCustomers();
  detailedSalesComparisonCustomer = customerOptions.some(customer => `${customer.code}|||${customer.name}` === previousComparisonCustomer)
    ? previousComparisonCustomer
    : (customerOptions[0] ? `${customerOptions[0].code}|||${customerOptions[0].name}` : '');
  detailedSalesErrorMessage = '';

  generateDetailedSalesReport();
  const availableGroups = new Set(getDetailedSalesGroups());
  detailedSalesSelectedGroups = previousGroups.filter(group => availableGroups.has(group));
  await generateDetailedSalesTopComparison({ suppressRender: true });
}

function normalizePersonalSalesConfig(source){
  if(!source) return null;
  const storageFolder = String(source.salesReportFolder || source.storageFolder || '').trim();
  if(!storageFolder) return null;
  const displayName = String(source.salesRepName || source.displayName || '').trim() || storageFolder;
  return { displayName, storageFolder };
}

async function getPersonalSalesConfig(uid, options = {}){
  if(!uid) return null;
  const forceRefresh = Boolean(options.forceRefresh);
  if(forceRefresh){
    personalSalesConfigCache.delete(uid);
  }
  if(personalSalesConfigCache.has(uid)){
    return personalSalesConfigCache.get(uid);
  }

  let config = null;
  if(db){
    try{
      const userDoc = await db.collection(USERS_COLLECTION).doc(uid).get();
      if(userDoc.exists){
        config = normalizePersonalSalesConfig(userDoc.data() || {});
      }
    }catch(error){
      console.error('Personal sales config read error', error);
    }
  }

  if(!config){
    config = normalizePersonalSalesConfig(PERSONAL_SALES_USERS[uid]);
  }

  personalSalesConfigCache.set(uid, config);
  return config;
}

async function ensurePhWeeklySalesDataLoaded(){
  if(isCurrentUserAdmin()) return;
  if(personalSalesDetailsLoading || phReportsAutoLoading) return;
  if(weeklySalesSourceRows.length) return;

  const user = auth?.currentUser;
  const config = await getPersonalSalesConfig(user?.uid);
  if(!config) return;

  phReportsAutoLoading = true;
  try{
    await applyPersonalWeeklySalesReport(config);
    renderReportsView();
  }catch(error){
    console.error('PH weekly sales auto-load error', error);
  }finally{
    phReportsAutoLoading = false;
    renderReportsView();
  }
}

async function ensurePhDetailedSalesDataLoaded(options = {}){
  const forceRefresh = Boolean(options.forceRefresh);
  if(isCurrentUserAdmin()) return;
  if(detailedSalesLoading) return;
  if(!forceRefresh && detailedSalesSourceRows.length && detailedSalesLoadedWeeksLimit === detailedSalesWeeksLimit){
    return;
  }

  const user = auth?.currentUser;
  const config = await getPersonalSalesConfig(user?.uid);
  if(!config){
    detailedSalesErrorMessage = 'Brak przypisanego raportu sprzedaży. Skontaktuj się z administratorem.';
    detailedSalesLoadedWeeksLimit = 0;
    detailedSalesLoadedReportsCount = 0;
    detailedSalesAvailableReportsCount = 0;
    renderReportsView();
    return;
  }

  detailedSalesLoading = true;
  detailedSalesErrorMessage = '';
  renderReportsView();

  try{
    await applyPersonalDetailedSalesReport(config, { forceRefresh });
  }catch(error){
    console.error('PH detailed sales auto-load error', error);
    detailedSalesSourceRows = [];
    detailedSalesGeneratedRows = [];
    detailedSalesImportFile = '';
    detailedSalesSelectedCustomer = '';
    detailedSalesComparisonCustomer = '';
    detailedSalesSelectedGroups = [];
    detailedSalesLoadedWeeksLimit = 0;
    detailedSalesLoadedReportsCount = 0;
    detailedSalesAvailableReportsCount = 0;
    detailedSalesUpdatedAt = '';
    detailedSalesErrorMessage = error?.message || 'Nie udało się wczytać szczegółowego raportu sprzedaży.';
    clearDetailedSalesTopComparisonResultState();
    detailedSalesTopComparisonLoading = false;
    detailedSalesTopComparisonRequestId += 1;
  }finally{
    detailedSalesLoading = false;
    renderReportsView();
  }
}

async function openPersonalSalesReportDetails(){
  if(personalSalesDetailsLoading) return;
  const user = auth?.currentUser;
  const config = await getPersonalSalesConfig(user?.uid, { forceRefresh: true });
  if(!config){
    alert('Brak przypisanego raportu sprzedaży. Skontaktuj się z administratorem.');
    return;
  }

  const detailsButton = personalSalesInsightContent?.querySelector('[data-personal-sales-details-btn="1"]');
  const initialLabel = detailsButton?.textContent || 'Pokaż szczegóły';

  personalSalesDetailsLoading = true;
  if(detailsButton){
    detailsButton.disabled = true;
    detailsButton.textContent = 'Ładowanie...';
  }

  try{
    reportsMode = 'weekly-sales';
    openReportsView();

    weeklySalesSourceRows = [];
    weeklySalesGeneratedRows = [];
    weeklySalesImportFile = '';
    weeklySalesSelectedRepresentative = '';
    weeklySalesSelectedCustomer = '';
    weeklySalesSelectedProducer = '';
    weeklySalesSelectedWeek = '';
    weeklySalesOnlyLastWeek250 = false;
    weeklySalesRepComparison = false;
    weeklySalesSortOrder = 'sales-desc';
    renderReportsView();

    await applyPersonalWeeklySalesReport(config);

    reportsMode = 'weekly-sales';
    renderReportsView();
  }catch(error){
    console.error('Personal sales details load error', error);
    weeklySalesSourceRows = [];
    weeklySalesGeneratedRows = [];
    weeklySalesImportFile = '';
    weeklySalesSelectedRepresentative = '';
    weeklySalesSelectedCustomer = '';
    weeklySalesSelectedProducer = '';
    weeklySalesSelectedWeek = '';
    weeklySalesOnlyLastWeek250 = false;
    weeklySalesRepComparison = false;
    weeklySalesSortOrder = 'sales-desc';
    reportsMode = 'weekly-sales';
    renderReportsView();
    alert(error?.message || 'Nie udało się wczytać raportu sprzedaży.');
  }finally{
    personalSalesDetailsLoading = false;
    if(detailsButton){
      detailsButton.disabled = false;
      detailsButton.textContent = initialLabel;
    }
  }
}

async function openPersonalSalesTopComparison(){
  if(personalSalesDetailsLoading) return;
  const user = auth?.currentUser;
  const config = await getPersonalSalesConfig(user?.uid, { forceRefresh: true });
  if(!config){
    alert('Brak przypisanego raportu sprzedaży. Skontaktuj się z administratorem.');
    return;
  }

  const comparisonButton = personalSalesInsightContent?.querySelector('[data-personal-sales-top-btn="1"]');
  const initialLabel = comparisonButton?.textContent || 'Top Rumunia';

  personalSalesDetailsLoading = true;
  if(comparisonButton){
    comparisonButton.disabled = true;
    comparisonButton.textContent = 'Ładowanie...';
  }

  try{
    reportsMode = 'top-rumunia-comparison';
    detailedSalesLoading = true;
    openReportsView();
    detailedSalesSourceRows = [];
    detailedSalesGeneratedRows = [];
    detailedSalesImportFile = '';
    detailedSalesSelectedCustomer = '';
    detailedSalesComparisonCustomer = '';
    detailedSalesSelectedGroups = [];
    detailedSalesLoadedWeeksLimit = 0;
    detailedSalesLoadedReportsCount = 0;
    detailedSalesAvailableReportsCount = 0;
    detailedSalesUpdatedAt = '';
    detailedSalesErrorMessage = '';
    clearDetailedSalesTopComparisonResultState();
    detailedSalesTopComparisonLoading = false;
    renderReportsView();

    await applyPersonalDetailedSalesReport(config);

    reportsMode = 'top-rumunia-comparison';
    detailedSalesLoading = false;
    renderReportsView();
  }catch(error){
    console.error('Personal sales Top Rumunia load error', error);
    detailedSalesSourceRows = [];
    detailedSalesGeneratedRows = [];
    detailedSalesImportFile = '';
    detailedSalesSelectedCustomer = '';
    detailedSalesComparisonCustomer = '';
    detailedSalesSelectedGroups = [];
    detailedSalesLoadedWeeksLimit = 0;
    detailedSalesLoadedReportsCount = 0;
    detailedSalesAvailableReportsCount = 0;
    detailedSalesUpdatedAt = '';
    detailedSalesErrorMessage = error?.message || 'Nie udało się wczytać raportu sprzedaży.';
    clearDetailedSalesTopComparisonResultState();
    detailedSalesTopComparisonLoading = false;
    reportsMode = 'top-rumunia-comparison';
    detailedSalesLoading = false;
    renderReportsView();
    alert(error?.message || 'Nie udało się wczytać raportu sprzedaży.');
  }finally{
    detailedSalesLoading = false;
    personalSalesDetailsLoading = false;
    if(comparisonButton){
      comparisonButton.disabled = false;
      comparisonButton.textContent = initialLabel;
    }
  }
}

async function getLatestPersonalSalesReport(config){
  const { weeklyEntry } = await listPersonalSalesReportEntries(config);
  return weeklyEntry || null;
}

function getPersonalSalesFolderPath(config){
  return `${SALES_REPORTS_ROOT}/${config.storageFolder}`;
}

function getReportEntryUpdatedAt(reportEntry){
  return String(reportEntry?.metadata?.updated || reportEntry?.metadata?.timeCreated || '').trim();
}

function getReportEntryCacheKey(reportEntry){
  const fullPath = String(reportEntry?.item?.fullPath || reportEntry?.item?.name || '').trim();
  return `${fullPath}::${getReportEntryUpdatedAt(reportEntry)}`;
}

async function listPersonalSalesReportEntries(config, options = {}){
  if(!storage){
    throw new Error('Firebase Storage nie jest dostępny.');
  }

  const forceRefresh = Boolean(options.forceRefresh);
  const folderPath = getPersonalSalesFolderPath(config);

  if(!forceRefresh && personalSalesFolderEntriesCache.has(folderPath)){
    return personalSalesFolderEntriesCache.get(folderPath);
  }

  const folderRef = storage.ref().child(folderPath);
  const listResult = await folderRef.listAll();
  const xlsxItems = listResult.items.filter(item => /\.xlsx$/i.test(item.name));

  if(!xlsxItems.length){
    const emptyResult = {
      allEntries: [],
      weeklyEntry: null,
      detailedEntries: []
    };
    personalSalesFolderEntriesCache.set(folderPath, emptyResult);
    return emptyResult;
  }

  const allEntries = await Promise.all(
    xlsxItems.map(async item => {
      const metadata = await item.getMetadata();
      const updatedAt = Date.parse(metadata.updated || metadata.timeCreated || '') || 0;
      return { item, metadata, updatedAt };
    })
  );

  allEntries.sort((a, b) => {
    if(b.updatedAt !== a.updatedAt) return b.updatedAt - a.updatedAt;
    return b.item.name.localeCompare(a.item.name, 'pl');
  });

  const result = {
    allEntries,
    weeklyEntry: allEntries.find(isWeeklySalesReportEntry) || null,
    detailedEntries: allEntries.filter(isSalesByIndexReportEntry)
  };
  personalSalesFolderEntriesCache.set(folderPath, result);
  return result;
}

async function getLatestPersonalDetailedSalesReports(config, limit = 1, options = {}){
  const { detailedEntries } = await listPersonalSalesReportEntries(config, options);
  return {
    entries: detailedEntries.slice(0, normalizeDetailedSalesWeeksLimit(limit)),
    totalCount: detailedEntries.length,
    allEntries: detailedEntries
  };
}

async function getCachedDetailedSalesReportData(reportEntry){
  const cacheKey = getReportEntryCacheKey(reportEntry);
  if(cacheKey && detailedSalesParsedReportCache.has(cacheKey)){
    return detailedSalesParsedReportCache.get(cacheKey);
  }
  if(cacheKey && detailedSalesParsedReportPromiseCache.has(cacheKey)){
    return detailedSalesParsedReportPromiseCache.get(cacheKey);
  }

  const loadPromise = (async () => {
    const buffer = await readReportEntryArrayBuffer(reportEntry);
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    if(!sheetName){
      throw new Error('Szczegółowy raport sprzedaży nie zawiera żadnego arkusza.');
    }

    const sheet = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const rows = normalizeDetailedSalesRowsFromMatrix(matrix);

    return {
      rows,
      fileName: reportEntry?.item?.name || '',
      updatedAt: getReportEntryUpdatedAt(reportEntry),
      reportLabel: extractDetailedSalesReportLabel(reportEntry)
    };
  })();

  if(cacheKey){
    detailedSalesParsedReportPromiseCache.set(cacheKey, loadPromise);
  }

  try{
    const data = await loadPromise;
    if(cacheKey){
      detailedSalesParsedReportCache.set(cacheKey, data);
    }
    return data;
  }finally{
    if(cacheKey){
      detailedSalesParsedReportPromiseCache.delete(cacheKey);
    }
  }
}

function warmDetailedSalesReportsCache(entries, limit = 3){
  const warmEntries = (entries || []).slice(0, Math.max(0, limit));
  if(!warmEntries.length) return;
  void Promise.allSettled(warmEntries.map(reportEntry => getCachedDetailedSalesReportData(reportEntry)));
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

function isSalesByIndexReportEntry(reportEntry){
  const customMetadata = reportEntry?.metadata?.customMetadata || {};
  const reportType = String(customMetadata.reportType || '').trim().toLowerCase();

  if(reportType === 'sales-by-index'){
    return true;
  }

  if(reportType === 'weekly-sales'){
    return false;
  }

  const fileName = String(reportEntry?.item?.name || '').trim().toLowerCase();
  return fileName.includes('per indeks');
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
    updatedAt: reportEntry?.metadata?.updated || reportEntry?.metadata?.timeCreated || ''
  };
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
  const hasFetchError = errorList.some(error => String(error?.message || '').toLowerCase().includes('failed to fetch'));
  if(hasFetchError){
    return 'Nie udało się pobrać raportu sprzedaży (błąd sieci/CORS). Odśwież stronę i spróbuj ponownie.';
  }
  return 'Nie udało się pobrać pliku raportu sprzedaży z Firebase Storage.';
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
    const downloadUrl = await reference.getDownloadURL();
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
  }catch(error){
    errors.push(error);
  }

  console.error('Personal sales report download attempts failed', errors);
  throw new Error(buildReportDownloadErrorMessage(errors));
}

async function readPersonalSalesInsight(reportEntry){
  try{
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
    const secondHeaderRow = Array.isArray(rows[1]) ? rows[1] : [];
    const salesColumns = secondHeaderRow
      .map((header, columnIndex) => ({ header, columnIndex }))
      .filter(item => normalizeHeaderKey(item.header).includes('sprzedaz waluta'));

    if(salesColumns.length >= 2){
      salesColumns.forEach(item => {
        const label = String(headerRow[item.columnIndex] ?? '').trim();
        if(label){
          weekLabels.push({ columnIndex: item.columnIndex, label });
        }
      });
    }else{
      for(let columnIndex = 3; columnIndex < headerRow.length; columnIndex += 1){
        const label = String(headerRow[columnIndex] ?? '').trim();
        if(label){
          weekLabels.push({ columnIndex, label });
        }
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
      fileName: reportEntry.item.name,
      updatedAt: reportEntry.metadata?.updated || reportEntry.metadata?.timeCreated || ''
    };
  }catch(parseError){
    const metadataInsight = readPersonalSalesInsightFromMetadata(reportEntry);
    if(metadataInsight){
      return metadataInsight;
    }
    throw parseError;
  }
}

async function loadPersonalSalesInsightForUser(user, version){
  const config = await getPersonalSalesConfig(user?.uid, { forceRefresh: true });
  if(version !== authStateVersion) return;

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
    renderPersonalSalesInsight(config, insight);
  }catch(error){
    console.error('Personal sales insight error', error);
    if(version !== authStateVersion) return;
    renderPersonalSalesInsightEmpty(config, error?.message || 'Nie udało się wczytać danych sprzedaży.');
  }
}

async function getUserBlockState(uid){
  if(!db || !uid) return false;

  try{
    const readStatePromise = db.collection(USERS_COLLECTION).doc(uid).get()
      .then(doc => Boolean(doc.exists && doc.data()?.blocked))
      .catch(error => {
        console.error('User block check error', error);
        return false;
      });
    return await withTimeout(readStatePromise, 5000, false);
  }catch(error){
    console.error('User block check error', error);
    return false;
  }
}

async function ensureUserProfileDoc(user){
  if(!db || !user?.uid || typeof firebase?.firestore !== 'function') return;

  try{
    const userRef = db.collection(USERS_COLLECTION).doc(user.uid);
    const userSnap = await userRef.get();
    const currentData = userSnap.exists ? (userSnap.data() || {}) : {};
    const payload = {
      uid: user.uid,
      email: user.email || '',
      lastLoginAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    };

    if(!String(currentData.source || '').trim()){
      payload.source = 'auth-login';
    }

    if(!userSnap.exists || !currentData.createdAt){
      payload.createdAt = firebase.firestore.FieldValue.serverTimestamp();
    }

    if(typeof currentData.blocked !== 'boolean'){
      payload.blocked = false;
    }

    await userRef.set(payload, { merge: true });
  }catch(error){
    console.error('Ensure user profile doc error', error);
  }
}

setAuthBootState(true, 'Sprawdzam sesję...');

if(typeof firebase !== 'undefined'){
  if(String(firebaseConfig.apiKey).startsWith('PASTE_')){
    showLogin();
    setLoginError('Uzupełnij firebaseConfig w script.js');
  }else{
    if(!firebase.apps.length){
      firebase.initializeApp(firebaseConfig);
    }
    auth = firebase.auth();
    if(typeof firebase.firestore === 'function'){
      db = firebase.firestore();
    }
    if(typeof firebase.storage === 'function'){
      storage = firebase.app().storage(STORAGE_BUCKET);
    }
    authReady = true;

    if(typeof window.__wfAuthUnsubscribe === 'function'){
      try{
        window.__wfAuthUnsubscribe();
      }catch(error){
        console.error('Previous auth listener cleanup error', error);
      }
    }

    window.__wfAuthUnsubscribe = auth.onAuthStateChanged(async user => {
      const currentVersion = ++authStateVersion;

      if(user){
        void ensureUserProfileDoc(user);
        const isBlocked = await getUserBlockState(user.uid);
        if(currentVersion !== authStateVersion){
          return;
        }
        if(isBlocked){
          setLoginError('To konto jest zablokowane.');
          try{
            await auth.signOut();
          }catch(error){
            console.error('Blocked user sign out error', error);
          }
          return;
        }
        showApp();
        toggleAdminPanel(user.email || '');
        resetAppState();
        setLoginError('');
        loadPersonalSalesInsightForUser(user, currentVersion);
      }else{
        toggleAdminPanel('');
        hidePersonalSalesInsight();
        resetAppState();
        showLogin();
      }
    });
  }
}else{
  showLogin();
  setLoginError('Firebase nie załadował się. Sprawdź połączenie.');
}

if(loginBtn && !window.__wfLoginHandlerBound){
  window.__wfLoginHandlerBound = true;
  loginBtn.addEventListener('click', async () => {
    if(!authReady || !auth){
      setLoginError('Firebase Auth nie jest gotowy.');
      return;
    }
    if(typeof navigator !== 'undefined' && navigator.onLine === false){
      setLoginError('Brak internetu. Sprawdź hotspot i spróbuj ponownie.');
      return;
    }
    setLoginError('Logowanie...');
    lastLoginEmail = loginEmail.value.trim();
    lastLoginPassword = loginPassword.value;
    const initialLabel = loginBtn.textContent;
    loginBtn.disabled = true;
    loginBtn.textContent = 'Logowanie...';
    setAuthBootState(true, 'Logowanie...');
    try{
      await signInWithTimeout(loginEmail.value.trim(), loginPassword.value);
    }catch(err){
      console.error('LOGIN ERROR:', err.code, err.message, err);
      setAuthBootState(false);
      setLoginError(getFriendlyAuthErrorMessage(err));
    }finally{
      loginBtn.disabled = false;
      loginBtn.textContent = initialLabel;
    }
  });
}


if(logoutBtn && !window.__wfLogoutHandlerBound){
  window.__wfLogoutHandlerBound = true;
  logoutBtn.addEventListener('click', async () => {
    try{
      if(auth){
        await auth.signOut();
      }
    }catch(e){
      console.error(e);
    }finally{
      toggleAdminPanel('');
      resetAppState();
      showLogin();
    }
  });
}

if(loginPassword && !window.__wfPasswordEnterHandlerBound){
  window.__wfPasswordEnterHandlerBound = true;
  loginPassword.addEventListener('keydown', (e) => {
    if(e.key === 'Enter' && loginBtn){
      loginBtn.click();
    }
  });
}

if(appBrandLink && !window.__wfBrandHandlerBound){
  window.__wfBrandHandlerBound = true;
  appBrandLink.addEventListener('click', (e) => {
    if(auth && auth.currentUser){
      e.preventDefault();
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
    if(!file){
      catalogCoverDataUrl = null;
      return;
    }
    catalogCoverDataUrl = await fileToDataUrl(file);
  });
}

if(catalogExcelInput){
  catalogExcelInput.addEventListener('change', async () => {
    const file = catalogExcelInput.files?.[0];
    if(!file){
      catalogPriceMap = null;
      catalogPriceRows = null;
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
    }catch(e){
      console.error(e);
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

// Admin panel jest osobną stroną (admin.html) — brak modala na stronie głównej.

if(passwordCancel){
  passwordCancel.addEventListener('click', () => {
    if(passwordModal) passwordModal.classList.add('hidden');
  });
}

if(worldFoodBtn && !window.__wfWorldFoodHandlerBound){
  window.__wfWorldFoodHandlerBound = true;
  worldFoodBtn.addEventListener('click', () => {
    const activeTab = document.querySelector('.tab.active')?.dataset.tab || 'rumunia';
    activateTab(activeTab);
  });
}

if(salesReportsBtn && !window.__wfSalesReportsHandlerBound){
  window.__wfSalesReportsHandlerBound = true;
  salesReportsBtn.addEventListener('click', () => {
    openReportsView();
  });
}

if(!window.__wfDetailedGroupsOutsideClickBound){
  window.__wfDetailedGroupsOutsideClickBound = true;
  document.addEventListener('click', handleDetailedSalesGroupsDropdownOutsideClick);
}

// OBSŁUGA ZAKŁADEK
if(!window.__wfTabsHandlerBound){
  window.__wfTabsHandlerBound = true;
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      activateTab(tab.dataset.tab);
    });
  });
}

// KLIK KAFELKA "SŁODYCZE"
document.getElementById('rumunia-slodycze').addEventListener('click', async () => {
  setActiveCard('rumunia-slodycze');
  setImageBase(imageBaseUrlRumunia);
  prepareCategoryView(slodyczeContainer, 'Slodycze', 'slodycze');
  const res = await fetch(jsonUrl);
  setFullData(await res.json());
  isLoading = false;
  render();
});

// KLIK KAFELKA "SŁODYCZE" - UKRAINA
const ukrainaSlodyczeCard = document.getElementById('ukraina-slodycze');
if(ukrainaSlodyczeCard){
  ukrainaSlodyczeCard.addEventListener('click', async () => {
    setActiveCard('ukraina-slodycze');
    setImageBase(imageBaseUrlUkraina);
    prepareCategoryView(ukrainaSlodyczeContainer, 'Slodycze_Ukraina', 'slodycze_ukraina');
    const res = await fetch(jsonUrlSlodyczeUkraina);
    setFullData(await res.json());
    isLoading = false;
    render();
  });
}

// KLIK KAFELKA "MIĘSO I WĘDLINY" - UKRAINA
const ukrainaMiesoCard = document.getElementById('ukraina-mieso-wedliny');
if(ukrainaMiesoCard){
  ukrainaMiesoCard.addEventListener('click', async () => {
    setActiveCard('ukraina-mieso-wedliny');
    setImageBase(imageBaseUrlUkraina);
    prepareCategoryView(ukrainaMiesoContainer, 'Mieso_i_wedliny_Ukraina', 'mieso_i_wedliny_ukraina');
    const res = await fetch(jsonUrlMiesoUkraina);
    setFullData(await res.json());
    isLoading = false;
    render();
  });
}

// KLIK KAFELKA "KAWY I HERBATY" - UKRAINA
const ukrainaKawyCard = document.getElementById('ukraina-kawy-herbaty');
if(ukrainaKawyCard){
  ukrainaKawyCard.addEventListener('click', async () => {
    setActiveCard('ukraina-kawy-herbaty');
    setImageBase(imageBaseUrlUkraina);
    prepareCategoryView(ukrainaKawyContainer, 'Kawy_i_herbaty_Ukraina', 'kawy_i_herbaty_ukraina');
    const res = await fetch(jsonUrlKawyUkraina);
    setFullData(await res.json());
    isLoading = false;
    render();
  });
}

// KLIK KAFELKA "PUSZKI I SŁOIKI" - UKRAINA
const ukrainaPuszkiCard = document.getElementById('ukraina-puszki-sloiki');
if(ukrainaPuszkiCard){
  ukrainaPuszkiCard.addEventListener('click', async () => {
    setActiveCard('ukraina-puszki-sloiki');
    setImageBase(imageBaseUrlUkraina);
    prepareCategoryView(ukrainaPuszkiContainer, 'Puszki_i_sloiki_Ukraina', 'puszki_i_sloiki_ukraina');
    setFullData(await loadXlsxAsJson(xlsxUrlPuszkiUkraina));
    isLoading = false;
    render();
  });
}

// KLIK KAFELKA "NAPOJE" - UKRAINA
const ukrainaNapojeCard = document.getElementById('ukraina-napoje');
if(ukrainaNapojeCard){
  ukrainaNapojeCard.addEventListener('click', async () => {
    setActiveCard('ukraina-napoje');
    setImageBase(imageBaseUrlUkraina);
    prepareCategoryView(ukrainaNapojeContainer, 'Napoje_Ukraina', 'napoje_ukraina');
    const res = await fetch(jsonUrlNapojeUkraina);
    setFullData(await res.json());
    isLoading = false;
    render();
  });
}

// KLIK KAFELKA "PRZYPRAWY I DODATKI W PROSZKU" - UKRAINA
const ukrainaPrzyprawyCard = document.getElementById('ukraina-przyprawy-proszek');
if(ukrainaPrzyprawyCard){
  ukrainaPrzyprawyCard.addEventListener('click', async () => {
    setActiveCard('ukraina-przyprawy-proszek');
    setImageBase(imageBaseUrlUkraina);
    prepareCategoryView(ukrainaPrzyprawyContainer, 'Przyprawy_Ukraina', 'przyprawy_ukraina');
    const res = await fetch(jsonUrlPrzyprawyUkraina);
    setFullData(await res.json());
    isLoading = false;
    render();
  });
}

// KLIK KAFELKA "MIĘSO I WĘDLINY"
document.getElementById('rumunia-mieso-wedliny').addEventListener('click', async () => {
  setActiveCard('rumunia-mieso-wedliny');
  setImageBase(imageBaseUrlRumunia);
  prepareCategoryView(miesoContainer, 'Mieso_i_wedliny', 'mieso_i_wedliny');
  const res = await fetch(jsonUrlMieso);
  setFullData(await res.json());
  isLoading = false;
  render();
});

// KLIK KAFELKA "NABIAŁ"
document.getElementById('rumunia-nabial').addEventListener('click', async () => {
  setActiveCard('rumunia-nabial');
  setImageBase(imageBaseUrlRumunia);
  prepareCategoryView(nabialContainer, 'Nabial', 'nabial');
  const res = await fetch(jsonUrlNabial);
  setFullData(await res.json());
  isLoading = false;
  render();
});

// KLIK KAFELKA "NAPOJE"
document.getElementById('rumunia-napoje').addEventListener('click', async () => {
  setActiveCard('rumunia-napoje');
  setImageBase(imageBaseUrlRumunia);
  prepareCategoryView(napojeContainer, 'Napoje', 'napoje');
  const res = await fetch(jsonUrlNapoje);
  setFullData(await res.json());
  isLoading = false;
  render();
});

// KLIK KAFELKA "PRZYPRAWY I DODATKI W PROSZKU"
document.getElementById('rumunia-przyprawy-proszek').addEventListener('click', async () => {
  setActiveCard('rumunia-przyprawy-proszek');
  setImageBase(imageBaseUrlRumunia);
  prepareCategoryView(przyprawyProszekContainer, 'Przyprawy_i_dodatki_w_proszku', 'przyprawy_i_dodatki_w_proszku');
  const res = await fetch(jsonUrlPrzyprawyProszek);
  setFullData(await res.json());
  isLoading = false;
  render();
});

// KLIK KAFELKA "PUSZKI I SŁOIKI"
document.getElementById('rumunia-puszki-sloiki').addEventListener('click', async () => {
  setActiveCard('rumunia-puszki-sloiki');
  setImageBase(imageBaseUrlRumunia);
  prepareCategoryView(puszkiSloikiContainer, 'Puszki_i_sloiki', 'puszki_i_sloiki');
  const res = await fetch(jsonUrlPuszkiSloiki);
  setFullData(await res.json());
  isLoading = false;
  render();
});

// KLIK KAFELKA "PRODUKTY PODSTAWOWE"
document.getElementById('rumunia-produkty-podstawowe').addEventListener('click', async () => {
  setActiveCard('rumunia-produkty-podstawowe');
  setImageBase(imageBaseUrlRumunia);
  prepareCategoryView(produktyPodstawoweContainer, 'Produkty_podstawowe', 'produkty_podstawowe');
  const res = await fetch(jsonUrlProduktyPodstawowe);
  setFullData(await res.json());
  isLoading = false;
  render();
});

// KLIK KAFELKA "KAWY I HERBATY"
document.getElementById('rumunia-kawy-herbaty').addEventListener('click', async () => {
  setActiveCard('rumunia-kawy-herbaty');
  setImageBase(imageBaseUrlRumunia);
  prepareCategoryView(kawyHerbatyContainer, 'Kawy_i_herbaty', 'kawy_i_herbaty');
  const res = await fetch(jsonUrlKawyHerbaty);
  setFullData(await res.json());
  isLoading = false;
  render();
});

// KLIK KAFELKA "TOP RUMUNIA"
document.getElementById('rumunia-top-rumunia').addEventListener('click', async () => {
  setActiveCard('rumunia-top-rumunia');
  setImageBase(imageBaseUrlRumunia);
  prepareCategoryView(topRumuniaContainer, 'Top_Rumunia', 'top_rumunia');
  const res = await fetch(jsonUrlTopRumunia);
  setFullData(await res.json());
  isLoading = false;
  render();
});

const importDaneCard = document.getElementById('import-dane');
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

if(reportsCard && !window.__wfReportsCardHandlerBound){
  window.__wfReportsCardHandlerBound = true;
  reportsCard.addEventListener('click', () => {
    openReportsView();
  });
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

function formatNumber(value){
  const number = Number(value);
  if(!Number.isFinite(number)) return '';
  return new Intl.NumberFormat('pl-PL', { maximumFractionDigits: 2 }).format(number);
}

function formatPercent(value, options = {}){
  const number = Number(value);
  if(!Number.isFinite(number)) return '0%';
  const {
    minimumFractionDigits = number > 0 && number < 10 ? 1 : 0,
    maximumFractionDigits = number > 0 && number < 10 ? 1 : 0
  } = options;
  return `${new Intl.NumberFormat('pl-PL', {
    minimumFractionDigits,
    maximumFractionDigits
  }).format(number)}%`;
}

function findHeaderKey(headers, variants){
  const map = new Map();
  const normalizedHeaders = headers.map(header => ({
    original: header,
    normalized: normalizeHeaderKey(header)
  }));
  headers.forEach(header => {
    map.set(normalizeHeaderKey(header), header);
  });
  for(const variant of variants){
    const found = map.get(normalizeHeaderKey(variant));
    if(found) return found;
  }
  for(const variant of variants){
    const normalizedVariant = normalizeHeaderKey(variant);
    const found = normalizedHeaders.find(header => header.normalized.includes(normalizedVariant));
    if(found) return found.original;
  }
  return '';
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

function dedupeReportRows(rows){
  const bestByIndex = new Map();
  rows.forEach(row => {
    const key = normalizeIndexValue(row.index);
    if(!key) return;
    const current = bestByIndex.get(key);
    if(!current || row.ranking < current.ranking){
      bestByIndex.set(key, row);
    }
  });
  return Array.from(bestByIndex.values());
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

function normalizeDetailedSalesRow(row, headers){
  const indexKey = findHeaderKey(headers, ['indeks', 'index', 'id', 'numer katalogowy', 'sku number']);
  const nameKey = findHeaderKey(headers, ['nazwa', 'name', 'nazwa towaru', 'nazwa produktu']);
  const groupNameKey = findHeaderKey(headers, ['nazwa grupa', 'nazwa_grupa', 'grupa']);
  const repKey = findHeaderKey(headers, ['opiekun klienta', 'opiekun_klienta', 'przedstawiciel handlowy', 'przedstawiciel', 'handlowiec']);
  const customerCodeKey = findHeaderKey(headers, ['kod kh', 'kod_kh', 'numer klienta', 'kod klienta']);
  const customerNameKey = findHeaderKey(headers, ['nazwa kh', 'nazwa_kh', 'nazwa klienta', 'klient']);
  const valueKey = findHeaderKey(headers, ['sprzedaż_waluta', 'sprzedaz_waluta', 'sprzedaż wartościowa', 'sprzedaz wartosciowa', 'wartość', 'wartosc', 'value']);

  const valueRaw = row[valueKey];
  const valueNumber = Number(String(valueRaw ?? '').replace(/\s+/g, '').replace(',', '.'));

  return {
    representative: String(row[repKey] ?? '').trim() || 'Bez przedstawiciela',
    customerCode: String(row[customerCodeKey] ?? '').trim(),
    customerName: String(row[customerNameKey] ?? '').trim() || 'Bez nazwy klienta',
    index: String(row[indexKey] ?? '').trim(),
    name: String(row[nameKey] ?? '').trim(),
    groupName: String(row[groupNameKey] ?? '').trim(),
    value: Number.isFinite(valueNumber) ? valueNumber : 0,
    reportLabel: '',
    reportOrder: 0
  };
}

function normalizeWeeklySalesRowsFromMatrix(matrix){
  const rows = Array.isArray(matrix) ? matrix : [];
  const firstRow = rows[0] || [];
  const secondRow = rows[1] || [];
  const isWideWeeklyReport =
    normalizeHeaderKey(firstRow[0]).includes('tydzien sprzedazy') ||
    normalizeHeaderKey(firstRow[0]).includes('tydzień sprzedaży');

  if(!isWideWeeklyReport){
    const headers = rows.length ? rows[0] : [];
    return rows.slice(1)
      .map(values => {
        const row = {};
        headers.forEach((header, index) => {
          row[header] = values[index] ?? '';
        });
        return normalizeWeeklySalesRow(row, headers);
      })
      .filter(row => row.index);
  }

  const repIndex = secondRow.findIndex(header => normalizeHeaderKey(header) === 'opiekun klienta');
  const customerCodeIndex = secondRow.findIndex(header => normalizeHeaderKey(header) === 'kod kh');
  const customerNameIndex = secondRow.findIndex(header => normalizeHeaderKey(header) === 'nazwa kh');
  const productIndex = secondRow.findIndex(header => normalizeHeaderKey(header) === 'kh kod prod');
  const producerIndex = secondRow.findIndex(header => normalizeHeaderKey(header) === 'skrot producenta');
  const hasProductColumn = productIndex >= 0;
  const salesColumnIndexes = secondRow
    .map((header, index) => ({ header, index }))
    .filter(item => normalizeHeaderKey(item.header).includes('sprzedaz waluta'))
    .map(item => item.index);

  return rows.slice(2).flatMap(values => {
    const representative = String(values[repIndex] ?? '').trim() || 'Bez przedstawiciela';
    const customerCode = String(values[customerCodeIndex] ?? '').trim();
    const customerName = String(values[customerNameIndex] ?? '').trim() || 'Bez nazwy klienta';
    const index = hasProductColumn ? String(values[productIndex] ?? '').trim() : '';
    const producer = producerIndex >= 0 ? String(values[producerIndex] ?? '').trim() : '';
    if(hasProductColumn && !index) return [];
    if(!hasProductColumn && !customerCode && !customerName) return [];

    return salesColumnIndexes.map((columnIndex, weekIndex) => {
      const valueRaw = values[columnIndex];
      const value = Number(String(valueRaw ?? '').replace(/\s+/g, '').replace(',', '.'));
      const weekHeader = String(firstRow[columnIndex] ?? '').trim();
      const weekLabel = weekHeader && !weekHeader.startsWith('=') ? weekHeader : `Tydzień ${weekIndex + 1}`;
      return {
        representative,
        customerCode,
        customerName,
        index,
        name: '',
        producer,
        week: weekLabel,
        quantity: 0,
        value: Number.isFinite(value) ? value : 0
      };
    });
  });
}

function normalizeDetailedSalesRowsFromMatrix(matrix){
  const rows = Array.isArray(matrix) ? matrix : [];
  const headers = rows[0] || [];

  return rows.slice(1)
    .map(values => {
      const row = {};
      headers.forEach((header, index) => {
        row[header] = values[index] ?? '';
      });
      return normalizeDetailedSalesRow(row, headers);
    })
    .filter(row => row.index || row.customerCode || row.customerName);
}

function getWeeklySalesRepresentatives(){
  return Array.from(new Set(weeklySalesSourceRows.map(row => row.representative).filter(Boolean)))
    .sort((a, b) => a.localeCompare(b, 'pl'));
}

function getDetailedSalesCustomers(){
  const map = new Map();
  detailedSalesSourceRows.forEach(row => {
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

function getDetailedSalesGroups(){
  return Array.from(new Set(
    detailedSalesGeneratedRows.map(row => String(row.groupName || '').trim()).filter(Boolean)
  )).sort((a, b) => a.localeCompare(b, 'pl'));
}

function normalizeDetailedSalesSelectedGroups(values){
  const source = Array.isArray(values)
    ? values
    : [values];
  return Array.from(new Set(
    source.map(value => String(value || '').trim()).filter(Boolean)
  ));
}

function expandDetailedSalesSelectedGroups(values){
  const expandedGroups = new Set();
  normalizeDetailedSalesSelectedGroups(values).forEach(value => {
    const config = getDetailedSalesDashboardGroupConfig(value);
    if(config?.groups?.length){
      config.groups.forEach(groupName => {
        const normalizedGroupName = String(groupName || '').trim();
        if(normalizedGroupName){
          expandedGroups.add(normalizedGroupName);
        }
      });
      return;
    }
    expandedGroups.add(String(value || '').trim());
  });
  return Array.from(expandedGroups).filter(Boolean);
}

function getDetailedSalesSelectedGroups(){
  return normalizeDetailedSalesSelectedGroups(detailedSalesSelectedGroups);
}

function matchesDetailedSalesRowGroup(groupName){
  const selectedGroups = expandDetailedSalesSelectedGroups(detailedSalesSelectedGroups);
  if(!selectedGroups.length) return true;
  return selectedGroups.includes(String(groupName || '').trim());
}

function getDetailedSalesSelectedGroupsLabel(maxItems = 2){
  const selectedGroups = Array.from(new Set(
    getDetailedSalesSelectedGroups().map(groupName => getDetailedSalesDashboardGroupLabel(groupName))
  ));
  if(!selectedGroups.length) return 'Wszystkie grupy';
  if(selectedGroups.length <= maxItems){
    return selectedGroups.join(', ');
  }
  const visibleGroups = selectedGroups.slice(0, maxItems).join(', ');
  return `${visibleGroups} +${selectedGroups.length - maxItems}`;
}

function getWeeklySalesCustomers(){
  if(!weeklySalesSelectedRepresentative) return [];
  const map = new Map();
  weeklySalesSourceRows
    .filter(row => row.representative === weeklySalesSelectedRepresentative)
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

function getWeeklySalesProducers(){
  return Array.from(new Set(
    weeklySalesGeneratedRows.map(row => String(row.producer || '').trim()).filter(Boolean)
  )).sort((a, b) => a.localeCompare(b, 'pl'));
}

function compareWeeklySalesWeekLabels(a, b){
  return String(a || '').localeCompare(String(b || ''), 'pl', { numeric: true });
}

function getWeeklySalesWeeks(options = {}){
  const descending = options.descending !== false;
  const rows = weeklySalesGeneratedRows.length ? weeklySalesGeneratedRows : weeklySalesSourceRows;
  return Array.from(new Set(
    rows.map(row => String(row.week || '').trim()).filter(Boolean)
  )).sort((a, b) => descending
    ? compareWeeklySalesWeekLabels(b, a)
    : compareWeeklySalesWeekLabels(a, b)
  );
}

function getLatestWeeklySalesWeek(){
  const weeks = getWeeklySalesWeeks();
  return weeks[0] || '';
}

function getWeeklySalesComparisonWeeks(){
  const weeks = Array.from(new Set(
    weeklySalesSourceRows.map(row => String(row.week || '').trim()).filter(Boolean)
  )).sort((a, b) => compareWeeklySalesWeekLabels(a, b));
  return weeks.slice(Math.max(weeks.length - 2, 0));
}

function getWeeklySalesRepresentativeComparisonRows(){
  const [previousWeek, latestWeek] = getWeeklySalesComparisonWeeks();
  if(!previousWeek || !latestWeek) return [];

  const map = new Map();
  weeklySalesSourceRows
    .filter(row => String(row.week || '') === previousWeek || String(row.week || '') === latestWeek)
    .forEach(row => {
      const representative = row.representative || 'Bez przedstawiciela';
      const current = map.get(representative) || {
        representative,
        previousWeek,
        latestWeek,
        previousValue: 0,
        latestValue: 0,
        trendValue: 0
      };
      if(String(row.week || '') === previousWeek){
        current.previousValue += row.quantity || row.value;
      }else if(String(row.week || '') === latestWeek){
        current.latestValue += row.quantity || row.value;
      }
      map.set(representative, current);
    });

  return Array.from(map.values())
    .map(row => ({
      ...row,
      trendValue: row.latestValue - row.previousValue
    }))
    .sort((a, b) => b.latestValue - a.latestValue || a.representative.localeCompare(b.representative, 'pl'));
}

function sortWeeklySalesRows(rows){
  const source = Array.isArray(rows) ? [...rows] : [];
  const sortByDropDesc = weeklySalesSortOrder === 'sales-drop-desc';
  const sortBySalesAsc = weeklySalesSortOrder === 'sales-asc';

  source.sort((a, b) => {
    if(sortByDropDesc){
      const aTrend = Number.isFinite(a?.trendValue) ? Number(a.trendValue) : null;
      const bTrend = Number.isFinite(b?.trendValue) ? Number(b.trendValue) : null;
      const aIsDrop = Number.isFinite(aTrend) && aTrend < 0;
      const bIsDrop = Number.isFinite(bTrend) && bTrend < 0;
      if(aIsDrop !== bIsDrop){
        return aIsDrop ? -1 : 1;
      }
      if(aIsDrop && bIsDrop){
        const aDrop = Math.abs(aTrend);
        const bDrop = Math.abs(bTrend);
        if(aDrop !== bDrop){
          return bDrop - aDrop;
        }
      }
    }

    const aValue = Number(a?.quantity || a?.value || 0);
    const bValue = Number(b?.quantity || b?.value || 0);
    if(aValue !== bValue){
      return sortBySalesAsc ? aValue - bValue : bValue - aValue;
    }
    const weekCompare = String(a?.week || '').localeCompare(String(b?.week || ''), 'pl', { numeric: true });
    if(weekCompare) return weekCompare;
    const customerCompare = String(a?.customerName || '').localeCompare(String(b?.customerName || ''), 'pl');
    if(customerCompare) return customerCompare;
    return String(a?.index || '').localeCompare(String(b?.index || ''), 'pl', { numeric: true });
  });

  return source;
}

function sortWeeklyComparisonRows(rows){
  const source = Array.isArray(rows) ? [...rows] : [];
  const sortByDropDesc = weeklySalesSortOrder === 'sales-drop-desc';
  const sortBySalesAsc = weeklySalesSortOrder === 'sales-asc';

  source.sort((a, b) => {
    if(sortByDropDesc){
      const aTrend = Number(a?.trendValue || 0);
      const bTrend = Number(b?.trendValue || 0);
      const aIsDrop = aTrend < 0;
      const bIsDrop = bTrend < 0;
      if(aIsDrop !== bIsDrop){
        return aIsDrop ? -1 : 1;
      }
      if(aIsDrop && bIsDrop){
        const aDrop = Math.abs(aTrend);
        const bDrop = Math.abs(bTrend);
        if(aDrop !== bDrop){
          return bDrop - aDrop;
        }
      }
    }

    const aValue = Number(a?.latestValue || 0);
    const bValue = Number(b?.latestValue || 0);
    if(aValue !== bValue){
      return sortBySalesAsc ? aValue - bValue : bValue - aValue;
    }
    return String(a?.representative || '').localeCompare(String(b?.representative || ''), 'pl');
  });

  return source;
}

function getFilteredWeeklySalesRows(){
  let rows = weeklySalesGeneratedRows;
  const latestWeek = getLatestWeeklySalesWeek();
  const showAllWeeks = weeklySalesSelectedWeek === '__all__';
  const activeWeek = showAllWeeks ? '' : (weeklySalesSelectedWeek || latestWeek);

  if(weeklySalesOnlyLastWeek250){
    const totalsByCustomer = new Map();
    rows.filter(row => String(row.week || '') === latestWeek).forEach(row => {
      const key = `${row.representative}|||${row.customerCode}|||${row.customerName}`;
      totalsByCustomer.set(key, (totalsByCustomer.get(key) || 0) + (row.quantity || row.value));
    });
    rows = rows.filter(row => {
      const key = `${row.representative}|||${row.customerCode}|||${row.customerName}`;
      return String(row.week || '') === latestWeek && (totalsByCustomer.get(key) || 0) >= 250;
    });
  }else if(!showAllWeeks && activeWeek){
    rows = rows.filter(row => String(row.week || '') === activeWeek);
  }

  if(weeklySalesSelectedProducer){
    rows = rows.filter(row => String(row.producer || '').trim() === weeklySalesSelectedProducer);
  }

  return sortWeeklySalesRows(rows);
}

function getLatestWeeksFromRows(rows, count){
  const weeks = Array.from(new Set(rows.map(row => String(row.week || '').trim()).filter(Boolean)))
    .sort((a, b) => a.localeCompare(b, 'pl', { numeric: true }));
  return weeks.slice(Math.max(weeks.length - count, 0));
}

function sortDetailedSalesRows(rows){
  const source = Array.isArray(rows) ? [...rows] : [];
  const sortBySalesAsc = detailedSalesSortOrder === 'sales-asc';

  source.sort((a, b) => {
    const aValue = Number(a?.value || 0);
    const bValue = Number(b?.value || 0);
    if(aValue !== bValue){
      return sortBySalesAsc ? aValue - bValue : bValue - aValue;
    }
    const reportOrderCompare = Number(a?.reportOrder || 0) - Number(b?.reportOrder || 0);
    if(reportOrderCompare) return reportOrderCompare;
    const customerCompare = String(a?.customerName || '').localeCompare(String(b?.customerName || ''), 'pl');
    if(customerCompare) return customerCompare;
    return String(a?.index || '').localeCompare(String(b?.index || ''), 'pl', { numeric: true });
  });

  return source;
}

function generateWeeklySalesReport(){
  weeklySalesGeneratedRows = [];
  if(!weeklySalesSelectedRepresentative) return;

  const selectedCustomer = weeklySalesSelectedCustomer;
  const filteredRows = weeklySalesSourceRows.filter(row => {
    if(row.representative !== weeklySalesSelectedRepresentative) return false;
    if(!selectedCustomer) return true;
    const [customerCode, customerName] = selectedCustomer.split('|||');
    return row.customerCode === customerCode && row.customerName === customerName;
  });

  const grouped = new Map();
  filteredRows.forEach(row => {
    const key = [
      row.representative,
      row.customerCode,
      row.customerName,
      normalizeIndexValue(row.index),
      row.week
    ].join('|||');
    const current = grouped.get(key) || {
      representative: row.representative,
      customerCode: row.customerCode,
      customerName: row.customerName,
      index: row.index,
      name: row.name,
      producer: row.producer,
      week: row.week,
      quantity: 0,
      value: 0
    };
    current.quantity += row.quantity;
    current.value += row.value;
    grouped.set(key, current);
  });

  const groupedRows = Array.from(grouped.values());
  const weeks = Array.from(new Set(groupedRows.map(row => String(row.week || '').trim()).filter(Boolean)))
    .sort((a, b) => a.localeCompare(b, 'pl', { numeric: true }));
  const weekIndexByLabel = new Map(weeks.map((week, index) => [week, index]));
  const salesByComparableKey = new Map();
  groupedRows.forEach(row => {
    const comparableKey = [
      row.representative,
      row.customerCode,
      row.customerName,
      normalizeIndexValue(row.index),
      normalizeProducerValue(row.producer),
      row.name
    ].join('|||');
    salesByComparableKey.set(`${comparableKey}|||${row.week}`, row.quantity || row.value);
  });
  groupedRows.forEach(row => {
    const comparableKey = [
      row.representative,
      row.customerCode,
      row.customerName,
      normalizeIndexValue(row.index),
      normalizeProducerValue(row.producer),
      row.name
    ].join('|||');
    const currentWeekIndex = weekIndexByLabel.get(String(row.week || ''));
    const previousWeek = Number.isInteger(currentWeekIndex) ? weeks[currentWeekIndex - 1] : '';
    const previousValue = previousWeek ? salesByComparableKey.get(`${comparableKey}|||${previousWeek}`) : undefined;
    const currentValue = row.quantity || row.value;
    row.previousWeek = previousWeek || '';
    row.previousValue = Number.isFinite(previousValue) ? previousValue : null;
    row.trendValue = Number.isFinite(previousValue) ? currentValue - previousValue : null;
    const discountInfo = getCustomerDiscountInfo(row.customerCode, row.representative);
    row.discountLabel = discountInfo.label;
    row.discountDetails = discountInfo.details;
    row.discountNetwork = discountInfo.network;
    row.discountShortcut = discountInfo.shortcut;
  });

  weeklySalesGeneratedRows = groupedRows.sort((a, b) => {
    const weekCompare = String(b.week).localeCompare(String(a.week), 'pl', { numeric: true });
    if(weekCompare) return weekCompare;
    const customerCompare = String(a.customerName).localeCompare(String(b.customerName), 'pl');
    if(customerCompare) return customerCompare;
    return String(a.index).localeCompare(String(b.index), 'pl', { numeric: true });
  });
}

function generateDetailedSalesReport(){
  const selectedCustomer = detailedSalesSelectedCustomer;
  const grouped = new Map();

  detailedSalesSourceRows.forEach(row => {
    if(selectedCustomer){
      const [customerCode, customerName] = selectedCustomer.split('|||');
      if(row.customerCode !== customerCode || row.customerName !== customerName){
        return;
      }
    }

    const key = [
      row.reportOrder,
      row.reportLabel,
      row.customerCode,
      row.customerName,
      normalizeIndexValue(row.index),
      row.name,
      row.groupName
    ].join('|||');
    const current = grouped.get(key) || {
      representative: row.representative,
      reportLabel: row.reportLabel,
      reportOrder: row.reportOrder,
      customerCode: row.customerCode,
      customerName: row.customerName,
      index: row.index,
      name: row.name,
      groupName: row.groupName,
      value: 0
    };
    current.value += row.value;
    grouped.set(key, current);
  });

  detailedSalesGeneratedRows = sortDetailedSalesRows(Array.from(grouped.values()));
}

function getFilteredDetailedSalesRows(){
  let rows = detailedSalesGeneratedRows;
  rows = rows.filter(row => matchesDetailedSalesRowGroup(row.groupName));
  return sortDetailedSalesRows(rows);
}

function normalizeDetailedSalesName(value){
  return normalizeHeaderKey(value);
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

function getDetailedSalesCustomerKey(code, name){
  return `${String(code || '').trim()}|||${String(name || '').trim()}`;
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

function buildWeeklySalesRepresentativePdfSummary(){
  const representatives = getWeeklySalesRepresentatives();
  const representativeLabel = weeklySalesSelectedRepresentative || (representatives.length === 1 ? representatives[0] : '');
  let sourceRows = Array.isArray(weeklySalesSourceRows) ? [...weeklySalesSourceRows] : [];
  if(representativeLabel){
    sourceRows = sourceRows.filter(row => row.representative === representativeLabel);
  }
  if(!sourceRows.length){
    return {
      representativeLabel,
      activeWeekLabel: '',
      activeSalesTotal: 0,
      activeCustomersCount: 0,
      previousWeekLabel: '',
      previousSalesTotal: 0,
      trendValue: null,
      topCustomers: []
    };
  }

  const availableWeeks = Array.from(new Set(
    sourceRows.map(row => String(row.week || '').trim()).filter(Boolean)
  )).sort((a, b) => compareWeeklySalesWeekLabels(a, b));
  const latestWeek = availableWeeks[availableWeeks.length - 1] || '';
  const explicitWeek = weeklySalesSelectedWeek === '__all__'
    ? '__all__'
    : (weeklySalesSelectedWeek || latestWeek);
  let activeRows = sourceRows;

  if(weeklySalesOnlyLastWeek250 && latestWeek){
    const totalsByCustomer = new Map();
    sourceRows
      .filter(row => String(row.week || '') === latestWeek)
      .forEach(row => {
        const key = `${row.customerCode}|||${row.customerName}`;
        totalsByCustomer.set(key, (totalsByCustomer.get(key) || 0) + Number(row.quantity || row.value || 0));
      });
    activeRows = sourceRows.filter(row => {
      const key = `${row.customerCode}|||${row.customerName}`;
      return String(row.week || '') === latestWeek && (totalsByCustomer.get(key) || 0) >= 250;
    });
  }else if(explicitWeek !== '__all__' && explicitWeek){
    activeRows = sourceRows.filter(row => String(row.week || '') === explicitWeek);
  }

  const activeSalesTotal = activeRows.reduce((sum, row) => sum + Number(row.quantity || row.value || 0), 0);
  const totalsByCustomer = new Map();
  activeRows.forEach(row => {
    const key = `${row.customerCode}|||${row.customerName}`;
    const current = totalsByCustomer.get(key) || {
      customerCode: row.customerCode,
      customerName: row.customerName,
      salesValue: 0
    };
    current.salesValue += Number(row.quantity || row.value || 0);
    totalsByCustomer.set(key, current);
  });

  const activeWeekLabel = weeklySalesOnlyLastWeek250
    ? `Ostatni tydzien ${latestWeek || ''} | min. 250 GBP`
    : (explicitWeek === '__all__' ? 'Wszystkie tygodnie' : (explicitWeek || latestWeek || ''));
  const activeWeekIndex = explicitWeek && explicitWeek !== '__all__'
    ? availableWeeks.indexOf(explicitWeek)
    : availableWeeks.length - 1;
  const previousWeekLabel = activeWeekIndex > 0 ? availableWeeks[activeWeekIndex - 1] : '';
  const previousSalesTotal = previousWeekLabel
    ? sourceRows
      .filter(row => String(row.week || '') === previousWeekLabel)
      .reduce((sum, row) => sum + Number(row.quantity || row.value || 0), 0)
    : 0;

  return {
    representativeLabel,
    activeWeekLabel,
    activeSalesTotal,
    activeCustomersCount: totalsByCustomer.size,
    previousWeekLabel,
    previousSalesTotal,
    trendValue: previousWeekLabel ? activeSalesTotal - previousSalesTotal : null,
    topCustomers: Array.from(totalsByCustomer.values())
      .sort((a, b) => b.salesValue - a.salesValue || String(a.customerName || '').localeCompare(String(b.customerName || ''), 'pl'))
      .slice(0, 5)
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
    const selectedCustomerRows = detailedSalesSourceRows.filter(row => {
      return row.customerCode === customerCode && row.customerName === customerName;
    });
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

async function loadCustomerDiscountMap(){
  if(customerDiscountMapCache) return customerDiscountMapCache;

  try{
    const res = await fetch(jsonUrlCustomerDiscounts);
    const data = await res.json();
    const map = new Map();
    (Array.isArray(data) ? data : []).forEach(row => {
      const code = normalizeCustomerCodeValue(row.KH_KOD);
      if(!code) return;
      if(!map.has(code)) map.set(code, []);
      map.get(code).push({
        discount: formatDiscountPercent(row.RABAT_PROCENTOWY),
        group: String(row.NAZWA_GRUPA ?? '').trim(),
        representative: String(row.OPIEKUN_KLIENTA ?? '').trim(),
        network: String(row.SIEC_OBCA ?? '').trim(),
        shortcut: String(row.SKROT ?? '').trim()
      });
    });
    customerDiscountMapCache = map;
  }catch(e){
    console.error('Customer discounts load error', e);
    customerDiscountMapCache = new Map();
  }

  return customerDiscountMapCache;
}

function getCustomerDiscountInfo(customerCode, representative){
  const rows = customerDiscountMapCache?.get(normalizeCustomerCodeValue(customerCode)) || [];
  if(!rows.length){
    return { label: '', details: '', network: '', shortcut: '' };
  }

  const representativeKey = normalizeProducerValue(representative);
  const representativeRows = representativeKey
    ? rows.filter(row => normalizeProducerValue(row.representative) === representativeKey)
    : [];
  const sourceRows = representativeRows.length ? representativeRows : rows;
  const seenEntries = new Set();
  const entries = [];

  sourceRows.forEach(row => {
    if(!row.discount) return;
    const key = `${row.group}|||${row.discount}`;
    if(seenEntries.has(key)) return;
    seenEntries.add(key);
    entries.push({
      group: row.group,
      discount: row.discount
    });
  });

  const networks = Array.from(new Set(sourceRows.map(row => row.network).filter(value => value && value !== '(puste)')));
  const shortcuts = Array.from(new Set(sourceRows.map(row => row.shortcut).filter(Boolean)));
  const label = entries
    .map(entry => [entry.group, entry.discount].filter(Boolean).join(' '))
    .join(', ');

  return {
    label,
    details: entries.map(entry => [entry.group, entry.discount].filter(Boolean).join(': ')).join('; '),
    network: networks.join(', '),
    shortcut: shortcuts.join(', ')
  };
}

function renderWeeklyTrendHtml(trendValue, title = ''){
  if(!Number.isFinite(trendValue)){
    return '<span class="weekly-trend weekly-trend--empty">-</span>';
  }
  const trendClass = trendValue > 0 ? 'weekly-trend--up' : trendValue < 0 ? 'weekly-trend--down' : 'weekly-trend--flat';
  const trendArrow = trendValue > 0 ? '↑' : trendValue < 0 ? '↓' : '→';
  const titleAttr = title ? ` title="${escapeAttr(title)}"` : '';
  return `<span class="weekly-trend ${trendClass}"${titleAttr}>${trendArrow} ${formatNumber(Math.abs(trendValue))}</span>`;
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

function getClientReportRepresentatives(){
  return Array.from(new Set(clientReportSourceRows.map(row => row.representative).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'pl'));
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
  const customerRows = topSuggestionSourceRows.filter(row =>
    row.representative === topSuggestionSelectedRepresentative &&
    row.customerCode === customerCode &&
    row.customerName === customerName
  );
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
  const clientRows = clientReportSourceRows.filter(row =>
    row.representative === clientReportSelectedRepresentative &&
    row.customerCode === customerCode &&
    row.customerName === customerName
  );
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
      const enrichedRow = {
        INDEKS: row.index,
        NAZWA: row.name,
        SKROT_PRODUCENTA: row.producer,
        'KOD EAN': row.ean || '',
        Ranking: row.rankingLabel || (Number.isFinite(row.ranking) ? String(row.ranking) : ''),
        'Grupa produktowa': group.label
      };
      purchased.push(enrichedRow);
      purchasedCount += 1;
    });

    missingTopRows.forEach(row => {
      const enrichedRow = {
        INDEKS: row.index,
        NAZWA: row.name,
        SKROT_PRODUCENTA: row.producer,
        'KOD EAN': row.ean || '',
        Ranking: row.rankingLabel || (Number.isFinite(row.ranking) ? String(row.ranking) : ''),
        'Grupa produktowa': group.label
      };
      missing.push(enrichedRow);
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

function getFilteredDetailedSalesTopRows(rows, filters){
  return rows.filter(row => {
    if(filters.producer && !includesText(row.SKROT_PRODUCENTA, filters.producer)) return false;
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

  return `
    <div class="reports-table-filters">
      <select onchange="setDetailedSalesTopTableFilter('${type}', 'producer', this.value)">
        <option value="">Wszyscy producenci</option>
        ${producerOptions}
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

function getDetailedSalesTopImageUrl(index){
  const idx = String(index || '').trim();
  if(!idx) return '';
  const ext = getPreferredImageExt(idx, imageBaseUrlRumunia);
  return buildImageUrl(idx, ext, imageBaseUrlRumunia);
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

function getDetailedSalesTopGroupKey(value){
  return String(value || '').trim().toLowerCase();
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

function wrapDetailedSalesTopCompactContent(content){
  return `<div class="reports-compact-mode">${content}</div>`;
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
              ${renderDetailedSalesTopInfoToolbar(purchasedRows, {
                emptyText: 'Brak kupionych produktów dla wybranego producenta.',
                label: 'Widoczne'
              })}
              ${renderDetailedSalesTopRowsTable(purchasedRows, 'Brak kupionych produktów dla wybranego producenta.')}
            </div>
            <div class="reports-column">
              <h3 class="reports-table-title">Brakujące produkty</h3>
              ${renderDetailedSalesTopQuickSelectionToolbar(missingRows, {
                scope: 'producer',
                selectLabel: 'Zaznacz produkty producenta',
                clearLabel: 'Odznacz produkty producenta'
              })}
              ${renderDetailedSalesTopRowsTable(missingRows, 'Brak brakujących produktów dla wybranego producenta.', { selectable: true })}
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

function renderWeeklySalesContent(){
  const representatives = getWeeklySalesRepresentatives();
  const customers = getWeeklySalesCustomers();
  const producers = getWeeklySalesProducers();
  const weeks = getWeeklySalesWeeks();
  const latestWeek = getLatestWeeklySalesWeek();
  const isAdminReports = isCurrentUserAdmin();
  const filteredRows = getFilteredWeeklySalesRows();
  const comparisonRows = sortWeeklyComparisonRows(getWeeklySalesRepresentativeComparisonRows());
  const comparisonWeeks = getWeeklySalesComparisonWeeks();
  const activeRowsCount = weeklySalesRepComparison ? comparisonRows.length : filteredRows.length;
  const activeSalesTotal = weeklySalesRepComparison
    ? comparisonRows.reduce((sum, row) => sum + row.latestValue, 0)
    : filteredRows.reduce((sum, row) => sum + (row.quantity || row.value), 0);
  const previousSalesTotal = comparisonRows.reduce((sum, row) => sum + row.previousValue, 0);
  const comparisonTrendValue = activeSalesTotal - previousSalesTotal;
  const customerLabel = weeklySalesSelectedCustomer ? 'wybrany klient' : 'wszyscy klienci';
  const selectedWeekLabel = weeklySalesSelectedWeek === '__all__'
    ? 'wszystkie tygodnie'
    : (weeklySalesSelectedWeek || latestWeek || 'brak tygodnia');
  const weekLabel = weeklySalesOnlyLastWeek250
    ? `ostatni tydzień ${latestWeek || ''}, min. 250 GBP`
    : selectedWeekLabel;
  const scopeLabel = weeklySalesRepComparison
    ? `wszyscy PH | ${comparisonWeeks.join(' do ') || 'brak tygodni'}`
    : `${weeklySalesSelectedRepresentative || 'Wybierz przedstawiciela'} | ${customerLabel} | ${weekLabel}`;
  const rowsHtml = filteredRows.map((row, index) => {
    const trendValue = row.trendValue;
    const trendHtml = renderWeeklyTrendHtml(trendValue, `Poprzedni tydzień: ${row.previousWeek || ''}`);
    const discountTitle = [
      row.discountDetails ? `Grupy: ${row.discountDetails}` : '',
      row.discountNetwork ? `Sieć: ${row.discountNetwork}` : '',
      row.discountShortcut ? `Skrót: ${row.discountShortcut}` : ''
    ].filter(Boolean).join(' | ');
    const discountHtml = row.discountLabel
      ? `<span title="${escapeAttr(discountTitle)}">${escapeHtml(row.discountLabel)}</span>`
      : '';
    const adminOnlyCellsHtml = isAdminReports
      ? `
        <td>${escapeHtml(row.index)}</td>
        <td>${escapeHtml(row.name)}</td>
        <td>${escapeHtml(row.producer)}</td>
      `
      : '';
    return `
      <tr>
        <td>${index + 1}</td>
        <td>${escapeHtml(row.week)}</td>
        <td>${escapeHtml(row.customerCode)}</td>
        <td>${escapeHtml(row.customerName)}</td>
        <td>${discountHtml}</td>
        ${adminOnlyCellsHtml}
        <td>${formatNumber(row.quantity || row.value)}</td>
        <td>${trendHtml}</td>
      </tr>
    `;
  }).join('');
  const comparisonRowsHtml = comparisonRows.map((row, index) => {
    const trendTitle = `${row.previousWeek}: ${formatNumber(row.previousValue)} | ${row.latestWeek}: ${formatNumber(row.latestValue)}`;
    return `
      <tr>
        <td>${index + 1}</td>
        <td>${escapeHtml(row.representative)}</td>
        <td>${escapeHtml(row.previousWeek)}</td>
        <td>${formatNumber(row.previousValue)}</td>
        <td>${escapeHtml(row.latestWeek)}</td>
        <td>${formatNumber(row.latestValue)}</td>
        <td>${renderWeeklyTrendHtml(row.trendValue, trendTitle)}</td>
      </tr>
    `;
  }).join('');
  const tableHtml = weeklySalesRepComparison
    ? `
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Lp.</th>
              <th>Przedstawiciel handlowy</th>
              <th>Poprzedni tydzień</th>
              <th>Sprzedaż poprzedni tydzień</th>
              <th>Aktualny tydzień</th>
              <th>Sprzedaż aktualny tydzień</th>
              <th>Trend</th>
            </tr>
          </thead>
          <tbody>
            ${comparisonRowsHtml || `<tr><td colspan="7" class="reports-empty">Zaimportuj plik z minimum dwoma tygodniami sprzedaży.</td></tr>`}
          </tbody>
        </table>
      </div>
    `
    : `
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Lp.</th>
              <th>Tydzień</th>
              <th>Kod klienta</th>
              <th>Nazwa klienta</th>
              <th>Rabat</th>
              ${isAdminReports ? `
                <th>INDEKS</th>
                <th>NAZWA</th>
                <th>Producent</th>
              ` : ''}
              <th>Sprzedaż</th>
              <th>Trend</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml || `<tr><td colspan="${isAdminReports ? '10' : '7'}" class="reports-empty">${isAdminReports ? 'Zaimportuj plik Excel i wybierz przedstawiciela.' : 'Brak danych tygodniowych dla Twojego konta.'}</td></tr>`}
          </tbody>
        </table>
      </div>
    `;
  const toolbarHtml = isAdminReports
    ? `
    <div class="reports-toolbar">
      <div>
        <div class="import-title">Sprzedaż per indeks per tydzień</div>
        <div class="import-sub">Importuj Excel, wybierz przedstawiciela i sprawdź sprzedaż dla wszystkich lub wybranego klienta.</div>
      </div>
      <div class="reports-actions">
        <input id="weekly-sales-excel-input" class="import-input reports-file-input" type="file" accept=".xlsx,.xls" onchange="importWeeklySalesExcel()">
        <button class="btn-outline" onclick="pickReportFile('weekly-sales-excel-input')">Importuj Excel</button>
        <button class="btn-outline" onclick="exportWeeklySalesXlsx()" ${(weeklySalesGeneratedRows.length || weeklySalesSourceRows.length) ? '' : 'disabled'}>Zapisz XLSX</button>
      </div>
      <div class="import-info">
        ${weeklySalesImportFile ? 'Status: plik zaimportowany' : 'Status: oczekuje na import'}
      </div>
    </div>
  `
    : '';

  return `
    ${toolbarHtml}

    ${weeklySalesRepComparison ? '' : `<div class="reports-filters">
      ${isAdminReports ? `<label class="reports-filter-card">
        <span class="reports-limit-title">Przedstawiciel handlowy</span>
        <select onchange="setWeeklySalesRepresentative(this.value)">
          <option value="">Wybierz</option>
          ${representatives.map(rep => `<option value="${escapeAttr(rep)}" ${weeklySalesSelectedRepresentative === rep ? 'selected' : ''}>${escapeHtml(rep)}</option>`).join('')}
        </select>
      </label>` : ''}
      <label class="reports-filter-card">
        <span class="reports-limit-title">Klient</span>
        <select onchange="setWeeklySalesCustomer(this.value)" ${weeklySalesSelectedRepresentative ? '' : 'disabled'}>
          <option value="">Wszyscy klienci</option>
          ${customers.map(customer => {
            const value = `${customer.code}|||${customer.name}`;
            const label = [customer.code, customer.name].filter(Boolean).join(' - ') || 'Bez nazwy klienta';
            return `<option value="${escapeAttr(value)}" ${weeklySalesSelectedCustomer === value ? 'selected' : ''}>${escapeHtml(label)}</option>`;
          }).join('')}
        </select>
      </label>
      ${isAdminReports ? `<label class="reports-filter-card">
        <span class="reports-limit-title">Producent</span>
        <select onchange="setWeeklySalesProducer(this.value)" ${weeklySalesGeneratedRows.length ? '' : 'disabled'}>
          <option value="">Wszyscy producenci</option>
          ${producers.map(producer => `<option value="${escapeAttr(producer)}" ${weeklySalesSelectedProducer === producer ? 'selected' : ''}>${escapeHtml(producer)}</option>`).join('')}
        </select>
      </label>` : ''}
      <label class="reports-filter-card">
        <span class="reports-limit-title">Tydzień sprzedaży</span>
        <select onchange="setWeeklySalesWeek(this.value)" ${weeklySalesGeneratedRows.length || weeklySalesOnlyLastWeek250 ? '' : 'disabled'}>
          <option value="" ${weeklySalesSelectedWeek === '' ? 'selected' : ''}>Ostatni tydzień</option>
          <option value="__all__" ${weeklySalesSelectedWeek === '__all__' ? 'selected' : ''}>Wszystkie tygodnie</option>
          ${weeks.map(week => `<option value="${escapeAttr(week)}" ${weeklySalesSelectedWeek === week ? 'selected' : ''}>${escapeHtml(week)}</option>`).join('')}
        </select>
      </label>
      <label class="reports-filter-card">
        <span class="reports-limit-title">Sortowanie sprzedaży</span>
        <select onchange="setWeeklySalesSortOrder(this.value)">
          <option value="sales-desc" ${weeklySalesSortOrder === 'sales-desc' ? 'selected' : ''}>Sprzedaż: od największej do najmniejszej</option>
          <option value="sales-asc" ${weeklySalesSortOrder === 'sales-asc' ? 'selected' : ''}>Sprzedaż: od najmniejszej do największej</option>
          <option value="sales-drop-desc" ${weeklySalesSortOrder === 'sales-drop-desc' ? 'selected' : ''}>Spadki sprzedaży: od największych do najmniejszych</option>
        </select>
      </label>
    </div>`}

    ${isAdminReports ? `<div class="reports-inline-actions">
      <button class="btn-outline ${weeklySalesOnlyLastWeek250 ? 'reports-mode-active' : ''}" onclick="setWeeklySalesLastWeek250(!weeklySalesOnlyLastWeek250)" ${weeklySalesGeneratedRows.length ? '' : 'disabled'}>Klienci >= 250 GBP w ostatnim tygodniu</button>
      <button class="btn-outline ${weeklySalesRepComparison ? 'reports-mode-active' : ''}" onclick="setWeeklySalesRepComparison(!weeklySalesRepComparison)" ${weeklySalesSourceRows.length ? '' : 'disabled'}>Porównanie PH tydzień do tygodnia</button>
    </div>` : ''}

    <div class="reports-summary">
      <div class="reports-summary-card">
        <div class="reports-summary-title">Zakres</div>
        <div class="reports-summary-meta">${escapeHtml(scopeLabel)}</div>
      </div>
      <div class="reports-summary-card">
        <div class="reports-summary-title">${weeklySalesRepComparison ? 'Przedstawiciele' : 'Pozycje'}</div>
        <div class="reports-summary-meta">${activeRowsCount}</div>
      </div>
      <div class="reports-summary-card">
        <div class="reports-summary-title">Suma sprzedaży</div>
        <div class="reports-summary-meta">${formatNumber(activeSalesTotal)}</div>
      </div>
      ${weeklySalesRepComparison ? `
        <div class="reports-summary-card">
          <div class="reports-summary-title">Porównanie tygodni</div>
          <div class="reports-summary-meta">
            ${escapeHtml(comparisonWeeks[0] || '')}: ${formatNumber(previousSalesTotal)} |
            ${escapeHtml(comparisonWeeks[1] || '')}: ${formatNumber(activeSalesTotal)} |
            ${renderWeeklyTrendHtml(comparisonTrendValue)}
          </div>
        </div>
      ` : ''}
    </div>

    ${tableHtml}
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

function renderDetailedSalesGroupsFilter(groups, options = {}){
  const disabled = Boolean(options.disabled);
  const selectedGroups = getDetailedSalesSelectedGroups();
  const isOpen = !disabled && detailedSalesGroupsDropdownOpen;
  const disabledAttr = disabled ? 'disabled' : '';
  const selectedCountLabel = selectedGroups.length
    ? `${selectedGroups.length} wybrane`
    : 'Wszystkie';

  return `
    <div class="reports-filter-card reports-filter-card--multi ${isOpen ? 'is-open' : ''}">
      <span class="reports-limit-title">Grupa</span>
      <button
        type="button"
        class="reports-multi-select-trigger"
        onclick="toggleDetailedSalesGroupsDropdown()"
        aria-expanded="${isOpen ? 'true' : 'false'}"
        ${disabledAttr}
      >
        <span class="reports-multi-select-trigger-main">
          <span class="reports-multi-select-summary">${escapeHtml(getDetailedSalesSelectedGroupsLabel())}</span>
          <span class="reports-multi-select-meta">${escapeHtml(selectedCountLabel)}</span>
        </span>
        <span class="reports-multi-select-chevron" aria-hidden="true">${isOpen ? '▴' : '▾'}</span>
      </button>
      ${isOpen ? `
        <div class="reports-multi-select-popover">
          <div class="reports-multi-select-popover-head">
            <div class="reports-multi-select-popover-title">Wybierz grupy</div>
            <button type="button" class="reports-multi-select-action" onclick="clearDetailedSalesGroups()">Wszystkie</button>
          </div>
          <div class="reports-multi-select-list">
            ${groups.length
              ? groups.map(group => {
                const serializedGroup = escapeAttr(JSON.stringify(group));
                return `
                  <label class="reports-multi-select-option">
                    <input
                      type="checkbox"
                      value="${escapeAttr(group)}"
                      onchange="toggleDetailedSalesGroupSelection(${serializedGroup}, this.checked)"
                      ${selectedGroups.includes(group) ? 'checked' : ''}
                    >
                    <span>${escapeHtml(group)}</span>
                  </label>
                `;
              }).join('')
              : `<div class="reports-multi-select-empty">Brak grup do wyboru.</div>`
            }
          </div>
          <div class="reports-multi-select-popover-foot">
            <button type="button" class="reports-multi-select-done" onclick="closeDetailedSalesGroupsDropdown()">Gotowe</button>
          </div>
        </div>
      ` : ''}
    </div>
  `;
}

function renderDetailedSalesContent(){
  const hasData = detailedSalesSourceRows.length > 0;

  if(detailedSalesLoading && !hasData){
    return `
      <div class="reports-toolbar">
        <div>
          <div class="import-title">Sprzedaż szczegółowa</div>
          <div class="import-sub">Pobieram raport per indeks przypisany do Twojego konta z Firebase.</div>
        </div>
        <div class="import-info">Status: ładowanie raportu</div>
      </div>
    `;
  }

  const customers = getDetailedSalesCustomers();
  const groups = getDetailedSalesGroups();
  const filteredRows = getFilteredDetailedSalesRows();
  const activeCustomers = new Set(
    filteredRows.map(row => `${row.customerCode}|||${row.customerName}`)
  ).size;
  const totalValue = filteredRows.reduce((sum, row) => sum + Number(row.value || 0), 0);
  const customerLabel = detailedSalesSelectedCustomer ? 'wybrany klient' : 'wszyscy klienci';
  const weeksScopeLabel = formatWeeksCountLabel(detailedSalesLoadedReportsCount || detailedSalesWeeksLimit);
  const fileInfo = getDetailedSalesFileInfoHtml();
  const rowsHtml = filteredRows.map((row, index) => `
      <tr>
        <td>${index + 1}</td>
        <td>${escapeHtml(row.reportLabel || '-')}</td>
        <td>${escapeHtml(row.customerCode)}</td>
        <td>${escapeHtml(row.customerName)}</td>
        <td>${escapeHtml(row.index)}</td>
        <td>${escapeHtml(row.name)}</td>
        <td>${escapeHtml(row.groupName || '-')}</td>
        <td>${formatNumber(row.value)}</td>
      </tr>
    `).join('');
  const emptyMessage = detailedSalesErrorMessage
    || (hasData ? 'Brak danych szczegółowych dla wybranych filtrów.' : 'Brak szczegółowego raportu sprzedaży dla Twojego konta.');
  const topComparisonHtml = renderDetailedSalesTopComparisonContent();

  return `
    <div class="reports-toolbar">
      <div>
        <div class="import-title">Sprzedaż szczegółowa</div>
        <div class="import-sub">Raport per indeks przypisany do Twojego konta w Firebase.</div>
      </div>
      <div class="import-info">${fileInfo}</div>
    </div>

    <div class="reports-filters">
      <label class="reports-filter-card">
        <span class="reports-limit-title">Klient</span>
        <select onchange="setDetailedSalesCustomer(this.value)" ${hasData ? '' : 'disabled'}>
          <option value="">Wszyscy klienci</option>
          ${customers.map(customer => {
            const value = `${customer.code}|||${customer.name}`;
            const label = [customer.code, customer.name].filter(Boolean).join(' - ') || 'Bez nazwy klienta';
            return `<option value="${escapeAttr(value)}" ${detailedSalesSelectedCustomer === value ? 'selected' : ''}>${escapeHtml(label)}</option>`;
          }).join('')}
        </select>
      </label>
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
      ${renderDetailedSalesGroupsFilter(groups, { disabled: !detailedSalesGeneratedRows.length })}
      <label class="reports-filter-card">
        <span class="reports-limit-title">Sortowanie sprzedaży</span>
        <select onchange="setDetailedSalesSortOrder(this.value)" ${detailedSalesGeneratedRows.length ? '' : 'disabled'}>
          <option value="sales-desc" ${detailedSalesSortOrder === 'sales-desc' ? 'selected' : ''}>Sprzedaż: od największej do najmniejszej</option>
          <option value="sales-asc" ${detailedSalesSortOrder === 'sales-asc' ? 'selected' : ''}>Sprzedaż: od najmniejszej do największej</option>
        </select>
      </label>
    </div>

    <div class="reports-summary">
      <div class="reports-summary-card">
        <div class="reports-summary-title">Zakres</div>
        <div class="reports-summary-meta">${escapeHtml(`${customerLabel} | ${weeksScopeLabel}`)}</div>
      </div>
      <div class="reports-summary-card">
        <div class="reports-summary-title">Pozycje</div>
        <div class="reports-summary-meta">${filteredRows.length}</div>
      </div>
      <div class="reports-summary-card">
        <div class="reports-summary-title">Klienci</div>
        <div class="reports-summary-meta">${activeCustomers}</div>
      </div>
      <div class="reports-summary-card">
        <div class="reports-summary-title">Suma sprzedaży</div>
        <div class="reports-summary-meta">${formatNumber(totalValue)}</div>
      </div>
    </div>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Lp.</th>
            <th>Tydzień</th>
            <th>Kod klienta</th>
            <th>Nazwa klienta</th>
            <th>INDEKS</th>
            <th>NAZWA</th>
            <th>Grupa</th>
            <th>Sprzedaż</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml || `<tr><td colspan="8" class="reports-empty">${escapeHtml(emptyMessage)}</td></tr>`}
        </tbody>
      </table>
    </div>

    ${topComparisonHtml}
  `;
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
        <select onchange="setDetailedSalesComparisonCustomer(this.value)" ${comparisonCustomers.length ? '' : 'disabled'}>
          <option value="">${comparisonCustomers.length ? 'Wybierz klienta do porównania' : 'Brak klientów w raporcie'}</option>
          ${comparisonCustomers.map(customer => {
            const value = `${customer.code}|||${customer.name}`;
            const label = [customer.code, customer.name].filter(Boolean).join(' - ') || 'Bez nazwy klienta';
            return `<option value="${escapeAttr(value)}" ${detailedSalesComparisonCustomer === value ? 'selected' : ''}>${escapeHtml(label)}</option>`;
          }).join('')}
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
        <select onchange="setDetailedSalesComparisonCustomer(this.value)" ${comparisonCustomers.length ? '' : 'disabled'}>
          <option value="">${comparisonCustomers.length ? 'Wybierz klienta do porównania' : 'Brak klientów w raporcie'}</option>
          ${comparisonCustomers.map(customer => {
            const value = `${customer.code}|||${customer.name}`;
            const label = [customer.code, customer.name].filter(Boolean).join(' - ') || 'Bez nazwy klienta';
            return `<option value="${escapeAttr(value)}" ${detailedSalesComparisonCustomer === value ? 'selected' : ''}>${escapeHtml(label)}</option>`;
          }).join('')}
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

function renderReportsBanner(){
  if(!reportsBannerActions) return;

  if(isCurrentUserAdmin()){
    reportsBannerActions.innerHTML = `
      <div class="flag-pill">
        <span class="flag-text">Raporty Sprzedaży</span>
      </div>
    `;
    return;
  }

  const areButtonsDisabled = phReportsAutoLoading || detailedSalesLoading;

  reportsBannerActions.innerHTML = `
    <button
      type="button"
      class="flag-pill reports-banner-pill ${reportsMode === 'weekly-sales' ? 'is-active' : ''}"
      onclick="setReportsMode('weekly-sales')"
      ${areButtonsDisabled ? 'disabled' : ''}
    >
      <span class="flag-text">Sprzedaż tygodniowa</span>
    </button>
    <button
      type="button"
      class="flag-pill reports-banner-pill ${reportsMode === 'detailed-sales' ? 'is-active' : ''}"
      onclick="setReportsMode('detailed-sales')"
      ${areButtonsDisabled ? 'disabled' : ''}
    >
      <span class="flag-text">Sprzedaż szczegółowa</span>
    </button>
    <button
      type="button"
      class="flag-pill reports-banner-pill ${reportsMode === 'top-rumunia-comparison' ? 'is-active' : ''}"
      onclick="setReportsMode('top-rumunia-comparison')"
      ${areButtonsDisabled ? 'disabled' : ''}
    >
      <span class="flag-text">Potencjał klienta</span>
    </button>
  `;
}

function isReportsModeAllowed(mode, isAdminReports = isCurrentUserAdmin()){
  const allowedModes = isAdminReports
    ? ['top-sales', 'client-gap', 'weekly-sales', 'top-suggestions']
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

function renderReportsView(){
  if(!reportsContainer) return;
  const isAdminReports = isCurrentUserAdmin();
  const contentByMode = {
    'top-sales': renderTopSalesReportContent,
    'client-gap': renderClientComparisonContent,
    'weekly-sales': renderWeeklySalesContent,
    'detailed-sales': renderDetailedSalesContent,
    'top-rumunia-comparison': renderDetailedSalesTopComparisonStandaloneContent,
    'top-suggestions': renderTopSuggestionsContent
  };
  const fallbackMode = isAdminReports ? 'top-sales' : 'weekly-sales';
  if(!isReportsModeAllowed(reportsMode, isAdminReports)){
    reportsMode = fallbackMode;
  }
  const renderContent = contentByMode[reportsMode] || (isAdminReports ? renderTopSalesReportContent : renderWeeklySalesContent);

  renderReportsBanner();

  reportsContainer.innerHTML = `
    <div class="reports-panel">
      ${isAdminReports ? `
      <div class="reports-mode-switch">
        <button class="btn-outline ${reportsMode === 'top-sales' ? 'reports-mode-active' : ''}" onclick="setReportsMode('top-sales')">Top sprzedaż</button>
        <button class="btn-outline ${reportsMode === 'client-gap' ? 'reports-mode-active' : ''}" onclick="setReportsMode('client-gap')">Potencjał klienta Top Rumunia</button>
        <button class="btn-outline ${reportsMode === 'weekly-sales' ? 'reports-mode-active' : ''}" onclick="setReportsMode('weekly-sales')">Sprzedaż tygodniowa</button>
        <button class="btn-outline ${reportsMode === 'top-suggestions' ? 'reports-mode-active' : ''}" onclick="setReportsMode('top-suggestions')">Propozycje Top</button>
      </div>` : ''}
      ${renderContent()}
    </div>
  `;
  registerLazyImages(reportsContainer);
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

async function setDetailedSalesCustomer(value){
  detailedSalesSelectedCustomer = value;
  generateDetailedSalesReport();
  const availableGroups = new Set(getDetailedSalesGroups());
  detailedSalesSelectedGroups = getDetailedSalesSelectedGroups().filter(group => availableGroups.has(group));
  clearDetailedSalesTopComparisonResultState();
  renderReportsView();
  await generateDetailedSalesTopComparison();
}

async function setDetailedSalesGroups(nextGroups){
  const availableGroups = new Set(getDetailedSalesGroups());
  detailedSalesSelectedGroups = normalizeDetailedSalesSelectedGroups(nextGroups)
    .filter(group => availableGroups.has(group));
  clearDetailedSalesTopComparisonResultState();
  renderReportsView();
  await generateDetailedSalesTopComparison();
}

function toggleDetailedSalesGroupsDropdown(){
  detailedSalesGroupsDropdownOpen = !detailedSalesGroupsDropdownOpen;
  renderReportsView();
}

function closeDetailedSalesGroupsDropdown(){
  if(!detailedSalesGroupsDropdownOpen) return;
  detailedSalesGroupsDropdownOpen = false;
  renderReportsView();
}

function handleDetailedSalesGroupsDropdownOutsideClick(event){
  if(!detailedSalesGroupsDropdownOpen) return;
  const target = event?.target;
  if(target?.closest?.('.reports-filter-card--multi')){
    return;
  }
  detailedSalesGroupsDropdownOpen = false;
  renderReportsView();
}

async function toggleDetailedSalesGroupSelection(value, checked){
  const nextGroups = new Set(getDetailedSalesSelectedGroups());
  if(checked){
    nextGroups.add(String(value || '').trim());
  }else{
    nextGroups.delete(String(value || '').trim());
  }
  await setDetailedSalesGroups(Array.from(nextGroups));
}

async function clearDetailedSalesGroups(){
  await setDetailedSalesGroups([]);
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

async function setDetailedSalesComparisonCustomer(value){
  detailedSalesComparisonCustomer = value;
  clearDetailedSalesTopComparisonResultState({ preserveCustomerSummaryRows: true });
  renderReportsView();
  await generateDetailedSalesTopComparison();
}

function setDetailedSalesSortOrder(value){
  detailedSalesSortOrder = ['sales-desc', 'sales-asc'].includes(value)
    ? value
    : 'sales-desc';
  renderReportsView();
}

async function setDetailedSalesWeeksLimit(value){
  detailedSalesWeeksLimit = normalizeDetailedSalesWeeksLimit(value);
  renderReportsView();
  if(!isCurrentUserAdmin() && (reportsMode === 'detailed-sales' || reportsMode === 'top-rumunia-comparison')){
    await ensurePhDetailedSalesDataLoaded();
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
    await ensurePhReportModeDataLoaded(reportsMode);
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

async function importWeeklySalesExcel(){
  const input = document.getElementById('weekly-sales-excel-input');
  const file = input?.files?.[0];
  if(!file) return;

  try{
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const sheetName = wb.SheetNames.includes('Export') ? 'Export' : wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    weeklySalesSourceRows = normalizeWeeklySalesRowsFromMatrix(matrix);
    weeklySalesImportFile = file.name;
    weeklySalesSelectedRepresentative = '';
    weeklySalesSelectedCustomer = '';
    weeklySalesSelectedProducer = '';
    weeklySalesSelectedWeek = '';
    weeklySalesOnlyLastWeek250 = false;
    weeklySalesRepComparison = false;
    weeklySalesSortOrder = 'sales-desc';
    weeklySalesGeneratedRows = [];
    await loadCustomerDiscountMap();

    const representatives = getWeeklySalesRepresentatives();
    if(representatives.length === 1){
      weeklySalesSelectedRepresentative = representatives[0];
      generateWeeklySalesReport();
    }
  }catch(e){
    console.error('Weekly sales import error', e);
    weeklySalesSourceRows = [];
    weeklySalesGeneratedRows = [];
    weeklySalesImportFile = '';
    weeklySalesSelectedRepresentative = '';
    weeklySalesSelectedCustomer = '';
    weeklySalesSelectedProducer = '';
    weeklySalesSelectedWeek = '';
    weeklySalesOnlyLastWeek250 = false;
    weeklySalesRepComparison = false;
    weeklySalesSortOrder = 'sales-desc';
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

async function setWeeklySalesRepresentative(value){
  weeklySalesSelectedRepresentative = value;
  weeklySalesSelectedCustomer = '';
  weeklySalesSelectedProducer = '';
  weeklySalesSelectedWeek = '';
  weeklySalesOnlyLastWeek250 = false;
  weeklySalesRepComparison = false;
  await loadCustomerDiscountMap();
  generateWeeklySalesReport();
  renderReportsView();
}

async function setWeeklySalesCustomer(value){
  weeklySalesSelectedCustomer = value;
  weeklySalesSelectedProducer = '';
  weeklySalesSelectedWeek = '';
  weeklySalesOnlyLastWeek250 = false;
  weeklySalesRepComparison = false;
  await loadCustomerDiscountMap();
  generateWeeklySalesReport();
  renderReportsView();
}

function setWeeklySalesProducer(value){
  weeklySalesSelectedProducer = value;
  weeklySalesRepComparison = false;
  renderReportsView();
}

function setWeeklySalesWeek(value){
  weeklySalesSelectedWeek = value;
  weeklySalesOnlyLastWeek250 = false;
  weeklySalesRepComparison = false;
  renderReportsView();
}

function setWeeklySalesSortOrder(value){
  weeklySalesSortOrder = ['sales-desc', 'sales-asc', 'sales-drop-desc'].includes(value)
    ? value
    : 'sales-desc';
  renderReportsView();
}

function setWeeklySalesLastWeek250(enabled){
  weeklySalesOnlyLastWeek250 = !!enabled;
  weeklySalesRepComparison = false;
  if(weeklySalesOnlyLastWeek250){
    weeklySalesSelectedWeek = getLatestWeeklySalesWeek();
  }
  renderReportsView();
}

function setWeeklySalesRepComparison(enabled){
  if(!isCurrentUserAdmin()){
    weeklySalesRepComparison = false;
    renderReportsView();
    return;
  }
  weeklySalesRepComparison = !!enabled;
  if(weeklySalesRepComparison){
    weeklySalesOnlyLastWeek250 = false;
  }
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

function exportWeeklySalesXlsx(){
  if(weeklySalesRepComparison){
    const comparisonRows = getWeeklySalesRepresentativeComparisonRows();
    if(!comparisonRows.length) return;
    const rows = [
      ['Przedstawiciel handlowy', 'Poprzedni tydzień', 'Sprzedaż poprzedni tydzień', 'Aktualny tydzień', 'Sprzedaż aktualny tydzień', 'Trend'],
      ...comparisonRows.map(row => [
        row.representative,
        row.previousWeek,
        row.previousValue,
        row.latestWeek,
        row.latestValue,
        row.trendValue
      ])
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'porownanie ph');
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    downloadBlob(blob, 'porownanie_ph_tydzien_do_tygodnia.xlsx');
    return;
  }

  const exportRows = getFilteredWeeklySalesRows();
  if(!exportRows.length) return;
  const cols = [
    'Przedstawiciel handlowy',
    'Kod klienta',
    'Nazwa klienta',
    'Rabat',
    'Szczegóły rabatu',
    'Sieć obca',
    'INDEKS',
    'NAZWA',
    'Producent',
    'Tydzień',
    'Sprzedaż',
    'Trend',
    'Poprzedni tydzień',
    'Sprzedaż poprzedni tydzień'
  ];
  const rows = [
    cols,
    ...exportRows.map(row => [
      row.representative,
      row.customerCode,
      row.customerName,
      row.discountLabel || '',
      row.discountDetails || '',
      row.discountNetwork || '',
      row.index,
      row.name,
      row.producer,
      row.week,
      row.quantity || row.value,
      Number.isFinite(row.trendValue) ? row.trendValue : '',
      row.previousWeek || '',
      Number.isFinite(row.previousValue) ? row.previousValue : ''
    ])
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'sprzedaz per tydzien');
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  downloadBlob(blob, 'sprzedaz_per_indeks_per_tydzien.xlsx');
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
    { name: 'Ukraina - Przyprawy', type: 'json', url: jsonUrlPrzyprawyUkraina }
  ];

  const results = [];
  for(const src of sources){
    try{
      const data = src.type === 'xlsx' ? await loadXlsxAsJson(src.url) : await (await fetch(src.url)).json();
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
  datasets.forEach(ds => {
    const cols = Object.keys(ds.data[0] || {});
    const eanKey = findColumn(cols, ['kod ean', 'ean']);
    const indexKey = findColumn(cols, ['indeks', 'index', 'id', 'numer katalogowy']);
    const nameKey = findColumn(cols, ['nazwa', 'name']);
    const producerKey = findColumn(cols, ['producent', 'producer']);
    const groupKey = findColumn(cols, ['grupa', 'group']);
    const rankingKey = findColumn(cols, ['ranking']);
    ds.data.forEach(row => {
      const eanVal = eanKey ? String(row[eanKey] ?? '').replace(/\D/g, '') : '';
      const idxVal = indexKey ? String(row[indexKey] ?? '').replace(/\D/g, '') : '';
      if(eanVal === searchCode || idxVal === searchCode){
        const imageUrl = buildListingImageUrl(ds.name, idxVal || row[indexKey], row);
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
    });
  });
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
  listingHead.innerHTML = cols.map(c => `<th>${escapeHtml(c)}</th>`).join('');
  listingBody.innerHTML = listingResults.map(r => `
    <tr>
      ${cols.map(c => c === 'Zdjęcie'
        ? `<td>${r[c] ? buildDeferredImageTag(String(r[c]), 'listing-img') : ''}</td>`
        : `<td>${escapeHtml(r[c] ?? '')}</td>`).join('')}
    </tr>
  `).join('');
  registerLazyImages(listingBody);
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
    const dataUrl = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
    const type = blob.type || '';
    const format = type.includes('png') ? 'PNG' : 'JPEG';
    return { dataUrl, format };
  }catch(e){
    return null;
  }
}


function buildListingImageUrl(source, indexValue, row){
  const idx = String(indexValue ?? '').replace(/\D/g, '');
  if(!idx) return '';
  if(String(source).toLowerCase().includes('ukraina')){
    return `${imageBaseUrlUkraina}${encodeURIComponent(idx)}.png?alt=media`;
  }
  return `${imageBaseUrlRumunia}${encodeURIComponent(idx)}.png?alt=media`;
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
  if(!row) return currentImageBaseUrl;
  const country = String(row.__country || row.__source || '').toLowerCase();
  if(country.includes('ukraina')) return imageBaseUrlUkraina;
  return imageBaseUrlRumunia;
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
    { name: 'Ukraina - Przyprawy', type: 'json', url: jsonUrlPrzyprawyUkraina }
  ];

  const allRows = [];
  const columns = ['INDEKS', 'NAZWA', 'RANKING', 'GRUPA', 'PRODUCENT', 'KOD EAN'];

  for(const src of sources){
    try{
      const data = src.type === 'xlsx'
        ? await loadXlsxAsJson(src.url)
        : await (await fetch(src.url)).json();
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

async function createCatalog(){
  viewMode = 'catalog';
  catalogLoading = true;
  catalogBlobUrl = null;
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
    return `<th class="index-header">
              <label class="select-box select-box--header">
                <input type="checkbox" onclick="toggleSelectAll(this)" ${checked}>
                <span></span>
              </label>
              ${escapeHtml(col)}
            </th>`;
  }
  return `<th>${escapeHtml(col)}</th>`;
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
  catalogPriceMap = null;
  catalogPriceRows = null;
  if(catalogError) catalogError.textContent = '';
  catalogModal.classList.remove('hidden');
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

function buildPriceMap(rows, config){
  if(!rows || !rows.length) return null;
  const headerRow = rows[0] || [];
  const header = headerRow.map(h => String(h).toLowerCase().trim());
  const hasHeader =
    header.some(h => h.includes('indeks')) ||
    header.some(h => h.includes('cena')) ||
    header.some(h => h.includes('price'));

  const idxCol = typeof config?.indexCol === 'number'
    ? config.indexCol
    : (hasHeader ? header.indexOf('indeks') : 0);
  const priceCol = typeof config?.priceCol === 'number'
    ? config.priceCol
    : (hasHeader ? header.indexOf('cena') : 1);
  const unitCol = hasHeader ? header.findIndex(h => h.includes('jednostka')) : -1;

  if(idxCol === -1 || priceCol === -1) return null;
  const map = new Map();
  const startRow = hasHeader ? 1 : 0;

  for(let i = startRow; i < rows.length; i++){
    const r = rows[i];
    const indexRaw = r[idxCol];
    const index = normalizeIndexValue(indexRaw);
    if(!index) continue;
    let priceRaw = r[priceCol];
    let price = String(priceRaw ?? '').trim();
    const numeric = Number(priceRaw);
    if(Number.isFinite(numeric)){
      price = numeric.toFixed(2);
    }else if(price){
      const parsed = Number(String(price).replace(',', '.'));
      if(Number.isFinite(parsed)) price = parsed.toFixed(2);
    }
    const unit = unitCol !== -1 ? String(r[unitCol] || '').trim() : '';
    map.set(index, { price, unit });
  }
  return map;
}

function populateCatalogColumnOptions(rows){
  if(!catalogIndexColSelect || !catalogPriceColSelect) return;
  const colCount = Math.max(...rows.map(r => r.length));
  const header = rows[0] || [];
  const makeLabel = (i) => {
    const name = header[i] ? ` (${header[i]})` : '';
    return `${columnLabel(i)}${name}`;
  };
  const options = Array.from({ length: colCount }, (_, i) =>
    `<option value="${i}">${makeLabel(i)}</option>`
  ).join('');
  catalogIndexColSelect.innerHTML = `<option value="auto">Auto</option>${options}`;
  catalogPriceColSelect.innerHTML = `<option value="auto">Auto</option>${options}`;
}

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

function isCurrentUserAdmin(){
  return Boolean(currentUserIsAdmin);
}

function toggleAdminPanel(email){
  const isAdminEmail = ADMIN_EMAILS.includes(String(email || '').toLowerCase());
  currentUserIsAdmin = isAdminEmail;
  if(adminPanelBtn) adminPanelBtn.classList.toggle('hidden', !isAdminEmail);
  if(reportsCard) reportsCard.classList.toggle('hidden', !isAdminEmail);
  if(!isAdminEmail){
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
    return `<td class="index-cell">
              <label class="select-box">
                <input type="checkbox" data-index="${escapeAttr(idx)}" onchange="toggleProductSelection(this)" ${checked}>
                <span></span>
              </label>
              <span class="row-number">${rowNumber || ''}.</span>
              ${img}
              <span>${escapeHtml(idx)}</span>
            </td>`;
  }
  if(eanKey && col === eanKey){
    const raw = String(value ?? '').trim();
    const normalized = normalizeEanForBarcode(raw);
    if(normalized){
      const barcode = getBarcodeDataUrl(normalized);
      return `<td class="ean-cell">
                ${barcode ? `<img class="ean-barcode" src="${barcode}" alt="${escapeAttr(normalized.code)}">` : ''}
                <div class="ean-text">${escapeHtml(normalized.code)}</div>
              </td>`;
    }
  }
  return `<td>${escapeHtml(value ?? '')}</td>`;
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
  return `${base}${encodeURIComponent(index)}.${ext}?alt=media`;
}

function createImageExtensionCache(){
  if(typeof localStorage === 'undefined') return new Map();
  try{
    const raw = localStorage.getItem(IMAGE_EXTENSION_CACHE_KEY);
    if(!raw) return new Map();
    const parsed = JSON.parse(raw);
    if(!parsed || typeof parsed !== 'object') return new Map();
    return new Map(
      Object.entries(parsed).filter(([key, value]) => key && (value === 'png' || value === 'jpg'))
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

function getPreferredImageExt(index, baseUrl){
  const key = getImageCacheKey(index, baseUrl);
  return (key && imageExtensionCache.get(key)) || 'png';
}

function rememberImageExt(index, baseUrl, ext){
  const normalizedExt = ext === 'jpg' ? 'jpg' : 'png';
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
  const attrs = `data-index="${escapeAttr(index)}" data-base="${escapeAttr(base)}" data-tried="${ext}"${options.attrs ? ` ${options.attrs}` : ''}`;
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

function loadDeferredImage(img){
  if(!img) return;
  const src = img.getAttribute('data-src');
  const state = img.getAttribute('data-state');
  if(!src || state === 'loading' || state === 'loaded') return;
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
  const tried = img.getAttribute('data-tried') || 'png';
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
  img.setAttribute('data-tried', ext);
  img.setAttribute('data-src', buildImageUrl(index, ext, base));
  loadDeferredImage(img);
}

function imageFallback(img){
  const index = img.getAttribute('data-index');
  const base = img.getAttribute('data-base') || currentImageBaseUrl;
  const tried = img.getAttribute('data-tried');
  const pop = img.closest('.img-hover')?.querySelector('.index-img-large');
  if(tried === 'png'){
    const nextUrl = buildImageUrl(index, 'jpg', base);
    img.setAttribute('data-tried', 'jpg');
    img.setAttribute('data-src', nextUrl);
    img.setAttribute('data-state', 'fallback');
    img.src = nextUrl;
    if(pop){
      pop.setAttribute('data-tried', 'jpg');
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
