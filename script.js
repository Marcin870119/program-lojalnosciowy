const jsonUrl =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Slodycze%20Ranking%20Rumunia.json';
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

const loginView = document.getElementById('login-view');
const appView = document.getElementById('app-view');
const loginEmail = document.getElementById('login-email');
const loginPassword = document.getElementById('login-password');
const loginBtn = document.getElementById('login-btn');
const logoutBtn = document.getElementById('logout-btn');
const loginError = document.getElementById('login-error');
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

let auth = null;
let authReady = false;
let lastLoginEmail = '';
let lastLoginPassword = '';

const REPORT_GROUP_CONFIGS = [
  {
    id: 'slodycze',
    label: 'SLODYCZE I PRZEKASKI RUMUNIA - SLODYCZE RUMUNIA',
    groups: ['SŁODYCZE I PRZEKĄSKI RUMUNIA', 'SŁODYCZE RUMUNIA'],
    sources: [jsonUrl]
  },
  {
    id: 'puszki',
    label: 'PUSZKI I SLOIKI CHLODZONE RUMUNIA - PUSZKI I SLOIKI RUMUNIA',
    groups: ['PUSZKI I SŁOIKI CHŁODZONE RUMUNIA', 'PUSZKI I SŁOIKI RUMUNIA'],
    sources: [jsonUrlPuszkiSloiki]
  },
  {
    id: 'kawa',
    label: 'HERBATA I KAWA RUMUNIA',
    groups: ['HERBATA I KAWA RUMUNIA'],
    sources: [jsonUrlKawyHerbaty]
  },
  {
    id: 'mieso',
    label: 'MIESO I WEDLINY RUMUNIA',
    groups: ['MIĘSO I WĘDLINY RUMUNIA'],
    sources: [jsonUrlMieso]
  },
  {
    id: 'nabial',
    label: 'NABIAL RUMUNIA - RUMUNIA CHLODNIA POZOSTALE',
    groups: ['NABIAŁ RUMUNIA', 'RUMUNIA CHŁODNIA POZOSTAŁE'],
    sources: [jsonUrlNabial]
  },
  {
    id: 'napoje',
    label: 'NAPOJE BEZALKOHOLOWE RUMUNIA',
    groups: ['NAPOJE BEZALKOHOLOWE RUMUNIA'],
    sources: [jsonUrlNapoje]
  },
  {
    id: 'podstawowe',
    label: 'PRODUKTY PODSTAWOWE RUMUNIA',
    groups: ['PRODUKTY PODSTAWOWE RUMUNIA'],
    sources: [jsonUrlProduktyPodstawowe]
  },
  {
    id: 'przyprawy',
    label: 'PRZYPRAWY I DODATKI W PROSZKU RUMUNIA',
    groups: ['PRZYPRAWY I DODATKI W PROSZKU RUMUNIA'],
    sources: [jsonUrlPrzyprawyProszek]
  }
];

function createDefaultReportLimits(){
  return REPORT_GROUP_CONFIGS.reduce((acc, group) => {
    acc[group.id] = 25;
    return acc;
  }, {});
}

let reportsGroupLimits = createDefaultReportLimits();

function showApp(){
  if(loginView) loginView.classList.add('hidden');
  if(appView) appView.classList.remove('hidden');
  showSecurityModalOnce();
}

function showLogin(){
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
  topSuggestionSourceRows = [];
  topSuggestionGeneratedRows = [];
  topSuggestionPurchasedRows = [];
  topSuggestionImportFile = '';
  topSuggestionSelectedRepresentative = '';
  topSuggestionSelectedCustomer = '';
  topSuggestionLimit = '';
  topSuggestionSummary = null;
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

if(typeof firebase !== 'undefined'){
  if(String(firebaseConfig.apiKey).startsWith('PASTE_')){
    setLoginError('Uzupełnij firebaseConfig w script.js');
  }else{
    if(!firebase.apps.length){
      firebase.initializeApp(firebaseConfig);
    }
    auth = firebase.auth();
    authReady = true;

    auth.onAuthStateChanged(user => {
      if(user){
        showApp();
        toggleAdminPanel(user.email || '');
        resetAppState();
        setLoginError('');
      }else{
        toggleAdminPanel('');
        resetAppState();
        showLogin();
      }
    });
  }
}else{
  setLoginError('Firebase nie załadował się. Sprawdź połączenie.');
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
    console.log('LOGIN CLICK');
    console.log('EMAIL:', loginEmail.value);
    console.log('AUTH READY:', authReady, auth);
    try{
      await auth.signInWithEmailAndPassword(loginEmail.value.trim(), loginPassword.value);
      showApp();
      toggleAdminPanel(lastLoginEmail);
      resetAppState();
    }catch(err){
      console.error('LOGIN ERROR:', err.code, err.message, err);
      setLoginError(err.code || err.message || 'Błąd logowania');
    }
  });
}


if(logoutBtn){
  logoutBtn.addEventListener('click', async () => {
    try{
      if(auth){
        await auth.signOut();
      }
    }catch(e){
      console.error(e);
    }finally{
      localStorage.removeItem('is_admin');
      toggleAdminPanel('');
      resetAppState();
      showLogin();
    }
  });
}

if(loginPassword){
  loginPassword.addEventListener('keydown', (e) => {
    if(e.key === 'Enter' && loginBtn){
      loginBtn.click();
    }
  });
}

if(appBrandLink){
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

// OBSŁUGA ZAKŁADEK
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    tab.classList.add('active');

    document.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));
    document.getElementById(tab.dataset.tab).classList.remove('hidden');
    currentImageBaseUrl = tab.dataset.tab === 'ukraina' ? imageBaseUrlUkraina : imageBaseUrlRumunia;
    window.imageBaseUrl = currentImageBaseUrl;

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
    viewMode = 'table';

    if(tab.dataset.tab !== 'listing'){
      stopListingScanner();
    }
  });
});

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

if(reportsCard){
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
    const imgUrl = idx ? buildImageUrl(idx, 'png', imageBaseUrlRumunia) : '';
    return `
      <tr>
        <td>${index + 1}</td>
        <td class="reports-image-cell">
          ${imgUrl ? `<img class="reports-image" src="${imgUrl}" data-index="${escapeAttr(idx)}" data-base="${escapeAttr(imageBaseUrlRumunia)}" data-tried="png" onerror="imageFallback(this)" alt="">` : ''}
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
        ${reportsImportFile ? `Źródło: ${escapeHtml(reportsImportFile)}` : 'Brak zaimportowanego pliku'}
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
    }).filter(row => row.value);
  });
}

function getWeeklySalesRepresentatives(){
  return Array.from(new Set(weeklySalesSourceRows.map(row => row.representative).filter(Boolean)))
    .sort((a, b) => a.localeCompare(b, 'pl'));
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

function getWeeklySalesWeeks(){
  const rows = weeklySalesGeneratedRows.length ? weeklySalesGeneratedRows : weeklySalesSourceRows;
  return Array.from(new Set(
    rows.map(row => String(row.week || '').trim()).filter(Boolean)
  )).sort((a, b) => a.localeCompare(b, 'pl', { numeric: true }));
}

function getLatestWeeklySalesWeek(){
  const weeks = getWeeklySalesWeeks();
  return weeks[weeks.length - 1] || '';
}

function getWeeklySalesComparisonWeeks(){
  const weeks = Array.from(new Set(
    weeklySalesSourceRows.map(row => String(row.week || '').trim()).filter(Boolean)
  )).sort((a, b) => a.localeCompare(b, 'pl', { numeric: true }));
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

function getFilteredWeeklySalesRows(){
  let rows = weeklySalesGeneratedRows;
  const latestWeek = getLatestWeeklySalesWeek();
  const activeWeek = weeklySalesSelectedWeek || latestWeek;

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
  }else if(activeWeek){
    rows = rows.filter(row => String(row.week || '') === activeWeek);
  }

  if(weeklySalesSelectedProducer){
    rows = rows.filter(row => String(row.producer || '').trim() === weeklySalesSelectedProducer);
  }

  return rows;
}

function getLatestWeeksFromRows(rows, count){
  const weeks = Array.from(new Set(rows.map(row => String(row.week || '').trim()).filter(Boolean)))
    .sort((a, b) => a.localeCompare(b, 'pl', { numeric: true }));
  return weeks.slice(Math.max(weeks.length - count, 0));
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
    const imgUrl = idx ? buildImageUrl(idx, 'png', imageBaseUrlRumunia) : '';
    const normalizedIndex = normalizeIndexValue(idx);
    const checked = selectable && clientReportSelectedMissingIndexes.has(normalizedIndex) ? 'checked' : '';
    return `
      <tr>
        <td>${index + 1}</td>
        ${selectable ? `<td><input type="checkbox" data-index="${escapeAttr(idx)}" onchange="toggleClientMissingSelection(this)" ${checked}></td>` : ''}
        <td class="reports-image-cell">
          ${imgUrl ? `<img class="reports-image" src="${imgUrl}" data-index="${escapeAttr(idx)}" data-base="${escapeAttr(imageBaseUrlRumunia)}" data-tried="png" onerror="imageFallback(this)" alt="">` : ''}
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
        ${clientReportImportFile ? `Źródło: ${escapeHtml(clientReportImportFile)}` : 'Brak zaimportowanego pliku'}
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
  const filteredRows = getFilteredWeeklySalesRows();
  const comparisonRows = getWeeklySalesRepresentativeComparisonRows();
  const comparisonWeeks = getWeeklySalesComparisonWeeks();
  const activeRowsCount = weeklySalesRepComparison ? comparisonRows.length : filteredRows.length;
  const activeSalesTotal = weeklySalesRepComparison
    ? comparisonRows.reduce((sum, row) => sum + row.latestValue, 0)
    : filteredRows.reduce((sum, row) => sum + (row.quantity || row.value), 0);
  const previousSalesTotal = comparisonRows.reduce((sum, row) => sum + row.previousValue, 0);
  const comparisonTrendValue = activeSalesTotal - previousSalesTotal;
  const customerLabel = weeklySalesSelectedCustomer ? 'wybrany klient' : 'wszyscy klienci';
  const weekLabel = weeklySalesOnlyLastWeek250
    ? `ostatni tydzień ${latestWeek || ''}, min. 250 GBP`
    : (weeklySalesSelectedWeek || latestWeek || 'brak tygodnia');
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
    return `
      <tr>
        <td>${index + 1}</td>
        <td>${escapeHtml(row.week)}</td>
        <td>${escapeHtml(row.customerCode)}</td>
        <td>${escapeHtml(row.customerName)}</td>
        <td>${discountHtml}</td>
        <td>${escapeHtml(row.index)}</td>
        <td>${escapeHtml(row.name)}</td>
        <td>${escapeHtml(row.producer)}</td>
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
              <th>INDEKS</th>
              <th>NAZWA</th>
              <th>Producent</th>
              <th>Sprzedaż</th>
              <th>Trend</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml || `<tr><td colspan="10" class="reports-empty">Zaimportuj plik Excel i wybierz przedstawiciela.</td></tr>`}
          </tbody>
        </table>
      </div>
    `;

  return `
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
        ${weeklySalesImportFile ? `Źródło: ${escapeHtml(weeklySalesImportFile)}` : 'Brak zaimportowanego pliku'}
      </div>
    </div>

    ${weeklySalesRepComparison ? '' : `<div class="reports-filters">
      <label class="reports-filter-card">
        <span class="reports-limit-title">Przedstawiciel handlowy</span>
        <select onchange="setWeeklySalesRepresentative(this.value)">
          <option value="">Wybierz</option>
          ${representatives.map(rep => `<option value="${escapeAttr(rep)}" ${weeklySalesSelectedRepresentative === rep ? 'selected' : ''}>${escapeHtml(rep)}</option>`).join('')}
        </select>
      </label>
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
      <label class="reports-filter-card">
        <span class="reports-limit-title">Producent</span>
        <select onchange="setWeeklySalesProducer(this.value)" ${weeklySalesGeneratedRows.length ? '' : 'disabled'}>
          <option value="">Wszyscy producenci</option>
          ${producers.map(producer => `<option value="${escapeAttr(producer)}" ${weeklySalesSelectedProducer === producer ? 'selected' : ''}>${escapeHtml(producer)}</option>`).join('')}
        </select>
      </label>
      <label class="reports-filter-card">
        <span class="reports-limit-title">Tydzień sprzedaży</span>
        <select onchange="setWeeklySalesWeek(this.value)" ${weeklySalesGeneratedRows.length || weeklySalesOnlyLastWeek250 ? '' : 'disabled'}>
          <option value="">Ostatni tydzień</option>
          ${weeks.map(week => `<option value="${escapeAttr(week)}" ${weeklySalesSelectedWeek === week ? 'selected' : ''}>${escapeHtml(week)}</option>`).join('')}
        </select>
      </label>
    </div>`}

    <div class="reports-inline-actions">
      <button class="btn-outline ${weeklySalesOnlyLastWeek250 ? 'reports-mode-active' : ''}" onclick="setWeeklySalesLastWeek250(!weeklySalesOnlyLastWeek250)" ${weeklySalesGeneratedRows.length ? '' : 'disabled'}>Klienci >= 250 GBP w ostatnim tygodniu</button>
      <button class="btn-outline ${weeklySalesRepComparison ? 'reports-mode-active' : ''}" onclick="setWeeklySalesRepComparison(!weeklySalesRepComparison)" ${weeklySalesSourceRows.length ? '' : 'disabled'}>Porównanie PH tydzień do tygodnia</button>
    </div>

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
        ${topSuggestionImportFile ? `Źródło: ${escapeHtml(topSuggestionImportFile)}` : 'Brak zaimportowanego pliku'}
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

function renderReportsView(){
  if(!reportsContainer) return;
  const contentByMode = {
    'top-sales': renderTopSalesReportContent,
    'client-gap': renderClientComparisonContent,
    'weekly-sales': renderWeeklySalesContent,
    'top-suggestions': renderTopSuggestionsContent
  };
  const renderContent = contentByMode[reportsMode] || renderTopSalesReportContent;

  reportsContainer.innerHTML = `
    <div class="reports-panel">
      <div class="reports-mode-switch">
        <button class="btn-outline ${reportsMode === 'top-sales' ? 'reports-mode-active' : ''}" onclick="setReportsMode('top-sales')">Top sprzedaż</button>
        <button class="btn-outline ${reportsMode === 'client-gap' ? 'reports-mode-active' : ''}" onclick="setReportsMode('client-gap')">Potencjał klienta Top Rumunia</button>
        <button class="btn-outline ${reportsMode === 'weekly-sales' ? 'reports-mode-active' : ''}" onclick="setReportsMode('weekly-sales')">Sprzedaż tygodniowa</button>
        <button class="btn-outline ${reportsMode === 'top-suggestions' ? 'reports-mode-active' : ''}" onclick="setReportsMode('top-suggestions')">Propozycje Top</button>
      </div>
      ${renderContent()}
    </div>
  `;
}

function openReportsView(){
  setActiveCard('reports-card');
  setImageBase(imageBaseUrlRumunia);
  setFullData([]);
  isLoading = false;
  activeContainer = reportsContainer;
  currentCategoryName = 'Raporty';
  currentCategorySlug = 'raporty';
  resetFilters();
  clearAllContentContainers();
  renderReportsView();
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

function setReportsMode(mode){
  reportsMode = mode;
  renderReportsView();
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

function setWeeklySalesLastWeek250(enabled){
  weeklySalesOnlyLastWeek250 = !!enabled;
  weeklySalesRepComparison = false;
  if(weeklySalesOnlyLastWeek250){
    weeklySalesSelectedWeek = getLatestWeeklySalesWeek();
  }
  renderReportsView();
}

function setWeeklySalesRepComparison(enabled){
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
        ? `<td>${r[c] ? `<img class="listing-img" src="${escapeAttr(r[c])}" alt="">` : ''}</td>`
        : `<td>${escapeHtml(r[c] ?? '')}</td>`).join('')}
    </tr>
  `).join('');
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

function toggleAdminPanel(email){
  const isAdminEmail = ADMIN_EMAILS.includes(String(email || '').toLowerCase());
  const isAdminSession = localStorage.getItem('is_admin') === '1';
  const isAdmin = isAdminEmail || (isAdminSession && isAdminEmail);
  if(adminPanelBtn) adminPanelBtn.classList.toggle('hidden', !isAdmin);
  if(reportsCard) reportsCard.classList.toggle('hidden', !isAdmin);
  if(isAdmin){
    localStorage.setItem('is_admin', '1');
  }else{
    localStorage.removeItem('is_admin');
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
           <img class="index-img" src="${buildImageUrl(idx, 'png', imgBase)}" data-index="${escapeAttr(idx)}" data-base="${escapeAttr(imgBase)}" data-tried="png" onerror="imageFallback(this)" alt="">
           <span class="img-pop">
             <img class="index-img-large" src="${buildImageUrl(idx, 'png', imgBase)}" alt="">
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

function imageFallback(img){
  const index = img.getAttribute('data-index');
  const base = img.getAttribute('data-base') || currentImageBaseUrl;
  const tried = img.getAttribute('data-tried');
  const pop = img.closest('.img-hover')?.querySelector('.index-img-large');
  if(tried === 'png'){
    img.setAttribute('data-tried', 'jpg');
    img.src = buildImageUrl(index, 'jpg', base);
    if(pop) pop.src = buildImageUrl(index, 'jpg', base);
    return;
  }
  img.classList.add('img-missing');
  if(pop) pop.classList.add('img-missing');
}

function positionPopup(el){
  const pop = el.querySelector('.img-pop');
  if(!pop) return;
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
