const jsonUrl =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Slodycze%20Ranking%20Rumunia.json';
const jsonUrlMieso =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Mieso%20Wedliny%20-%20Rumunia.json';
const xlsxUrlNabial =
  'https://raw.githubusercontent.com/Marcin870119/program-lojalnosciowy/main/Nabial%20-%20Rumunia%20Ranking.xlsx';
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
let currentImageBaseUrl = imageBaseUrlRumunia;
window.imageBaseUrl = currentImageBaseUrl;

const firebaseConfig = {
  apiKey: 'AIzaSyDwzVRS5W2lklGMLcZJn-YPCK9OtBQZ7bI',
  authDomain: 'pdf-creator-f7a8b.firebaseapp.com',
  projectId: 'pdf-creator-f7a8b',
  appId: '1:606744201676:web:6f8c1b2c323fbaf6f3b569'
};

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
const ukrainaSlodyczeContainer = document.getElementById('ukraina-slodycze-content');
const ukrainaMiesoContainer = document.getElementById('ukraina-mieso-wedliny-content');
const ukrainaKawyContainer = document.getElementById('ukraina-kawy-herbaty-content');
const ukrainaPuszkiContainer = document.getElementById('ukraina-puszki-sloiki-content');
const ukrainaNapojeContainer = document.getElementById('ukraina-napoje-content');
const ukrainaPrzyprawyContainer = document.getElementById('ukraina-przyprawy-proszek-content');

var fullData = [];
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
const LIMIT = 25;
const barcodeCache = new Map();
let importedIndexSet = null;
let importedIndexCount = 0;
let importedIndexFile = '';
let listingResults = [];
let listingAllDataCache = null;
let listingScannedCodes = new Set();
let listingCooldown = false;

let auth = null;
let authReady = false;
let lastLoginEmail = '';
let lastLoginPassword = '';

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
  importedIndexSet = null;
  importedIndexCount = 0;
  importedIndexFile = '';
  resetFilters();

  slodyczeContainer.innerHTML = '';
  miesoContainer.innerHTML = '';
  nabialContainer.innerHTML = '';
  napojeContainer.innerHTML = '';
  przyprawyProszekContainer.innerHTML = '';
  puszkiSloikiContainer.innerHTML = '';
  produktyPodstawoweContainer.innerHTML = '';
  kawyHerbatyContainer.innerHTML = '';
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
      return;
    }
    if(priceWarningModal) priceWarningModal.classList.remove('hidden');
    try{
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      catalogPriceMap = buildPriceMap(rows);
    }catch(e){
      console.error(e);
      catalogPriceMap = null;
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
  const res = await fetch(jsonUrl);
  fullData = await res.json();
  expanded = false;
  viewMode = 'table';
  activeContainer = slodyczeContainer;
  currentCategoryName = 'Slodycze';
  currentCategorySlug = 'slodycze';
  resetFilters();
  slodyczeContainer.innerHTML = '';
  miesoContainer.innerHTML = '';
  nabialContainer.innerHTML = '';
  napojeContainer.innerHTML = '';
  przyprawyProszekContainer.innerHTML = '';
  puszkiSloikiContainer.innerHTML = '';
  produktyPodstawoweContainer.innerHTML = '';
  kawyHerbatyContainer.innerHTML = '';
  if(ukrainaSlodyczeContainer) ukrainaSlodyczeContainer.innerHTML = '';
  if(ukrainaMiesoContainer) ukrainaMiesoContainer.innerHTML = '';
  if(ukrainaKawyContainer) ukrainaKawyContainer.innerHTML = '';
  if(ukrainaPuszkiContainer) ukrainaPuszkiContainer.innerHTML = '';
  if(ukrainaNapojeContainer) ukrainaNapojeContainer.innerHTML = '';
  if(ukrainaPrzyprawyContainer) ukrainaPrzyprawyContainer.innerHTML = '';
  render();
});

// KLIK KAFELKA "SŁODYCZE" - UKRAINA
const ukrainaSlodyczeCard = document.getElementById('ukraina-slodycze');
if(ukrainaSlodyczeCard){
  ukrainaSlodyczeCard.addEventListener('click', async () => {
    setActiveCard('ukraina-slodycze');
    const res = await fetch(jsonUrlSlodyczeUkraina);
    fullData = await res.json();
    expanded = false;
    viewMode = 'table';
    activeContainer = ukrainaSlodyczeContainer;
    currentCategoryName = 'Slodycze_Ukraina';
    currentCategorySlug = 'slodycze_ukraina';
    resetFilters();
    slodyczeContainer.innerHTML = '';
    miesoContainer.innerHTML = '';
    nabialContainer.innerHTML = '';
    napojeContainer.innerHTML = '';
    przyprawyProszekContainer.innerHTML = '';
    puszkiSloikiContainer.innerHTML = '';
    produktyPodstawoweContainer.innerHTML = '';
    kawyHerbatyContainer.innerHTML = '';
    if(ukrainaSlodyczeContainer) ukrainaSlodyczeContainer.innerHTML = '';
    if(ukrainaMiesoContainer) ukrainaMiesoContainer.innerHTML = '';
    render();
  });
}

// KLIK KAFELKA "MIĘSO I WĘDLINY" - UKRAINA
const ukrainaMiesoCard = document.getElementById('ukraina-mieso-wedliny');
if(ukrainaMiesoCard){
  ukrainaMiesoCard.addEventListener('click', async () => {
    setActiveCard('ukraina-mieso-wedliny');
    const res = await fetch(jsonUrlMiesoUkraina);
    fullData = await res.json();
    expanded = false;
    viewMode = 'table';
    activeContainer = ukrainaMiesoContainer;
    currentCategoryName = 'Mieso_i_wedliny_Ukraina';
    currentCategorySlug = 'mieso_i_wedliny_ukraina';
    resetFilters();
    slodyczeContainer.innerHTML = '';
    miesoContainer.innerHTML = '';
    nabialContainer.innerHTML = '';
    napojeContainer.innerHTML = '';
    przyprawyProszekContainer.innerHTML = '';
    puszkiSloikiContainer.innerHTML = '';
    produktyPodstawoweContainer.innerHTML = '';
    kawyHerbatyContainer.innerHTML = '';
    if(ukrainaSlodyczeContainer) ukrainaSlodyczeContainer.innerHTML = '';
    if(ukrainaMiesoContainer) ukrainaMiesoContainer.innerHTML = '';
    if(ukrainaKawyContainer) ukrainaKawyContainer.innerHTML = '';
    render();
  });
}

// KLIK KAFELKA "KAWY I HERBATY" - UKRAINA
const ukrainaKawyCard = document.getElementById('ukraina-kawy-herbaty');
if(ukrainaKawyCard){
  ukrainaKawyCard.addEventListener('click', async () => {
    setActiveCard('ukraina-kawy-herbaty');
    const res = await fetch(jsonUrlKawyUkraina);
    fullData = await res.json();
    expanded = false;
    viewMode = 'table';
    activeContainer = ukrainaKawyContainer;
    currentCategoryName = 'Kawy_i_herbaty_Ukraina';
    currentCategorySlug = 'kawy_i_herbaty_ukraina';
    resetFilters();
    slodyczeContainer.innerHTML = '';
    miesoContainer.innerHTML = '';
    nabialContainer.innerHTML = '';
    napojeContainer.innerHTML = '';
    przyprawyProszekContainer.innerHTML = '';
    puszkiSloikiContainer.innerHTML = '';
    produktyPodstawoweContainer.innerHTML = '';
    kawyHerbatyContainer.innerHTML = '';
    if(ukrainaSlodyczeContainer) ukrainaSlodyczeContainer.innerHTML = '';
    if(ukrainaMiesoContainer) ukrainaMiesoContainer.innerHTML = '';
    if(ukrainaKawyContainer) ukrainaKawyContainer.innerHTML = '';
    if(ukrainaPuszkiContainer) ukrainaPuszkiContainer.innerHTML = '';
    render();
  });
}

// KLIK KAFELKA "PUSZKI I SŁOIKI" - UKRAINA
const ukrainaPuszkiCard = document.getElementById('ukraina-puszki-sloiki');
if(ukrainaPuszkiCard){
  ukrainaPuszkiCard.addEventListener('click', async () => {
    setActiveCard('ukraina-puszki-sloiki');
    fullData = await loadXlsxAsJson(xlsxUrlPuszkiUkraina);
    expanded = false;
    viewMode = 'table';
    activeContainer = ukrainaPuszkiContainer;
    currentCategoryName = 'Puszki_i_sloiki_Ukraina';
    currentCategorySlug = 'puszki_i_sloiki_ukraina';
    resetFilters();
    slodyczeContainer.innerHTML = '';
    miesoContainer.innerHTML = '';
    nabialContainer.innerHTML = '';
    napojeContainer.innerHTML = '';
    przyprawyProszekContainer.innerHTML = '';
    puszkiSloikiContainer.innerHTML = '';
    produktyPodstawoweContainer.innerHTML = '';
    kawyHerbatyContainer.innerHTML = '';
    if(ukrainaSlodyczeContainer) ukrainaSlodyczeContainer.innerHTML = '';
    if(ukrainaMiesoContainer) ukrainaMiesoContainer.innerHTML = '';
    if(ukrainaKawyContainer) ukrainaKawyContainer.innerHTML = '';
    if(ukrainaPuszkiContainer) ukrainaPuszkiContainer.innerHTML = '';
    if(ukrainaNapojeContainer) ukrainaNapojeContainer.innerHTML = '';
    render();
  });
}

// KLIK KAFELKA "NAPOJE" - UKRAINA
const ukrainaNapojeCard = document.getElementById('ukraina-napoje');
if(ukrainaNapojeCard){
  ukrainaNapojeCard.addEventListener('click', async () => {
    setActiveCard('ukraina-napoje');
    const res = await fetch(jsonUrlNapojeUkraina);
    fullData = await res.json();
    expanded = false;
    viewMode = 'table';
    activeContainer = ukrainaNapojeContainer;
    currentCategoryName = 'Napoje_Ukraina';
    currentCategorySlug = 'napoje_ukraina';
    resetFilters();
    slodyczeContainer.innerHTML = '';
    miesoContainer.innerHTML = '';
    nabialContainer.innerHTML = '';
    napojeContainer.innerHTML = '';
    przyprawyProszekContainer.innerHTML = '';
    puszkiSloikiContainer.innerHTML = '';
    produktyPodstawoweContainer.innerHTML = '';
    kawyHerbatyContainer.innerHTML = '';
    if(ukrainaSlodyczeContainer) ukrainaSlodyczeContainer.innerHTML = '';
    if(ukrainaMiesoContainer) ukrainaMiesoContainer.innerHTML = '';
    if(ukrainaKawyContainer) ukrainaKawyContainer.innerHTML = '';
    if(ukrainaPuszkiContainer) ukrainaPuszkiContainer.innerHTML = '';
    if(ukrainaNapojeContainer) ukrainaNapojeContainer.innerHTML = '';
    if(ukrainaPrzyprawyContainer) ukrainaPrzyprawyContainer.innerHTML = '';
    render();
  });
}

// KLIK KAFELKA "PRZYPRAWY I DODATKI W PROSZKU" - UKRAINA
const ukrainaPrzyprawyCard = document.getElementById('ukraina-przyprawy-proszek');
if(ukrainaPrzyprawyCard){
  ukrainaPrzyprawyCard.addEventListener('click', async () => {
    setActiveCard('ukraina-przyprawy-proszek');
    const res = await fetch(jsonUrlPrzyprawyUkraina);
    fullData = await res.json();
    expanded = false;
    viewMode = 'table';
    activeContainer = ukrainaPrzyprawyContainer;
    currentCategoryName = 'Przyprawy_Ukraina';
    currentCategorySlug = 'przyprawy_ukraina';
    resetFilters();
    slodyczeContainer.innerHTML = '';
    miesoContainer.innerHTML = '';
    nabialContainer.innerHTML = '';
    napojeContainer.innerHTML = '';
    przyprawyProszekContainer.innerHTML = '';
    puszkiSloikiContainer.innerHTML = '';
    produktyPodstawoweContainer.innerHTML = '';
    kawyHerbatyContainer.innerHTML = '';
    if(ukrainaSlodyczeContainer) ukrainaSlodyczeContainer.innerHTML = '';
    if(ukrainaMiesoContainer) ukrainaMiesoContainer.innerHTML = '';
    if(ukrainaKawyContainer) ukrainaKawyContainer.innerHTML = '';
    if(ukrainaPuszkiContainer) ukrainaPuszkiContainer.innerHTML = '';
    if(ukrainaNapojeContainer) ukrainaNapojeContainer.innerHTML = '';
    if(ukrainaPrzyprawyContainer) ukrainaPrzyprawyContainer.innerHTML = '';
    render();
  });
}

// KLIK KAFELKA "MIĘSO I WĘDLINY"
document.getElementById('rumunia-mieso-wedliny').addEventListener('click', async () => {
  setActiveCard('rumunia-mieso-wedliny');
  const res = await fetch(jsonUrlMieso);
  fullData = await res.json();
  expanded = false;
  viewMode = 'table';
  activeContainer = miesoContainer;
  currentCategoryName = 'Mieso_i_wedliny';
  currentCategorySlug = 'mieso_i_wedliny';
  resetFilters();
  slodyczeContainer.innerHTML = '';
  miesoContainer.innerHTML = '';
  nabialContainer.innerHTML = '';
  napojeContainer.innerHTML = '';
  przyprawyProszekContainer.innerHTML = '';
  puszkiSloikiContainer.innerHTML = '';
  produktyPodstawoweContainer.innerHTML = '';
  kawyHerbatyContainer.innerHTML = '';
  render();
});

// KLIK KAFELKA "NABIAŁ"
document.getElementById('rumunia-nabial').addEventListener('click', async () => {
  setActiveCard('rumunia-nabial');
  fullData = await loadXlsxAsJson(xlsxUrlNabial);
  expanded = false;
  viewMode = 'table';
  activeContainer = nabialContainer;
  currentCategoryName = 'Nabial';
  currentCategorySlug = 'nabial';
  resetFilters();
  slodyczeContainer.innerHTML = '';
  miesoContainer.innerHTML = '';
  nabialContainer.innerHTML = '';
  napojeContainer.innerHTML = '';
  przyprawyProszekContainer.innerHTML = '';
  puszkiSloikiContainer.innerHTML = '';
  produktyPodstawoweContainer.innerHTML = '';
  kawyHerbatyContainer.innerHTML = '';
  render();
});

// KLIK KAFELKA "NAPOJE"
document.getElementById('rumunia-napoje').addEventListener('click', async () => {
  setActiveCard('rumunia-napoje');
  const res = await fetch(jsonUrlNapoje);
  fullData = await res.json();
  expanded = false;
  viewMode = 'table';
  activeContainer = napojeContainer;
  currentCategoryName = 'Napoje';
  currentCategorySlug = 'napoje';
  resetFilters();
  slodyczeContainer.innerHTML = '';
  miesoContainer.innerHTML = '';
  nabialContainer.innerHTML = '';
  napojeContainer.innerHTML = '';
  przyprawyProszekContainer.innerHTML = '';
  puszkiSloikiContainer.innerHTML = '';
  produktyPodstawoweContainer.innerHTML = '';
  kawyHerbatyContainer.innerHTML = '';
  render();
});

// KLIK KAFELKA "PRZYPRAWY I DODATKI W PROSZKU"
document.getElementById('rumunia-przyprawy-proszek').addEventListener('click', async () => {
  setActiveCard('rumunia-przyprawy-proszek');
  const res = await fetch(jsonUrlPrzyprawyProszek);
  fullData = await res.json();
  expanded = false;
  viewMode = 'table';
  activeContainer = przyprawyProszekContainer;
  currentCategoryName = 'Przyprawy_i_dodatki_w_proszku';
  currentCategorySlug = 'przyprawy_i_dodatki_w_proszku';
  resetFilters();
  slodyczeContainer.innerHTML = '';
  miesoContainer.innerHTML = '';
  nabialContainer.innerHTML = '';
  napojeContainer.innerHTML = '';
  przyprawyProszekContainer.innerHTML = '';
  puszkiSloikiContainer.innerHTML = '';
  produktyPodstawoweContainer.innerHTML = '';
  kawyHerbatyContainer.innerHTML = '';
  render();
});

// KLIK KAFELKA "PUSZKI I SŁOIKI"
document.getElementById('rumunia-puszki-sloiki').addEventListener('click', async () => {
  setActiveCard('rumunia-puszki-sloiki');
  const res = await fetch(jsonUrlPuszkiSloiki);
  fullData = await res.json();
  expanded = false;
  viewMode = 'table';
  activeContainer = puszkiSloikiContainer;
  currentCategoryName = 'Puszki_i_sloiki';
  currentCategorySlug = 'puszki_i_sloiki';
  resetFilters();
  slodyczeContainer.innerHTML = '';
  miesoContainer.innerHTML = '';
  nabialContainer.innerHTML = '';
  napojeContainer.innerHTML = '';
  przyprawyProszekContainer.innerHTML = '';
  puszkiSloikiContainer.innerHTML = '';
  produktyPodstawoweContainer.innerHTML = '';
  kawyHerbatyContainer.innerHTML = '';
  render();
});

// KLIK KAFELKA "PRODUKTY PODSTAWOWE"
document.getElementById('rumunia-produkty-podstawowe').addEventListener('click', async () => {
  setActiveCard('rumunia-produkty-podstawowe');
  const res = await fetch(jsonUrlProduktyPodstawowe);
  fullData = await res.json();
  expanded = false;
  viewMode = 'table';
  activeContainer = produktyPodstawoweContainer;
  currentCategoryName = 'Produkty_podstawowe';
  currentCategorySlug = 'produkty_podstawowe';
  resetFilters();
  slodyczeContainer.innerHTML = '';
  miesoContainer.innerHTML = '';
  nabialContainer.innerHTML = '';
  napojeContainer.innerHTML = '';
  przyprawyProszekContainer.innerHTML = '';
  puszkiSloikiContainer.innerHTML = '';
  produktyPodstawoweContainer.innerHTML = '';
  kawyHerbatyContainer.innerHTML = '';
  render();
});

// KLIK KAFELKA "KAWY I HERBATY"
document.getElementById('rumunia-kawy-herbaty').addEventListener('click', async () => {
  setActiveCard('rumunia-kawy-herbaty');
  const res = await fetch(jsonUrlKawyHerbaty);
  fullData = await res.json();
  expanded = false;
  viewMode = 'table';
  activeContainer = kawyHerbatyContainer;
  currentCategoryName = 'Kawy_i_herbaty';
  currentCategorySlug = 'kawy_i_herbaty';
  resetFilters();
  slodyczeContainer.innerHTML = '';
  miesoContainer.innerHTML = '';
  nabialContainer.innerHTML = '';
  napojeContainer.innerHTML = '';
  przyprawyProszekContainer.innerHTML = '';
  puszkiSloikiContainer.innerHTML = '';
  produktyPodstawoweContainer.innerHTML = '';
  kawyHerbatyContainer.innerHTML = '';
  render();
});

function render(){
  if(!fullData.length) return;
  slodyczeContainer.innerHTML = '';
  miesoContainer.innerHTML = '';
  nabialContainer.innerHTML = '';
  napojeContainer.innerHTML = '';
  przyprawyProszekContainer.innerHTML = '';
  puszkiSloikiContainer.innerHTML = '';
  produktyPodstawoweContainer.innerHTML = '';
  kawyHerbatyContainer.innerHTML = '';

  const cols = Object.keys(fullData[0]);
  const producerKey = findColumn(cols, ['producent', 'producer']);
  const groupKey = findColumn(cols, ['grupa', 'group']);
  const nameKey = findColumn(cols, ['nazwa', 'name']);
  const eanKey = findColumn(cols, ['kod ean', 'ean']);
  const indexKey = findColumn(cols, ['indeks', 'index', 'id']);

  let html = `
    ${viewMode === 'table' ? `
      <div class="import-bar">
        <div>
          <div class="import-title">Import Excel</div>
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
    ` : ''}

    <div class="actions">
      <button onclick="exportCSV()">Zapisz CSV</button>
      <button onclick="exportXLS()">Zapisz XLS</button>
      <button onclick="createCatalog()">Utwórz katalog</button>
      ${catalogBlobUrl ? `<button onclick="downloadCatalog()">Zapisz katalog PDF</button>` : ''}
    </div>

    ${viewMode === 'pdf' ? `
      <div class="pdf-preview">
        <iframe src="${pdfPreviewUrl + encodeURIComponent(getActivePdfUrl())}&t=${Date.now()}" title="Podgląd PDF"></iframe>
      </div>
    ` : viewMode === 'catalog' ? `
      <div class="pdf-preview">
        ${catalogLoading ? `<div class="catalog-loading">
          <div class="catalog-spinner"></div>
          <div>Tworzę katalog...</div>
        </div>` : ''}
        ${catalogBlobUrl ? `<iframe src="${catalogBlobUrl}#zoom=50" title="Katalog PDF"></iframe>` : ''}
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
            <tr>${cols.map(c => renderHeaderCell(c, indexKey)).join('')}</tr>
          </thead>
          <tbody id="data-tbody"></tbody>
        </table>
      </div>
      <div id="expand-container"></div>
    `}
  `;

  activeContainer.innerHTML = html;

  if(viewMode === 'table'){
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

const listingCodeInput = document.getElementById('listing-code');
const listingStartBtn = document.getElementById('listing-start-btn');
const listingStopBtn = document.getElementById('listing-stop-btn');
const listingSearchBtn = document.getElementById('listing-search-btn');
const listingExportBtn = document.getElementById('listing-export-btn');
const listingHead = document.getElementById('listing-head');
const listingBody = document.getElementById('listing-body');
const listingConfirmModal = document.getElementById('listing-confirm-modal');
const listingConfirmYes = document.getElementById('listing-confirm-yes');
const listingConfirmNo = document.getElementById('listing-confirm-no');
let pendingListingAdd = null;

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
if(listingCodeInput){
  listingCodeInput.addEventListener('keydown', (e) => {
    if(e.key === 'Enter'){
      e.preventDefault();
      const code = listingCodeInput.value.trim();
      if(code) searchListingByCode(code);
    }
  });
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
    { name: 'Rumunia - Nabiał', type: 'xlsx', url: xlsxUrlNabial },
    { name: 'Rumunia - Napoje', type: 'json', url: jsonUrlNapoje },
    { name: 'Rumunia - Przyprawy', type: 'json', url: jsonUrlPrzyprawyProszek },
    { name: 'Rumunia - Puszki i słoiki', type: 'json', url: jsonUrlPuszkiSloiki },
    { name: 'Rumunia - Produkty podstawowe', type: 'json', url: jsonUrlProduktyPodstawowe },
    { name: 'Rumunia - Kawy i herbaty', type: 'json', url: jsonUrlKawyHerbaty },
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
        matches.push({
          Źródło: ds.name,
          Indeks: indexKey ? row[indexKey] ?? '' : '',
          Nazwa: nameKey ? row[nameKey] ?? '' : '',
          Producent: producerKey ? row[producerKey] ?? '' : '',
          Grupa: groupKey ? row[groupKey] ?? '' : '',
          Ranking: rankingKey ? row[rankingKey] ?? '' : '',
          'Kod EAN': eanKey ? row[eanKey] ?? '' : ''
        });
      }
    });
  });
  listingResults = matches;
  if(fromScan){
    maybeAddListingResult(searchCode, matches);
  }else{
    renderListingTable();
  }
}

function maybeAddListingResult(code, matches){
  if(!matches.length){
    return;
  }
  if(listingScannedCodes.has(code)){
    pendingListingAdd = { code, matches };
    if(listingConfirmModal) listingConfirmModal.classList.remove('hidden');
    return;
  }
  applyListingAdd(code, matches, false);
}

function applyListingAdd(code, matches, force){
  if(!force && listingScannedCodes.has(code)) return;
  listingScannedCodes.add(code);
  listingResults = listingResults.concat(matches);
  renderListingTable();
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
  const cols = Object.keys(listingResults[0]);
  listingHead.innerHTML = cols.map(c => `<th>${escapeHtml(c)}</th>`).join('');
  listingBody.innerHTML = listingResults.map(r => `
    <tr>
      ${cols.map(c => `<td>${escapeHtml(r[c] ?? '')}</td>`).join('')}
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

function resetFilters(){
  filters = { producer:'', group:'', name:'', ean:'', index:'', limit:'' };
}

function exportCSV(){
  const cols = Object.keys(fullData[0]);
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
  const cols = Object.keys(fullData[0]);
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
    const blob = await window.createCatalogPdf(products);
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
  const cols = Object.keys(fullData[0] || {});
  const indexKey = findColumn(cols, ['indeks', 'index', 'id']);
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
  const cols = Object.keys(fullData[0] || {});
  const indexKey = findColumn(cols, ['indeks', 'index', 'id']);
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
    const options = { coverDataUrl: catalogCoverDataUrl, priceMap: catalogPriceMap, currency, priceColor };
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

function buildPriceMap(rows){
  if(!rows || !rows.length) return null;
  const header = rows[0].map(h => String(h).toLowerCase().trim());
  const idxCol = header.indexOf('indeks');
  const priceCol = header.indexOf('cena');
  const unitCol = header.findIndex(h => h.includes('jednostka'));
  if(idxCol === -1 || priceCol === -1) return null;
  const map = new Map();
  for(let i=1;i<rows.length;i++){
    const r = rows[i];
    const index = String(r[idxCol] || '').trim();
    if(!index) continue;
    const price = String(r[priceCol] || '').trim();
    const unit = unitCol !== -1 ? String(r[unitCol] || '').trim() : '';
    map.set(index, { price, unit });
  }
  return map;
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
  const cols = Object.keys(fullData[0]);
  const producerKey = findColumn(cols, ['producent', 'producer']);
  const groupKey = findColumn(cols, ['grupa', 'group']);
  const nameKey = findColumn(cols, ['nazwa', 'name']);
  const eanKey = findColumn(cols, ['kod ean', 'ean']);
  const indexKey = findColumn(cols, ['indeks', 'index', 'id']);

  return fullData.filter(r => {
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
}

function updateTable(){
  const cols = Object.keys(fullData[0]);
  const filteredData = getFilteredData();
  const limitValue = getLimitValue();
  const data = expanded ? filteredData : filteredData.slice(0, limitValue);
  const indexKey = findColumn(cols, ['indeks', 'index', 'id']);
  const eanKey = findColumn(cols, ['kod ean', 'ean']);

  const tbody = document.getElementById('data-tbody');
  if(tbody){
    tbody.innerHTML = data.map(r => `
      <tr>
        ${cols.map(c => renderCell(c, r[c], indexKey, eanKey)).join('')}
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
  if(!adminPanelBtn) return;
  const isAdminEmail = String(email || '').toLowerCase() === 'admin@admin.com';
  const isAdminSession = localStorage.getItem('is_admin') === '1';
  const isAdminPassword = lastLoginPassword === 'ADMIN1';
  const isAdmin = isAdminEmail && (isAdminPassword || isAdminSession);
  adminPanelBtn.classList.toggle('hidden', !isAdmin);
  if(isAdmin){
    localStorage.setItem('is_admin', '1');
  }else{
    localStorage.removeItem('is_admin');
  }
}

function toggleProductSelection(checkbox){
  const index = checkbox.getAttribute('data-index');
  if(!index) return;
  const cols = Object.keys(fullData[0] || {});
  const indexKey = findColumn(cols, ['indeks', 'index', 'id']);
  if(!indexKey) return;
  const row = fullData.find(r => String(r[indexKey]).trim() === index);
  if(checkbox.checked && row){
    selectedProducts.set(index, row);
  }else{
    selectedProducts.delete(index);
  }
}

function renderCell(col, value, indexKey, eanKey){
  if(indexKey && col === indexKey){
    const idx = String(value ?? '').trim();
    const checked = selectedProducts.has(idx) ? 'checked' : '';
    const img = idx
      ? `<span class="img-hover" onmouseenter="positionPopup(this)">
           <img class="index-img" src="${buildImageUrl(idx, 'png')}" data-index="${escapeAttr(idx)}" data-tried="png" onerror="imageFallback(this)" alt="">
           <span class="img-pop">
             <img class="index-img-large" src="${buildImageUrl(idx, 'png')}" alt="">
           </span>
         </span>`
      : '';
    return `<td class="index-cell">
              <label class="select-box">
                <input type="checkbox" data-index="${escapeAttr(idx)}" onchange="toggleProductSelection(this)" ${checked}>
                <span></span>
              </label>
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

function buildImageUrl(index, ext){
  return `${currentImageBaseUrl}${encodeURIComponent(index)}.${ext}?alt=media`;
}

function imageFallback(img){
  const index = img.getAttribute('data-index');
  const tried = img.getAttribute('data-tried');
  const pop = img.closest('.img-hover')?.querySelector('.index-img-large');
  if(tried === 'png'){
    img.setAttribute('data-tried', 'jpg');
    img.src = buildImageUrl(index, 'jpg');
    if(pop) pop.src = buildImageUrl(index, 'jpg');
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
