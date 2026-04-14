const firebaseConfig = {
  apiKey: 'AIzaSyDwzVRS5W2lklGMLcZJn-YPCK9OtBQZ7bI',
  authDomain: 'pdf-creator-f7a8b.firebaseapp.com',
  projectId: 'pdf-creator-f7a8b',
  appId: '1:606744201676:web:6f8c1b2c323fbaf6f3b569'
};

const emailInput = document.getElementById('admin-user-email');
const passwordInput = document.getElementById('admin-user-password');
const createBtn = document.getElementById('admin-create-user');
const msg = document.getElementById('admin-create-msg');
const backBtn = document.getElementById('admin-back-btn');
const userList = document.getElementById('admin-user-list');
const salesReportFileInput = document.getElementById('admin-sales-report-file');
const salesReportUploadBtn = document.getElementById('admin-sales-report-upload');
const salesReportMsg = document.getElementById('admin-sales-report-msg');
const salesReportResults = document.getElementById('admin-sales-report-results');
const USERS_COLLECTION = 'app_users';
const SALES_REPORTS_ROOT = 'raporty - maspo';
const STORAGE_BUCKET = 'gs://pdf-creator-f7a8b.firebasestorage.app';
const SALES_REP_HEADER_KEYS = ['opiekun_klienta'];
const PINNED_TEST_USER = {
  uid: '9MB4MF77sRayF6jjgvLCMpphBJM2',
  email: 'robert.bubula@maspo.pl',
  source: 'test-seed'
};

let unsubscribeUsers = null;
let pinnedUserEnsured = false;
let userBlockControlsEnabled = true;
let firestoreDb = null;
let storageService = null;

if(!localStorage.getItem('is_admin')){
  window.location.href = 'index.html';
}

if(backBtn){
  backBtn.addEventListener('click', () => {
    window.location.href = 'index.html';
  });
}

function escapeHtml(value){
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatUserDate(timestamp){
  if(!timestamp?.toDate) return 'brak daty';
  try{
    return timestamp.toDate().toLocaleString('pl-PL');
  }catch(e){
    return 'brak daty';
  }
}

function renderUserListMessage(message){
  if(!userList) return;
  userList.innerHTML = `<div class="admin-user-empty">${escapeHtml(message)}</div>`;
}

function setSalesReportMessage(message, type = 'info'){
  if(!salesReportMsg) return;
  salesReportMsg.textContent = message || '';
  salesReportMsg.classList.toggle('is-error', type === 'error');
  salesReportMsg.classList.toggle('is-success', type === 'success');
}

function renderSalesReportResults(items, summary){
  if(!salesReportResults) return;

  if(!items?.length){
    salesReportResults.innerHTML = '<div class="admin-user-empty">Wybierz plik i uruchom podział na przedstawicieli.</div>';
    return;
  }

  const summaryLine = summary
    ? `<div class="admin-user-empty">Utworzono ${escapeHtml(summary.filesCount)} plików dla ${escapeHtml(summary.repsCount)} przedstawicieli. Przetworzono ${escapeHtml(summary.rowsCount)} wierszy.</div>`
    : '';

  salesReportResults.innerHTML = `${summaryLine}${items.map(item => `
    <div class="admin-import-item">
      <div class="admin-import-item-title">
        <span>${escapeHtml(item.repName)}</span>
        <span class="admin-import-badge">${escapeHtml(item.rowsCount)} wierszy</span>
      </div>
      <div class="admin-import-meta">Plik: ${escapeHtml(item.fileName)}</div>
      <div class="admin-import-path">${escapeHtml(item.path)}</div>
    </div>
  `).join('')}`;
}

function normalizeLookupValue(value){
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim()
    .replace(/[\s-]+/g, '_');
}

function sanitizeStorageSegment(value){
  const cleaned = String(value ?? '')
    .replace(/[\\/#?%]+/g, ' ')
    .replace(/[:*"<>|]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return cleaned || 'bez_nazwy';
}

function stripFileExtension(filename){
  return String(filename || '').replace(/\.[^.]+$/, '') || 'raport_sprzedazy';
}

function truncateSheetName(name){
  return String(name || 'Export').slice(0, 31) || 'Export';
}

function findSalesRepColumn(rows){
  const maxRowsToCheck = Math.min(rows.length, 5);

  for(let rowIndex = 0; rowIndex < maxRowsToCheck; rowIndex += 1){
    const row = Array.isArray(rows[rowIndex]) ? rows[rowIndex] : [];
    for(let columnIndex = 0; columnIndex < row.length; columnIndex += 1){
      const cellKey = normalizeLookupValue(row[columnIndex]);
      if(SALES_REP_HEADER_KEYS.includes(cellKey)){
        return { rowIndex, columnIndex };
      }
    }
  }

  return null;
}

function copyWorksheetLayout(sourceSheet, targetSheet, rowCount){
  if(Array.isArray(sourceSheet?.['!cols'])){
    targetSheet['!cols'] = sourceSheet['!cols'].map(column => ({ ...column }));
  }

  if(Array.isArray(sourceSheet?.['!merges'])){
    targetSheet['!merges'] = sourceSheet['!merges']
      .filter(range => range?.s?.r < rowCount && range?.e?.r < rowCount)
      .map(range => ({
        s: { ...range.s },
        e: { ...range.e }
      }));
  }
}

function groupRowsBySalesRep(rows, headerInfo){
  const headerRows = rows.slice(0, headerInfo.rowIndex + 1);
  const dataRows = rows.slice(headerInfo.rowIndex + 1);
  const groupedRows = new Map();

  dataRows.forEach(row => {
    if(!Array.isArray(row)) return;
    const repName = String(row[headerInfo.columnIndex] ?? '').trim();
    if(!repName) return;

    const repKey = normalizeLookupValue(repName);
    if(!groupedRows.has(repKey)){
      groupedRows.set(repKey, {
        repName,
        rows: []
      });
    }

    groupedRows.get(repKey).rows.push(row);
  });

  return {
    headerRows,
    groups: Array.from(groupedRows.values()).sort((a, b) => a.repName.localeCompare(b.repName, 'pl'))
  };
}

function normalizeHeaderKey(value){
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[_\-.]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseNumericCell(value){
  if(typeof value === 'number'){
    return Number.isFinite(value) ? value : 0;
  }
  const normalized = String(value ?? '')
    .replace(/\s+/g, '')
    .replace(',', '.');
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function buildWeeklySalesInsightFromRows(rows){
  if(!Array.isArray(rows) || rows.length < 3){
    return null;
  }

  const firstRow = Array.isArray(rows[0]) ? rows[0] : [];
  const secondRow = Array.isArray(rows[1]) ? rows[1] : [];
  const firstCellKey = normalizeHeaderKey(firstRow[0]);
  const isWideWeeklyReport =
    firstCellKey.includes('tydzien sprzedazy') ||
    firstCellKey.includes('tydzien sprzedaży');

  if(isWideWeeklyReport){
    const salesColumns = secondRow
      .map((header, index) => ({ header: normalizeHeaderKey(header), index }))
      .filter(item => item.header.includes('sprzedaz waluta'))
      .map(item => item.index);

    if(salesColumns.length < 2){
      return null;
    }

    const weekTotals = salesColumns.map((columnIndex, weekPosition) => {
      const rawWeekLabel = String(firstRow[columnIndex] ?? '').trim();
      const weekLabel = rawWeekLabel && !rawWeekLabel.startsWith('=')
        ? rawWeekLabel
        : `Tydzień ${weekPosition + 1}`;
      const total = rows
        .slice(2)
        .reduce((sum, row) => sum + parseNumericCell(Array.isArray(row) ? row[columnIndex] : ''), 0);
      return {
        weekLabel,
        total
      };
    });

    const previous = weekTotals[weekTotals.length - 2];
    const last = weekTotals[weekTotals.length - 1];
    return {
      previousWeek: previous.weekLabel,
      lastWeek: last.weekLabel,
      previousTotal: previous.total,
      lastTotal: last.total
    };
  }

  const headers = firstRow.map(header => normalizeHeaderKey(header));
  const weekIndex = headers.findIndex(header => header.includes('tydzien') || header.includes('week'));
  const valueIndex = headers.findIndex(header =>
    header.includes('sprzedaz wartosciowa') ||
    header.includes('sprzedaz waluta') ||
    header === 'wartosc' ||
    header === 'value'
  );

  if(weekIndex < 0 || valueIndex < 0){
    return null;
  }

  const weekTotalsMap = new Map();
  rows.slice(1).forEach(row => {
    if(!Array.isArray(row)) return;
    const weekLabel = String(row[weekIndex] ?? '').trim();
    if(!weekLabel) return;
    const value = parseNumericCell(row[valueIndex]);
    weekTotalsMap.set(weekLabel, (weekTotalsMap.get(weekLabel) || 0) + value);
  });

  const weekTotals = Array.from(weekTotalsMap.entries()).map(([weekLabel, total]) => ({
    weekLabel,
    total
  }));

  if(weekTotals.length < 2){
    return null;
  }

  const previous = weekTotals[weekTotals.length - 2];
  const last = weekTotals[weekTotals.length - 1];
  return {
    previousWeek: previous.weekLabel,
    lastWeek: last.weekLabel,
    previousTotal: previous.total,
    lastTotal: last.total
  };
}

function buildSalesInsightMetadataPayload(insight){
  if(!insight){
    return {};
  }
  return {
    insightPreviousWeek: String(insight.previousWeek || ''),
    insightLastWeek: String(insight.lastWeek || ''),
    insightPreviousTotal: String(Number(insight.previousTotal || 0)),
    insightLastTotal: String(Number(insight.lastTotal || 0)),
    insightGeneratedAt: new Date().toISOString()
  };
}

async function uploadSalesReportGroups(file){
  if(typeof XLSX === 'undefined'){
    throw new Error('Biblioteka Excel nie została załadowana.');
  }
  if(!storageService){
    throw new Error('Firebase Storage nie jest gotowy.');
  }

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, {
    type: 'array',
    raw: true,
    cellDates: true
  });

  const sheetName = workbook.SheetNames?.[0];
  if(!sheetName){
    throw new Error('Plik Excel nie zawiera żadnego arkusza.');
  }

  const sourceSheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sourceSheet, {
    header: 1,
    raw: true,
    defval: ''
  });

  const headerInfo = findSalesRepColumn(rows);
  if(!headerInfo){
    throw new Error('Nie znaleziono kolumny OPIEKUN_KLIENTA.');
  }

  const { headerRows, groups } = groupRowsBySalesRep(rows, headerInfo);
  if(!groups.length){
    throw new Error('Nie znaleziono wierszy przypisanych do przedstawicieli handlowych.');
  }

  const sourceBaseName = sanitizeStorageSegment(stripFileExtension(file.name));
  const safeSheetName = truncateSheetName(sheetName);
  const uploadResults = [];
  let processedRows = 0;

  for(let index = 0; index < groups.length; index += 1){
    const group = groups[index];
    const repFolderName = sanitizeStorageSegment(group.repName);
    const outputFileName = `${sourceBaseName} - ${repFolderName}.xlsx`;
    const storagePath = `${SALES_REPORTS_ROOT}/${repFolderName}/${outputFileName}`;
    const sheetRows = [...headerRows, ...group.rows];
    const insightMetadata = buildSalesInsightMetadataPayload(buildWeeklySalesInsightFromRows(sheetRows));
    const outputSheet = XLSX.utils.aoa_to_sheet(sheetRows);

    copyWorksheetLayout(sourceSheet, outputSheet, sheetRows.length);

    const outputWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, safeSheetName);

    const outputBuffer = XLSX.write(outputWorkbook, {
      bookType: 'xlsx',
      type: 'array',
      compression: true
    });

    const outputBlob = new Blob(
      [outputBuffer],
      { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
    );

    setSalesReportMessage(`Wysyłam ${index + 1} z ${groups.length}: ${group.repName}`);

    await storageService.ref().child(storagePath).put(outputBlob, {
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      customMetadata: {
        salesRep: group.repName,
        sourceFile: file.name,
        ...insightMetadata
      }
    });

    processedRows += group.rows.length;
    uploadResults.push({
      repName: group.repName,
      rowsCount: group.rows.length,
      fileName: outputFileName,
      path: storagePath
    });
  }

  return {
    items: uploadResults,
    summary: {
      filesCount: uploadResults.length,
      repsCount: uploadResults.length,
      rowsCount: processedRows
    }
  };
}

function updateUserItemStatus(toggle, blocked){
  const userItem = toggle?.closest('.admin-user-item');
  const statusNode = userItem?.querySelector('.admin-user-status');
  if(!statusNode) return;

  statusNode.textContent = blocked ? 'Konto zablokowane' : 'Konto aktywne';
  statusNode.classList.toggle('is-blocked', blocked);
}

function buildUserItem({ uid, email, createdAt, source, blocked }){
  const createdLabel = formatUserDate(createdAt);
  const sourceLabel = source === 'test-seed' ? 'Testowy wpis panelu' : 'Konto zapisane w panelu';
  const statusLabel = blocked ? 'Konto zablokowane' : 'Konto aktywne';
  const switchDisabled = userBlockControlsEnabled ? '' : 'disabled';
  const switchChecked = blocked ? 'checked' : '';
  const statusClass = blocked ? ' is-blocked' : '';
  return `
    <div class="admin-user-item">
      <div class="admin-user-top">
        <div class="admin-user-email">${escapeHtml(email || 'Brak emaila')}</div>
        <label class="admin-user-switch">
          <span class="admin-user-switch-text">Blokada</span>
          <span class="admin-switch">
            <input
              class="admin-user-block-toggle"
              type="checkbox"
              data-uid="${escapeHtml(uid || '')}"
              data-email="${escapeHtml(email || '')}"
              ${switchChecked}
              ${switchDisabled}
            >
            <span class="admin-switch-slider" aria-hidden="true"></span>
          </span>
        </label>
      </div>
      <div class="admin-user-status${statusClass}">${escapeHtml(statusLabel)}</div>
      <div class="admin-user-meta">UID: ${escapeHtml(uid || 'Brak UID')}</div>
      <div class="admin-user-meta">Dodano: ${escapeHtml(createdLabel)}</div>
      <div class="admin-user-meta">${escapeHtml(sourceLabel)}</div>
    </div>
  `;
}

function renderFallbackUsers(){
  if(!userList) return;
  userList.innerHTML = buildUserItem({
    ...PINNED_TEST_USER,
    createdAt: null,
    blocked: false
  });
}

function getTimestampValue(timestamp){
  if(!timestamp?.toDate) return 0;
  try{
    return timestamp.toDate().getTime();
  }catch(error){
    return 0;
  }
}

function renderUsers(docs){
  if(!userList) return;
  if(!docs.length){
    renderFallbackUsers();
    return;
  }

  const items = docs.map(doc => {
    const data = doc.data() || {};
    return {
      uid: doc.id,
      email: data.email,
      createdAt: data.createdAt,
      source: data.source,
      blocked: Boolean(data.blocked)
    };
  });

  if(!items.some(item => item.uid === PINNED_TEST_USER.uid)){
    items.unshift({
      ...PINNED_TEST_USER,
      createdAt: null,
      blocked: false
    });
  }

  items.sort((a, b) => getTimestampValue(b.createdAt) - getTimestampValue(a.createdAt));
  userList.innerHTML = items.map(buildUserItem).join('');
}

async function ensurePinnedTestUser(db){
  if(pinnedUserEnsured) return;
  pinnedUserEnsured = true;
  try{
    const ref = db.collection(USERS_COLLECTION).doc(PINNED_TEST_USER.uid);
    const snap = await ref.get();
    const currentData = snap.exists ? (snap.data() || {}) : {};
    if(!snap.exists || !currentData.createdAt){
      await ref.set({
        email: PINNED_TEST_USER.email,
        uid: PINNED_TEST_USER.uid,
        createdAt: firebase.firestore.FieldValue.serverTimestamp(),
        source: PINNED_TEST_USER.source,
        blocked: false
      }, { merge: true });
    }
  }catch(error){
    console.error('Pinned user seed error', error);
  }
}

async function updateUserBlockedState(uid, email, blocked){
  if(!firestoreDb || !uid){
    throw new Error('Brak dostepu do bazy danych.');
  }

  await firestoreDb.collection(USERS_COLLECTION).doc(uid).set({
    uid,
    email: email || '',
    blocked: Boolean(blocked),
    updatedAt: firebase.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
}

async function handleUserBlockToggle(event){
  const toggle = event.target.closest('.admin-user-block-toggle');
  if(!toggle) return;

  const { uid, email } = toggle.dataset;
  const nextBlocked = toggle.checked;
  const previousMessage = msg ? msg.textContent : '';

  if(!uid){
    toggle.checked = !nextBlocked;
    return;
  }

  toggle.disabled = true;
  updateUserItemStatus(toggle, nextBlocked);

  try{
    await updateUserBlockedState(uid, email, nextBlocked);
    if(msg){
      msg.textContent = nextBlocked
        ? 'Konto zostało zablokowane.'
        : 'Konto zostało odblokowane.';
    }
  }catch(error){
    console.error('Block user error', error);
    toggle.checked = !nextBlocked;
    updateUserItemStatus(toggle, !nextBlocked);
    if(msg){
      msg.textContent = 'Nie udało się zmienić blokady konta.';
    }
    setTimeout(() => {
      if(msg && msg.textContent === 'Nie udało się zmienić blokady konta.'){
        msg.textContent = previousMessage;
      }
    }, 2500);
  }finally{
    toggle.disabled = !userBlockControlsEnabled;
  }
}

if(userList){
  userList.addEventListener('change', handleUserBlockToggle);
}

if(salesReportUploadBtn){
  salesReportUploadBtn.addEventListener('click', async () => {
    const file = salesReportFileInput?.files?.[0];

    setSalesReportMessage('');

    if(!file){
      setSalesReportMessage('Wybierz plik Excel do importu.', 'error');
      return;
    }

    salesReportUploadBtn.disabled = true;
    renderSalesReportResults([], null);

    try{
      const result = await uploadSalesReportGroups(file);
      renderSalesReportResults(result.items, result.summary);
      setSalesReportMessage(`Import zakończony. Utworzono ${result.summary.filesCount} plików.`, 'success');
    }catch(error){
      console.error('Sales report import error', error);
      setSalesReportMessage(error?.message || 'Nie udało się podzielić i wysłać raportu.', 'error');
      renderSalesReportResults([], null);
    }finally{
      salesReportUploadBtn.disabled = false;
    }
  });
}

if(typeof firebase !== 'undefined'){
  if(!firebase.apps.length){
    firebase.initializeApp(firebaseConfig);
  }
  if(!firebase.apps.some(a => a.name === 'adminApp')){
    firebase.initializeApp(firebaseConfig, 'adminApp');
  }
  const adminAuth = firebase.app('adminApp').auth();
  const db = firebase.firestore();
  firestoreDb = db;
  if(typeof firebase.storage === 'function'){
    storageService = firebase.app().storage(STORAGE_BUCKET);
  }
  ensurePinnedTestUser(db);

  unsubscribeUsers = db
    .collection(USERS_COLLECTION)
    .onSnapshot(snapshot => {
      renderUsers(snapshot.docs);
    }, error => {
      console.error('Users list error', error);
      renderFallbackUsers();
    });

  createBtn.addEventListener('click', async () => {
    const email = emailInput.value.trim();
    const pass = passwordInput.value;
    msg.textContent = '';
    if(!email || !pass){
      msg.textContent = 'Podaj email i hasło.';
      return;
    }
    if(pass.length < 6){
      msg.textContent = 'Hasło min. 6 znaków.';
      return;
    }
    createBtn.disabled = true;
    try{
      const credential = await adminAuth.createUserWithEmailAndPassword(email, pass);
      const createdUser = credential.user;
      let savedToList = false;
      if(createdUser){
        try{
          await db.collection(USERS_COLLECTION).doc(createdUser.uid).set({
            email,
            uid: createdUser.uid,
            createdAt: firebase.firestore.FieldValue.serverTimestamp(),
            source: 'admin-panel',
            blocked: false
          }, { merge: true });
          savedToList = true;
        }catch(saveError){
          console.error('Save user list error', saveError);
        }
      }
      await adminAuth.signOut();
      msg.textContent = savedToList
        ? 'Konto utworzone.'
        : 'Konto utworzone, ale nie zapisano go na liście.';
      emailInput.value = '';
      passwordInput.value = '';
    }catch(e){
      console.error('Create user error', e);
      msg.textContent = e.code || 'Nie udało się utworzyć konta.';
    }finally{
      createBtn.disabled = false;
    }
  });
}else{
  renderFallbackUsers();
}
