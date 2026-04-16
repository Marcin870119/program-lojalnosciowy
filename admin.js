const firebaseConfig = {
  apiKey: 'AIzaSyDwzVRS5W2lklGMLcZJn-YPCK9OtBQZ7bI',
  authDomain: 'pdf-creator-f7a8b.firebaseapp.com',
  projectId: 'pdf-creator-f7a8b',
  appId: '1:606744201676:web:6f8c1b2c323fbaf6f3b569'
};

const authSplash = document.getElementById('auth-splash');
const authSplashMessage = document.getElementById('auth-splash-message');
const emailInput = document.getElementById('admin-user-email');
const passwordInput = document.getElementById('admin-user-password');
const createBtn = document.getElementById('admin-create-user');
const msg = document.getElementById('admin-create-msg');
const backBtn = document.getElementById('admin-back-btn');
const userList = document.getElementById('admin-user-list');
const userPicker = document.getElementById('admin-user-picker');
const userCount = document.getElementById('admin-user-count');
const salesReportFileInput = document.getElementById('admin-sales-report-file');
const salesReportUploadBtn = document.getElementById('admin-sales-report-upload');
const salesReportMsg = document.getElementById('admin-sales-report-msg');
const salesReportResults = document.getElementById('admin-sales-report-results');
const salesIndexFileInput = document.getElementById('admin-sales-index-file');
const salesIndexUploadBtn = document.getElementById('admin-sales-index-upload');
const salesIndexMsg = document.getElementById('admin-sales-index-msg');
const salesIndexResults = document.getElementById('admin-sales-index-results');
const USERS_COLLECTION = 'app_users';
const SALES_REPORTS_ROOT = 'raporty - maspo';
const STORAGE_BUCKET = 'gs://pdf-creator-f7a8b.firebasestorage.app';
const SALES_REP_HEADER_KEYS = ['opiekun_klienta'];
const ADMIN_EMAILS = ['admin@admin.info', 'admin@admin.com'];
const PENDING_USER_SYNC_STORAGE_KEY = 'admin_pending_app_users_sync_v1';
const SALES_REPORT_EMPTY_MESSAGE = 'Wybierz plik i uruchom podział na przedstawicieli.';
const SALES_INDEX_EMPTY_MESSAGE = 'Wybierz plik per indeks i uruchom podział na przedstawicieli.';
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
let salesReportRepOptions = [];
let latestUserDocs = [];
let latestRenderedUsers = [];
let pendingUserSyncInFlight = false;
let pendingUserSyncTimer = null;
let selectedAdminUserUid = '';
let redirectToIndexScheduled = false;

if(backBtn){
  backBtn.addEventListener('click', () => {
    window.location.href = 'index.html';
  });
}

function setAuthBootState(isLoading, message = 'Sprawdzam uprawnienia...'){
  if(authSplashMessage){
    authSplashMessage.textContent = message || 'Sprawdzam uprawnienia...';
  }
  if(document.body){
    document.body.classList.toggle('auth-booting', Boolean(isLoading));
  }
  if(authSplash){
    authSplash.classList.toggle('hidden', !isLoading);
  }
}

function redirectToIndex(){
  if(redirectToIndexScheduled) return;
  redirectToIndexScheduled = true;
  window.location.replace('index.html');
}

function escapeHtml(value){
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function isAdminEmail(email){
  return ADMIN_EMAILS.includes(String(email || '').toLowerCase());
}

function parseDateLike(value){
  if(!value) return null;
  if(value instanceof Date){
    return Number.isNaN(value.getTime()) ? null : value;
  }
  const parsedDate = new Date(value);
  return Number.isNaN(parsedDate.getTime()) ? null : parsedDate;
}

function formatUserDate(timestamp){
  try{
    if(timestamp?.toDate){
      return timestamp.toDate().toLocaleString('pl-PL');
    }
    const parsedDate = parseDateLike(timestamp);
    return parsedDate ? parsedDate.toLocaleString('pl-PL') : 'brak daty';
  }catch(e){
    return 'brak daty';
  }
}

function renderUserListMessage(message){
  if(!userList) return;
  userList.innerHTML = `<div class="admin-user-empty">${escapeHtml(message)}</div>`;
}

function setInlineMessage(node, message, type = 'info'){
  if(!node) return;
  node.textContent = message || '';
  node.classList.toggle('is-error', type === 'error');
  node.classList.toggle('is-success', type === 'success');
}

function setSalesReportMessage(message, type = 'info'){
  setInlineMessage(salesReportMsg, message, type);
}

function setSalesIndexMessage(message, type = 'info'){
  setInlineMessage(salesIndexMsg, message, type);
}

function renderImportResults(container, items, summary, emptyMessage){
  if(!container) return;

  if(!items?.length){
    container.innerHTML = `<div class="admin-user-empty">${escapeHtml(emptyMessage)}</div>`;
    return;
  }

  const summaryLine = summary
    ? `<div class="admin-user-empty">Utworzono ${escapeHtml(summary.filesCount)} plików dla ${escapeHtml(summary.repsCount)} przedstawicieli. Przetworzono ${escapeHtml(summary.rowsCount)} wierszy.</div>`
    : '';

  container.innerHTML = `${summaryLine}${items.map(item => `
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

function renderSalesReportResults(items, summary){
  renderImportResults(salesReportResults, items, summary, SALES_REPORT_EMPTY_MESSAGE);
}

function renderSalesIndexResults(items, summary){
  renderImportResults(salesIndexResults, items, summary, SALES_INDEX_EMPTY_MESSAGE);
}

function normalizePendingUserSyncEntry(entry){
  const uid = String(entry?.uid || '').trim();
  if(!uid) return null;

  return {
    uid,
    email: String(entry?.email || '').trim(),
    source: String(entry?.source || 'admin-panel').trim() || 'admin-panel',
    createdAt: parseDateLike(entry?.createdAt)?.toISOString() || '',
    blocked: Boolean(entry?.blocked),
    syncBlocked: Boolean(entry?.syncBlocked),
    salesReportFolder: String(entry?.salesReportFolder || '').trim(),
    salesRepName: String(entry?.salesRepName || '').trim(),
    syncSalesReport: Boolean(entry?.syncSalesReport),
    queuedAt: parseDateLike(entry?.queuedAt)?.toISOString() || new Date().toISOString(),
    lastSyncAttemptAt: parseDateLike(entry?.lastSyncAttemptAt)?.toISOString() || '',
    lastSyncError: String(entry?.lastSyncError || '').trim()
  };
}

function readPendingUserSyncQueue(){
  try{
    const raw = localStorage.getItem(PENDING_USER_SYNC_STORAGE_KEY);
    const parsed = JSON.parse(raw || '[]');
    if(!Array.isArray(parsed)) return [];
    return parsed
      .map(normalizePendingUserSyncEntry)
      .filter(Boolean);
  }catch(error){
    console.error('Pending user sync queue read error', error);
    return [];
  }
}

function writePendingUserSyncQueue(entries){
  const normalizedEntries = (entries || [])
    .map(normalizePendingUserSyncEntry)
    .filter(entry => entry && (entry.syncBlocked || entry.syncSalesReport));

  if(!normalizedEntries.length){
    localStorage.removeItem(PENDING_USER_SYNC_STORAGE_KEY);
    return;
  }

  localStorage.setItem(PENDING_USER_SYNC_STORAGE_KEY, JSON.stringify(normalizedEntries));
}

function upsertPendingUserSyncEntry(entry){
  const normalizedEntry = normalizePendingUserSyncEntry(entry);
  if(!normalizedEntry) return;

  const hasSyncBlocked = Object.prototype.hasOwnProperty.call(entry || {}, 'syncBlocked');
  const hasSyncSalesReport = Object.prototype.hasOwnProperty.call(entry || {}, 'syncSalesReport');
  const hasBlocked = Object.prototype.hasOwnProperty.call(entry || {}, 'blocked');
  const hasSalesReportFolder = Object.prototype.hasOwnProperty.call(entry || {}, 'salesReportFolder');
  const hasSalesRepName = Object.prototype.hasOwnProperty.call(entry || {}, 'salesRepName');
  const hasLastSyncError = Object.prototype.hasOwnProperty.call(entry || {}, 'lastSyncError');

  const queue = readPendingUserSyncQueue();
  const queueMap = new Map(queue.map(item => [item.uid, item]));
  const current = queueMap.get(normalizedEntry.uid) || {};

  queueMap.set(normalizedEntry.uid, normalizePendingUserSyncEntry({
    ...current,
    ...normalizedEntry,
    uid: normalizedEntry.uid,
    email: normalizedEntry.email || current.email || '',
    source: normalizedEntry.source || current.source || 'admin-panel',
    queuedAt: current.queuedAt || normalizedEntry.queuedAt || new Date().toISOString(),
    createdAt: normalizedEntry.createdAt || current.createdAt || '',
    blocked: hasBlocked ? normalizedEntry.blocked : Boolean(current.blocked),
    syncBlocked: hasSyncBlocked ? normalizedEntry.syncBlocked : Boolean(current.syncBlocked),
    salesReportFolder: hasSalesReportFolder ? normalizedEntry.salesReportFolder : String(current.salesReportFolder || ''),
    salesRepName: hasSalesRepName ? normalizedEntry.salesRepName : String(current.salesRepName || ''),
    syncSalesReport: hasSyncSalesReport ? normalizedEntry.syncSalesReport : Boolean(current.syncSalesReport),
    lastSyncError: hasLastSyncError ? normalizedEntry.lastSyncError : (current.lastSyncError || '')
  }));

  writePendingUserSyncQueue(Array.from(queueMap.values()));
}

function buildPendingUserSyncPayload(entry){
  const payload = {
    uid: entry.uid,
    email: entry.email || '',
    source: entry.source || 'admin-panel',
    updatedAt: firebase.firestore.FieldValue.serverTimestamp()
  };

  const createdAtDate = parseDateLike(entry.createdAt);
  if(createdAtDate){
    payload.createdAt = createdAtDate;
  }

  if(entry.syncBlocked){
    payload.blocked = Boolean(entry.blocked);
  }

  if(entry.syncSalesReport){
    const normalizedFolder = String(entry.salesReportFolder || '').trim();
    payload.salesReportFolder = normalizedFolder;
    payload.salesRepName = normalizedFolder
      ? (String(entry.salesRepName || '').trim() || normalizedFolder)
      : '';
  }

  return payload;
}

function schedulePendingUserSync(delay = 0){
  if(pendingUserSyncTimer){
    clearTimeout(pendingUserSyncTimer);
  }

  pendingUserSyncTimer = setTimeout(() => {
    pendingUserSyncTimer = null;
    void flushPendingUserSyncQueue();
  }, Math.max(0, delay));
}

async function flushPendingUserSyncQueue(){
  if(pendingUserSyncInFlight || !firestoreDb || typeof firebase?.firestore !== 'function'){
    return;
  }

  const queue = readPendingUserSyncQueue();
  if(!queue.length){
    return;
  }

  pendingUserSyncInFlight = true;

  try{
    const remainingEntries = [];

    for(const entry of queue){
      try{
        await firestoreDb
          .collection(USERS_COLLECTION)
          .doc(entry.uid)
          .set(buildPendingUserSyncPayload(entry), { merge: true });
      }catch(error){
        console.error('Pending user sync write error', error);
        remainingEntries.push(normalizePendingUserSyncEntry({
          ...entry,
          lastSyncAttemptAt: new Date().toISOString(),
          lastSyncError: error?.code || error?.message || 'unknown'
        }));
      }
    }

    writePendingUserSyncQueue(remainingEntries);
  }finally{
    pendingUserSyncInFlight = false;
    renderUsers(latestUserDocs);
  }
}

function normalizeUserSalesConfig(data){
  if(!data) return { reportFolder: '', reportName: '' };
  const reportFolder = String(data.salesReportFolder || data.storageFolder || '').trim();
  const reportName = String(data.salesRepName || data.displayName || '').trim() || reportFolder;
  return { reportFolder, reportName };
}

function mergeSalesReportRepOptions(options){
  const map = new Map();
  (options || []).forEach(item => {
    const folder = String(item?.folder || '').trim();
    if(!folder) return;
    const name = String(item?.name || '').trim() || folder;
    if(!map.has(folder)){
      map.set(folder, { folder, name });
      return;
    }
    const current = map.get(folder);
    if(!current.name && name){
      current.name = name;
    }
  });
  return Array.from(map.values()).sort((a, b) => a.name.localeCompare(b.name, 'pl'));
}

function buildSalesReportRepSelectOptions(selectedFolder, selectedName){
  const fallbackOption = selectedFolder
    ? [{ folder: selectedFolder, name: selectedName || selectedFolder }]
    : [];
  const options = mergeSalesReportRepOptions([
    ...salesReportRepOptions,
    ...fallbackOption
  ]);

  return `
    <option value="">Brak przypisania</option>
    ${options.map(option => `
      <option
        value="${escapeHtml(option.folder)}"
        data-name="${escapeHtml(option.name)}"
        ${selectedFolder === option.folder ? 'selected' : ''}
      >${escapeHtml(option.name)}</option>
    `).join('')}
  `;
}

function extractFolderFromStoragePath(path){
  const normalizedPath = String(path || '').trim();
  const prefix = `${SALES_REPORTS_ROOT}/`;
  if(!normalizedPath.startsWith(prefix)) return '';
  return normalizedPath.slice(prefix.length).split('/')[0] || '';
}

function collectAssignedReportOptionsFromDocs(docs){
  return (docs || []).map(doc => {
    const data = doc?.data?.() || {};
    const config = normalizeUserSalesConfig(data);
    return {
      folder: config.reportFolder,
      name: config.reportName
    };
  }).filter(item => item.folder);
}

async function refreshSalesReportRepOptionsFromStorage(){
  if(!storageService) return;
  try{
    const rootRef = storageService.ref().child(SALES_REPORTS_ROOT);
    const listResult = await rootRef.listAll();
    const fromStorage = (listResult.prefixes || []).map(prefix => {
      const folder = String(prefix?.name || '').trim();
      return folder ? { folder, name: folder } : null;
    }).filter(Boolean);

    salesReportRepOptions = mergeSalesReportRepOptions([
      ...salesReportRepOptions,
      ...fromStorage,
      ...collectAssignedReportOptionsFromDocs(latestUserDocs)
    ]);

    if(latestUserDocs.length){
      renderUsers(latestUserDocs);
    }
  }catch(error){
    console.error('Sales report options load error', error);
  }
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

async function uploadReportGroups(file, options = {}){
  if(typeof XLSX === 'undefined'){
    throw new Error('Biblioteka Excel nie została załadowana.');
  }
  if(!storageService){
    throw new Error('Firebase Storage nie jest gotowy.');
  }

  const setProgressMessage = typeof options.setMessage === 'function'
    ? options.setMessage
    : () => {};
  const buildMetadata = typeof options.buildMetadata === 'function'
    ? options.buildMetadata
    : () => ({});
  const progressLabel = String(options.progressLabel || 'Wysyłam').trim();
  const reportType = String(options.reportType || 'generic-report').trim() || 'generic-report';

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
    const extraMetadata = buildMetadata({
      headerRows,
      sheetRows,
      group,
      storagePath,
      outputFileName
    }) || {};
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

    setProgressMessage(`${progressLabel} ${index + 1} z ${groups.length}: ${group.repName}`);

    await storageService.ref().child(storagePath).put(outputBlob, {
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      customMetadata: {
        salesRep: group.repName,
        sourceFile: file.name,
        reportType,
        ...extraMetadata
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

async function uploadSalesReportGroups(file){
  return uploadReportGroups(file, {
    reportType: 'weekly-sales',
    progressLabel: 'Wysyłam',
    setMessage: setSalesReportMessage,
    buildMetadata: ({ sheetRows }) => buildSalesInsightMetadataPayload(buildWeeklySalesInsightFromRows(sheetRows))
  });
}

async function uploadSalesIndexGroups(file){
  return uploadReportGroups(file, {
    reportType: 'sales-by-index',
    progressLabel: 'Wysyłam raport per indeks',
    setMessage: setSalesIndexMessage
  });
}

function updateUserItemStatus(toggle, blocked){
  const userItem = toggle?.closest('.admin-user-item');
  const statusNode = userItem?.querySelector('.admin-user-status');
  if(!statusNode) return;

  statusNode.textContent = blocked ? 'Konto zablokowane' : 'Konto aktywne';
  statusNode.classList.toggle('is-blocked', blocked);
}

function createPinnedUserItem(){
  return {
    ...PINNED_TEST_USER,
    createdAt: null,
    blocked: false,
    reportFolder: '',
    reportName: '',
    syncPending: false,
    syncError: ''
  };
}

function formatAdminUserCountLabel(count){
  if(count === 1) return '1 konto';
  const lastTwoDigits = count % 100;
  const lastDigit = count % 10;
  if(lastDigit >= 2 && lastDigit <= 4 && (lastTwoDigits < 12 || lastTwoDigits > 14)){
    return `${count} konta`;
  }
  return `${count} kont`;
}

function buildUserPickerLabel(item){
  const email = String(item?.email || '').trim();
  return email || `UID: ${String(item?.uid || 'brak')}`;
}

function renderAdminUserPicker(items){
  if(userPicker){
    userPicker.innerHTML = items.map(item => `
      <option value="${escapeHtml(item.uid || '')}">
        ${escapeHtml(buildUserPickerLabel(item))}
      </option>
    `).join('');
    userPicker.value = selectedAdminUserUid;
  }

  if(userCount){
    userCount.textContent = formatAdminUserCountLabel(items.length);
  }
}

function renderSelectedAdminUserCard(){
  if(!userList) return;
  if(!latestRenderedUsers.length){
    userList.innerHTML = '<div class="admin-user-empty">Brak użytkowników do wyświetlenia.</div>';
    if(userPicker){
      userPicker.innerHTML = '<option value="">Brak użytkowników</option>';
      userPicker.value = '';
    }
    if(userCount){
      userCount.textContent = '';
    }
    return;
  }

  const selectedUser = latestRenderedUsers.find(item => item.uid === selectedAdminUserUid)
    || latestRenderedUsers[0];
  selectedAdminUserUid = selectedUser.uid;
  renderAdminUserPicker(latestRenderedUsers);
  userList.innerHTML = buildUserItem(selectedUser);
}

function buildUserItem({
  uid,
  email,
  createdAt,
  source,
  blocked,
  reportFolder,
  reportName,
  syncPending,
  syncError
}){
  const createdLabel = formatUserDate(createdAt);
  const sourceLabel = syncPending
    ? 'Oczekuje na synchronizację z bazą'
    : source === 'test-seed'
      ? 'Testowy wpis panelu'
      : source === 'auth-login'
        ? 'Konto dopisane po pierwszym logowaniu'
        : 'Konto zapisane w panelu';
  const statusLabel = blocked ? 'Konto zablokowane' : 'Konto aktywne';
  const switchDisabled = userBlockControlsEnabled ? '' : 'disabled';
  const switchChecked = blocked ? 'checked' : '';
  const statusClass = blocked ? ' is-blocked' : '';
  const assignmentLabel = reportFolder
    ? (reportName && reportName !== reportFolder ? `${reportName} (${reportFolder})` : (reportName || reportFolder))
    : 'Brak przypisania';
  const assignDisabled = userBlockControlsEnabled ? '' : 'disabled';
  const reportOptions = buildSalesReportRepSelectOptions(reportFolder, reportName);
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
      <div class="admin-user-meta">Raport PH: ${escapeHtml(assignmentLabel)}</div>
      ${syncError ? `<div class="admin-user-meta">Ostatni błąd synchronizacji: ${escapeHtml(syncError)}</div>` : ''}
      <div class="admin-user-assignment">
        <select
          class="admin-user-report-select"
          data-uid="${escapeHtml(uid || '')}"
          data-email="${escapeHtml(email || '')}"
          ${assignDisabled}
        >
          ${reportOptions}
        </select>
        <button
          class="btn-outline admin-user-assign-btn"
          type="button"
          data-uid="${escapeHtml(uid || '')}"
          data-email="${escapeHtml(email || '')}"
          ${assignDisabled}
        >Zapisz raport</button>
      </div>
    </div>
  `;
}

function renderFallbackUsers(){
  latestRenderedUsers = [createPinnedUserItem()];
  selectedAdminUserUid = latestRenderedUsers[0].uid;
  renderSelectedAdminUserCard();
}

function getTimestampValue(timestamp){
  try{
    if(timestamp?.toDate){
      return timestamp.toDate().getTime();
    }
    const parsedDate = parseDateLike(timestamp);
    return parsedDate ? parsedDate.getTime() : 0;
  }catch(error){
    return 0;
  }
}

function renderUsers(docs){
  if(!userList) return;
  latestUserDocs = Array.isArray(docs) ? [...docs] : [];
  const items = (docs || []).map(doc => {
    const data = doc.data() || {};
    const config = normalizeUserSalesConfig(data);
    return {
      uid: doc.id,
      email: data.email,
      createdAt: data.createdAt,
      source: data.source,
      blocked: Boolean(data.blocked),
      reportFolder: config.reportFolder,
      reportName: config.reportName,
      syncPending: false,
      syncError: ''
    };
  });

  const itemMap = new Map(items.map(item => [item.uid, item]));
  const pendingQueue = readPendingUserSyncQueue();

  pendingQueue.forEach(entry => {
    const existingItem = itemMap.get(entry.uid);
    const pendingValues = {
      uid: entry.uid,
      email: entry.email || existingItem?.email || '',
      createdAt: existingItem?.createdAt || entry.createdAt || entry.queuedAt || null,
      source: existingItem?.source || entry.source || 'admin-panel',
      blocked: entry.syncBlocked ? Boolean(entry.blocked) : Boolean(existingItem?.blocked),
      reportFolder: entry.syncSalesReport ? entry.salesReportFolder : (existingItem?.reportFolder || ''),
      reportName: entry.syncSalesReport ? (entry.salesRepName || entry.salesReportFolder) : (existingItem?.reportName || ''),
      syncPending: true,
      syncError: entry.lastSyncError || ''
    };

    if(existingItem){
      Object.assign(existingItem, pendingValues);
      return;
    }

    itemMap.set(entry.uid, pendingValues);
    items.push(pendingValues);
  });

  if(!items.length){
    renderFallbackUsers();
    return;
  }

  if(!items.some(item => item.uid === PINNED_TEST_USER.uid)){
    items.unshift({
      ...PINNED_TEST_USER,
      createdAt: null,
      blocked: false,
      reportFolder: '',
      reportName: '',
      syncPending: false,
      syncError: ''
    });
  }

  salesReportRepOptions = mergeSalesReportRepOptions([
    ...salesReportRepOptions,
    ...items.map(item => ({ folder: item.reportFolder, name: item.reportName })).filter(item => item.folder)
  ]);

  items.sort((a, b) => getTimestampValue(b.createdAt) - getTimestampValue(a.createdAt));
  latestRenderedUsers = items;
  if(!items.some(item => item.uid === selectedAdminUserUid)){
    selectedAdminUserUid = items[0]?.uid || '';
  }
  renderSelectedAdminUserCard();
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
    if(!String(currentData.salesReportFolder || '').trim()){
      await ref.set({
        salesReportFolder: 'Bubula Robert',
        salesRepName: 'Robert Bubula',
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
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

async function updateUserSalesReportAssignment(uid, email, reportFolder, reportName){
  if(!firestoreDb || !uid){
    throw new Error('Brak dostepu do bazy danych.');
  }

  const normalizedFolder = String(reportFolder || '').trim();
  const normalizedName = String(reportName || '').trim() || normalizedFolder;

  await firestoreDb.collection(USERS_COLLECTION).doc(uid).set({
    uid,
    email: email || '',
    salesReportFolder: normalizedFolder,
    salesRepName: normalizedFolder ? normalizedName : '',
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
    upsertPendingUserSyncEntry({
      uid,
      email,
      syncBlocked: false,
      lastSyncError: ''
    });
    if(msg){
      msg.textContent = nextBlocked
        ? 'Konto zostało zablokowane.'
        : 'Konto zostało odblokowane.';
    }
  }catch(error){
    console.error('Block user error', error);
    upsertPendingUserSyncEntry({
      uid,
      email,
      blocked: nextBlocked,
      syncBlocked: true,
      lastSyncError: error?.code || error?.message || 'unknown'
    });
    renderUsers(latestUserDocs);
    schedulePendingUserSync(1200);
    if(msg){
      msg.textContent = 'Nie udało się teraz zmienić blokady. Synchronizacja zostanie ponowiona automatycznie.';
    }
    setTimeout(() => {
      if(msg && msg.textContent === 'Nie udało się teraz zmienić blokady. Synchronizacja zostanie ponowiona automatycznie.'){
        msg.textContent = previousMessage;
      }
    }, 2500);
  }finally{
    toggle.disabled = !userBlockControlsEnabled;
  }
}

async function handleUserAssignmentSave(button){
  if(!button) return;
  const { uid, email } = button.dataset;
  if(!uid){
    return;
  }

  const userItem = button.closest('.admin-user-item');
  const select = userItem?.querySelector('.admin-user-report-select');
  if(!select){
    return;
  }

  const selectedOption = select.options[select.selectedIndex];
  const selectedFolder = String(select.value || '').trim();
  const selectedName = String(selectedOption?.dataset?.name || selectedOption?.textContent || '').trim();
  const previousMessage = msg ? msg.textContent : '';

  button.disabled = true;
  select.disabled = true;

  try{
    await updateUserSalesReportAssignment(uid, email, selectedFolder, selectedName);
    upsertPendingUserSyncEntry({
      uid,
      email,
      syncSalesReport: false,
      lastSyncError: ''
    });
    if(msg){
      msg.textContent = selectedFolder
        ? 'Przypisanie raportu zapisane.'
        : 'Przypisanie raportu usunięte.';
    }
  }catch(error){
    console.error('User report assignment error', error);
    upsertPendingUserSyncEntry({
      uid,
      email,
      salesReportFolder: selectedFolder,
      salesRepName: selectedFolder ? selectedName : '',
      syncSalesReport: true,
      lastSyncError: error?.code || error?.message || 'unknown'
    });
    renderUsers(latestUserDocs);
    schedulePendingUserSync(1200);
    if(msg){
      msg.textContent = 'Nie udało się teraz zapisać przypisania. Synchronizacja zostanie ponowiona automatycznie.';
    }
    setTimeout(() => {
      if(msg && msg.textContent === 'Nie udało się teraz zapisać przypisania. Synchronizacja zostanie ponowiona automatycznie.'){
        msg.textContent = previousMessage;
      }
    }, 2500);
  }finally{
    button.disabled = !userBlockControlsEnabled;
    select.disabled = !userBlockControlsEnabled;
  }
}

function handleUserListClick(event){
  const assignButton = event.target.closest('.admin-user-assign-btn');
  if(!assignButton) return;
  void handleUserAssignmentSave(assignButton);
}

function handleUserPickerChange(event){
  const nextUid = String(event?.target?.value || '').trim();
  if(!nextUid) return;
  selectedAdminUserUid = nextUid;
  renderSelectedAdminUserCard();
}

if(userList){
  userList.addEventListener('change', handleUserBlockToggle);
  userList.addEventListener('click', handleUserListClick);
}

if(userPicker){
  userPicker.addEventListener('change', handleUserPickerChange);
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
      const uploadedOptions = result.items.map(item => ({
        folder: extractFolderFromStoragePath(item.path),
        name: item.repName
      })).filter(item => item.folder);
      salesReportRepOptions = mergeSalesReportRepOptions([
        ...salesReportRepOptions,
        ...uploadedOptions
      ]);
      if(latestUserDocs.length){
        renderUsers(latestUserDocs);
      }
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

if(salesIndexUploadBtn){
  salesIndexUploadBtn.addEventListener('click', async () => {
    const file = salesIndexFileInput?.files?.[0];

    setSalesIndexMessage('');

    if(!file){
      setSalesIndexMessage('Wybierz plik Excel raportu per indeks.', 'error');
      return;
    }

    salesIndexUploadBtn.disabled = true;
    renderSalesIndexResults([], null);

    try{
      const result = await uploadSalesIndexGroups(file);
      renderSalesIndexResults(result.items, result.summary);
      setSalesIndexMessage(`Import per indeks zakończony. Utworzono ${result.summary.filesCount} plików.`, 'success');
    }catch(error){
      console.error('Sales index import error', error);
      setSalesIndexMessage(error?.message || 'Nie udało się podzielić i wysłać raportu per indeks.', 'error');
      renderSalesIndexResults([], null);
    }finally{
      salesIndexUploadBtn.disabled = false;
    }
  });
}

setAuthBootState(true, 'Sprawdzam uprawnienia...');

if(typeof firebase !== 'undefined'){
  if(!firebase.apps.length){
    firebase.initializeApp(firebaseConfig);
  }
  if(!firebase.apps.some(a => a.name === 'adminApp')){
    firebase.initializeApp(firebaseConfig, 'adminApp');
  }
  const appAuth = firebase.auth();
  const adminAuth = firebase.app('adminApp').auth();
  const db = firebase.firestore();
  firestoreDb = db;
  if(typeof firebase.storage === 'function'){
    storageService = firebase.app().storage(STORAGE_BUCKET);
    void refreshSalesReportRepOptionsFromStorage();
  }
  ensurePinnedTestUser(db);
  window.addEventListener('online', () => {
    schedulePendingUserSync(400);
  });
  schedulePendingUserSync(250);

  appAuth.onAuthStateChanged(user => {
    if(user && isAdminEmail(user.email)){
      setAuthBootState(false);
      return;
    }
    setAuthBootState(true, 'Przekierowuję do logowania...');
    renderUserListMessage('Brak aktywnej sesji administratora. Trwa przekierowanie...');
    redirectToIndex();
  });

  unsubscribeUsers = db
    .collection(USERS_COLLECTION)
    .onSnapshot(snapshot => {
      renderUsers(snapshot.docs);
    }, error => {
      console.error('Users list error', error);
      renderUsers([]);
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
      let saveErrorMessage = '';
      if(createdUser){
        const pendingUserSeed = {
          uid: createdUser.uid,
          email,
          createdAt: createdUser.metadata?.creationTime || new Date().toISOString(),
          source: 'admin-panel',
          blocked: false,
          syncBlocked: true
        };
        try{
          await db.collection(USERS_COLLECTION).doc(createdUser.uid).set({
            email,
            uid: createdUser.uid,
            createdAt: firebase.firestore.FieldValue.serverTimestamp(),
            source: 'admin-panel',
            blocked: false,
            salesReportFolder: '',
            salesRepName: ''
          }, { merge: true });
          savedToList = true;
          writePendingUserSyncQueue(
            readPendingUserSyncQueue().filter(item => item.uid !== createdUser.uid)
          );
        }catch(saveError){
          console.error('Save user list error', saveError);
          saveErrorMessage = saveError?.code || saveError?.message || '';
          upsertPendingUserSyncEntry({
            ...pendingUserSeed,
            lastSyncError: saveErrorMessage
          });
          renderUsers(latestUserDocs);
          schedulePendingUserSync(1200);
        }
      }
      await adminAuth.signOut();
      msg.textContent = savedToList
        ? 'Konto utworzone.'
        : `Konto utworzone. Profil jest w kolejce automatycznej synchronizacji${saveErrorMessage ? ` (${saveErrorMessage})` : ''}.`;
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
  setAuthBootState(false);
  renderFallbackUsers();
}
