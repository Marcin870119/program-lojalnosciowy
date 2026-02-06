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

const slodyczeContainer = document.getElementById('slodycze-content');
const miesoContainer = document.getElementById('mieso-wedliny-content');
const nabialContainer = document.getElementById('nabial-content');
const napojeContainer = document.getElementById('napoje-content');
const przyprawyProszekContainer = document.getElementById('przyprawy-proszek-content');
const puszkiSloikiContainer = document.getElementById('puszki-sloiki-content');
const produktyPodstawoweContainer = document.getElementById('produkty-podstawowe-content');
const kawyHerbatyContainer = document.getElementById('kawy-herbaty-content');

let fullData = [];
let expanded = false;
let viewMode = 'table';
let activeContainer = slodyczeContainer;
let currentCategoryName = 'Slodycze';
let currentCategorySlug = 'slodycze';
let filters = {
  producer: '',
  group: '',
  name: '',
  ean: '',
  index: '',
  limit: ''
};
const LIMIT = 25;

// OBSŁUGA ZAKŁADEK
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    tab.classList.add('active');

    document.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));
    document.getElementById(tab.dataset.tab).classList.remove('hidden');

    slodyczeContainer.innerHTML = '';
    miesoContainer.innerHTML = '';
    nabialContainer.innerHTML = '';
    napojeContainer.innerHTML = '';
    przyprawyProszekContainer.innerHTML = '';
    puszkiSloikiContainer.innerHTML = '';
    produktyPodstawoweContainer.innerHTML = '';
    kawyHerbatyContainer.innerHTML = '';
    viewMode = 'table';
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
  render();
});

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
    <div class="actions">
      <button onclick="exportCSV()">Zapisz CSV</button>
      <button onclick="exportXLS()">Zapisz XLS</button>
      <button onclick="previewPDF()">Podgląd PDF</button>
      <button onclick="downloadPDF()">Zapisz PDF</button>
    </div>

    ${viewMode === 'pdf' ? `
      <div class="pdf-preview">
        <iframe src="${pdfPreviewUrl + encodeURIComponent(getActivePdfUrl())}" title="Podgląd PDF"></iframe>
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

      <table>
        <thead>
          <tr>${cols.map(c => `<th>${c}</th>`).join('')}</tr>
        </thead>
        <tbody id="data-tbody"></tbody>
      </table>
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

function setFilter(key, value, soft){
  filters[key] = value;
  expanded = false;
  if(soft){
    updateTable();
  }else{
    render();
  }
}

function resetFilters(){
  filters = { producer:'', group:'', name:'', ean:'', index:'', limit:'' };
}

function exportCSV(){
  const cols = Object.keys(fullData[0]);
  let csv = cols.join(',') + '\n';

  getFilteredData().forEach(r => {
    csv += cols
      .map(c => `"${(r[c] ?? '').toString().replace(/"/g, '""')}"`)
      .join(',') + '\n';
  });

  download(csv, `${currentCategorySlug}_rumunia.csv`, 'text/csv');
}

function exportXLS(){
  const cols = Object.keys(fullData[0]);
  let html = '<table><tr>' + cols.map(c => `<th>${c}</th>`).join('') + '</tr>';

  getFilteredData().forEach(r => {
    html += '<tr>' + cols.map(c => `<td>${r[c] ?? ''}</td>`).join('') + '</tr>';
  });

  html += '</table>';
  download(html, `${currentCategorySlug}_rumunia.xls`, 'application/vnd.ms-excel');
}

function previewPDF(){
  viewMode = 'pdf';
  render();
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

  const tbody = document.getElementById('data-tbody');
  if(tbody){
    tbody.innerHTML = data.map(r => `
      <tr>
        ${cols.map(c => `<td>${r[c] ?? ''}</td>`).join('')}
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
