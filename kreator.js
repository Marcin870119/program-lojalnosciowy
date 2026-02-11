// =======================================================
// script.js – FINALNA WERSJA
// PDF: bez ucinania + bez „skoku” layoutu + równe marginesy
// =======================================================

// FORMAT WALUTY
const fmt = new Intl.NumberFormat('en-GB', {
  style: 'currency',
  currency: 'GBP',
  maximumFractionDigits: 0
});

let clientsData = [];

// =======================================================
// DOM READY
// =======================================================
document.addEventListener('DOMContentLoaded', () => {
  const targetInput = document.getElementById('target-world');
  const salesInput = document.getElementById('sales-world');
  const clientSelect = document.getElementById('clientSelect');
  const clientNameEl = document.getElementById('selected-client-name');
  const summaryBox = document.getElementById('points-summary');
  const congratsEl = document.getElementById('congrats-text');
  const endOfPeriodEl = document.getElementById('end-of-period');
  const pointRows = summaryBox.querySelectorAll('.point-row');

  // ─── PRZYCISK WYCZYŚĆ ───
  const clearBtn = document.createElement('button');
  clearBtn.textContent = 'Wyczyść dane';
  clearBtn.id = 'clear-data';
  clearBtn.style.marginLeft = '20px';
  clearBtn.style.padding = '10px 20px';
  clearBtn.style.background = '#6c757d';
  clearBtn.style.color = '#fff';
  clearBtn.style.border = 'none';
  clearBtn.style.borderRadius = '8px';
  clearBtn.style.cursor = 'pointer';
  clearBtn.style.fontSize = '15px';

  const actionsDiv = document.querySelector('.actions');
  if (actionsDiv) actionsDiv.prepend(clearBtn);

  clearBtn.addEventListener('click', () => {
    clientNameEl.textContent = '';
    summaryBox.style.display = 'none';
    congratsEl.textContent = '';
    congratsEl.className = 'congrats';
    pointRows.forEach(r => r.style.display = 'none');
    targetInput.value = 0;
    salesInput.value = 0;
    document.getElementById('progress-fill').style.width = '0%';
    document.getElementById('progress-pct').textContent = '0%';
    document.getElementById('target-display').textContent = fmt.format(0);
    document.getElementById('sales-display').textContent = fmt.format(0);
    document.getElementById('remain-display').textContent = fmt.format(0);
    clientSelect.value = '';
  });

  targetInput.addEventListener('input', update);
  salesInput.addEventListener('input', update);

  document.getElementById('reset').addEventListener('click', () => {
    salesInput.value = 0;
    update();
  });

  document.getElementById('excelFile').addEventListener('change', handleExcelUpload);
  document.getElementById('pdf').addEventListener('click', generatePDF);

  clientSelect.addEventListener('change', () => {
    const idx = Number(clientSelect.value);
    if (!clientsData[idx]) {
      clientNameEl.textContent = '';
      summaryBox.style.display = 'none';
      congratsEl.textContent = '';
      congratsEl.className = 'congrats';
      return;
    }

    const row = clientsData[idx];
    clientNameEl.textContent = `${row[1] || ''} ${row[2] || ''}`.trim();

    const target = Math.round(Number(row[4]) || 0);
    const sales = Math.round(Number(row[5]) || 0);
    targetInput.value = target;
    salesInput.value = sales;

    const wykonany = (row[6] || '').toString().trim().toUpperCase() === 'TAK';

    if (!wykonany) {
      summaryBox.style.display = 'block';
      pointRows.forEach(r => r.style.display = 'none');
      congratsEl.textContent =
        endOfPeriodEl && endOfPeriodEl.checked
          ? 'Niestety, tym razem się nie udało. Popracuj ze swoim przedstawicielem w kolejnym etapie i razem osiągniesz cel.'
          : 'Jesteś już bardzo blisko! Jeszcze chwila i osiągniesz swój cel 💪';
      congratsEl.className =
        endOfPeriodEl && endOfPeriodEl.checked
          ? 'congrats end'
          : 'congrats progress';
      update();
      return;
    }

    document.getElementById('points-h').textContent = Math.round(Number(row[7]) || 0);
    document.getElementById('points-i').textContent = Math.round(Number(row[8]) || 0);
    document.getElementById('points-j').textContent = Math.round(Number(row[9]) || 0);
    document.getElementById('points-n').textContent = Math.round(Number(row[13]) || 0);

    pointRows.forEach(r => r.style.display = 'block');
    summaryBox.style.display = 'block';
    congratsEl.textContent = 'Brawo! Osiągnąłeś swój cel 🎉';
    congratsEl.className = 'congrats success';

    update();
  });

  update();
});

// =======================================================
// IMPORT EXCEL
// =======================================================
function handleExcelUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = evt => {
    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });

    clientsData = rows.slice(1).filter(r => r && r[1]);

    const select = document.getElementById('clientSelect');
    select.innerHTML = '<option value="">Wybierz klienta...</option>';

    clientsData.forEach((row, i) => {
      const opt = document.createElement('option');
      opt.value = i;
      opt.textContent = `${row[1] || ''} ${row[2] || ''}`.trim();
      select.appendChild(opt);
    });

    select.disabled = clientsData.length === 0;
  };
  reader.readAsArrayBuffer(file);
}

// =======================================================
// UPDATE
// =======================================================
function update() {
  const target = Math.round(Number(document.getElementById('target-world').value) || 0);
  const sales = Math.round(Number(document.getElementById('sales-world').value) || 0);
  const remain = Math.max(target - sales, 0);

  let pct = target > 0 ? (sales / target) * 100 : 0;
  pct = Math.round(Math.min(100, Math.max(0, pct)));

  document.getElementById('target-display').textContent = fmt.format(target);
  document.getElementById('sales-display').textContent = fmt.format(sales);
  document.getElementById('remain-display').textContent = fmt.format(remain);
  document.getElementById('progress-fill').style.width = pct + '%';
  document.getElementById('progress-pct').textContent = pct + '%';
}

// =======================================================
// PDF – wersja z przesunięciem mapy w prawo + mniejsza rozdzielczość (scale 1.4)
// =======================================================
function generatePDF() {
  const source = document.getElementById('capture');
  if (!source) return;

  const clone = source.cloneNode(true);

  // Przygotowanie klonu – staramy się zachować naturalny układ
  clone.style.position = 'absolute';
  clone.style.left = '-9999px';
  clone.style.top = '0';
  clone.style.width = 'auto';
  clone.style.minWidth = 'unset';
  clone.style.maxWidth = 'unset';
  clone.style.height = 'auto';
  clone.style.margin = '0';
  clone.style.padding = '0';
  clone.style.overflow = 'visible';
  clone.style.boxSizing = 'border-box';
  clone.style.backgroundColor = '#ffffff';
  document.body.appendChild(clone);

  // Usuwamy interaktywne elementy i niepotrzebne etykiety
  clone.querySelectorAll(
    'select, input, button, label[for="end-of-period"], #excelFile, .actions, #clear-data, #reset'
  ).forEach(el => el.remove());

  Array.from(clone.querySelectorAll('*')).forEach(el => {
    const txt = (el.textContent || '').trim();
    if (txt === 'Importuj Excel:' || txt === 'Koniec okresu programu') {
      el.remove();
    }
  });

  // Główny kontener – usuwamy ograniczenia szerokości
  const main = clone.querySelector('#main-container');
  if (main) {
    main.style.maxWidth = 'none';
    main.style.width = '100%';
    main.style.margin = '0';
    main.style.padding = '0';
    main.style.boxShadow = 'none';
  }

  // Karta – bez sztucznego skalowania
  const card = clone.querySelector('.world-food-card');
  if (card) {
    card.style.transform = 'none';
    card.style.width = 'auto';
    card.style.maxWidth = 'none';
    card.style.margin = '0 auto';
    card.style.padding = '15px'; // możesz zmienić na 0 jeśli chcesz mniej odstępu
  }

  // Mapa – wypełnia przestrzeń + przesunięcie w prawo
  const map = clone.querySelector('.europe-map');
  if (map) {
    map.style.width = '100%';
    map.style.maxWidth = 'none';
    map.style.height = 'auto';
    map.style.margin = '0';
    map.style.padding = '0';
    map.style.display = 'block';
    map.style.objectFit = 'contain';
    map.style.position = 'relative';
    map.style.left = '80px';           // przesunięcie mapy w prawo
  }

  // Kontener mapy + flag – minimalne boczne marginesy
  const mapCont = clone.querySelector('.map-container');
  if (mapCont) {
    mapCont.style.width = '100%';
    mapCont.style.maxWidth = 'none';
    mapCont.style.margin = '0';
    mapCont.style.padding = '30px 10px 50px 10px';
    mapCont.style.position = 'relative';
    mapCont.style.overflow = 'visible';
  }

  const flagsColumn = clone.querySelector('.flags-column');
  if (flagsColumn) {
    flagsColumn.style.marginLeft = '0';
    flagsColumn.style.paddingLeft = '8px';
    flagsColumn.style.left = '0';
  }

  // Czekamy aż przeglądarka przeliczy rzeczywiste wymiary
  setTimeout(() => {
    const rect = clone.getBoundingClientRect();
    const w = Math.ceil(rect.width);
    const h = Math.ceil(rect.height);

    html2canvas(clone, {
      scale: 1.4,                    // <--- ZMIENIONE – mniejsza rozdzielczość → PDF otwiera się w czytelniejszej skali
      useCORS: true,
      backgroundColor: '#ffffff',
      width: w,
      height: h,
      windowWidth: w,
      windowHeight: h,
      x: 0,
      y: 0,
      scrollX: 0,
      scrollY: 0,
      logging: false,
      allowTaint: true
    }).then(canvas => {
      const { jsPDF } = window.jspdf;

      let orientation = 'landscape';
      if (h > w * 1.2) orientation = 'portrait';

      const pdf = new jsPDF({
        orientation: orientation,
        unit: 'px',
        format: [canvas.width, canvas.height],
        compress: true
      });

      pdf.addImage(
        canvas.toDataURL('image/jpeg', 0.92),
        'JPEG',
        0,
        0,
        canvas.width,
        canvas.height
      );

      pdf.save('Euroeast_World_Food.pdf');

      document.body.removeChild(clone);
    }).catch(err => {
      console.error('PDF error:', err);
      alert('Błąd generowania PDF – sprawdź konsolę');
      document.body.removeChild(clone);
    });
  }, 1200);
}
