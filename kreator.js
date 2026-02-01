// =======================================================
// script.js – FINALNA WERSJA (poprawiony PDF – bez ucinania dołu)
// PDF: bez „Koniec okresu programu” + widoczny pasek postępu
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
          ? 'Niestety, tym razem się nie udało. Popracuj ze swoim przedstawicielem w kolejnym etapie i razem osiągniecie cel.'
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
// PDF – poprawiona wersja – pełna wysokość + margines bezpieczeństwa
// =======================================================
function generatePDF() {
  const capture = document.getElementById('capture');

  // Ukrywamy elementy, które nie mają być w PDF
  const hiddenEls = [
    ...capture.querySelectorAll(
      'select, input[type="file"], #end-of-period, label[for="end-of-period"], #excelFile'
    ),
    ...Array.from(capture.querySelectorAll('*')).filter(el => {
      const t = el.textContent ? el.textContent.trim() : '';
      return t === 'Importuj Excel:' || t === 'Koniec okresu programu';
    })
  ];

  hiddenEls.forEach(el => (el.style.visibility = 'hidden'));

  // Dajemy przeglądarce chwilę na przeliczenie layoutu
  setTimeout(() => {
    // Pobieramy rzeczywistą wysokość zawartości + spory zapas
    const contentHeight = capture.scrollHeight + 120;

    html2canvas(capture, {
      scale: 2,
      useCORS: true,
      backgroundColor: '#ffffff',
      width: capture.scrollWidth,
      height: contentHeight,
      windowWidth: capture.scrollWidth,
      windowHeight: contentHeight,
      logging: false
    }).then(canvas => {
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF('landscape', 'pt', 'a4');

      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();

      const margin = 40;
      const maxWidth = pageWidth - margin * 2;
      const ratio = maxWidth / canvas.width;

      const imgWidth = canvas.width * ratio;
      const imgHeight = canvas.height * ratio;

      // Centrujemy pionowo i poziomo
      const x = (pageWidth - imgWidth) / 2;
      let y = (pageHeight - imgHeight) / 2;

      // Jeśli obraz jest wyższy niż strona → dopasowujemy do wysokości
      if (imgHeight > pageHeight - margin * 2) {
        y = margin;
      }

      pdf.addImage(canvas.toDataURL('image/png'), 'PNG', x, y, imgWidth, imgHeight);
      pdf.save('EuroEast_Dashboard.pdf');

      // Przywracamy widoczność
      hiddenEls.forEach(el => (el.style.visibility = 'visible'));
    });
  }, 600);
}