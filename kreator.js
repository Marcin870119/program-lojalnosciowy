(() => {
  function findColumn(cols, variants){
    const lower = cols.map(c => c.toLowerCase());
    for(const v of variants){
      const i = lower.indexOf(v.toLowerCase());
      if(i !== -1) return cols[i];
    }
    return null;
  }

  const imageCache = new Map();

  async function loadImageAsJpeg(url, maxDim){
    if(imageCache.has(url)) return imageCache.get(url);
    const res = await fetch(url);
    if(!res.ok) throw new Error('Image not found');
    const blob = await res.blob();
    const img = new Image();
    img.crossOrigin = 'anonymous';
    const objectUrl = URL.createObjectURL(blob);
    const loadedImg = await new Promise((resolve, reject) => {
      img.onload = () => resolve(img);
      img.onerror = reject;
      img.src = objectUrl;
    });
    const canvas = document.createElement('canvas');
    let w = loadedImg.naturalWidth;
    let h = loadedImg.naturalHeight;
    if(maxDim && (w > maxDim || h > maxDim)){
      const scale = Math.min(maxDim / w, maxDim / h);
      w = Math.round(w * scale);
      h = Math.round(h * scale);
    }
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(loadedImg, 0, 0, w, h);
    const jpg = canvas.toDataURL('image/jpeg', 0.75);
    URL.revokeObjectURL(objectUrl);
    const result = { dataUrl: jpg, width: canvas.width, height: canvas.height };
    imageCache.set(url, result);
    return result;
  }

  async function getProductImage(index){
    const base = window.imageBaseUrl || '';
    const png = `${base}${encodeURIComponent(index)}.png?alt=media`;
    const jpg = `${base}${encodeURIComponent(index)}.jpg?alt=media`;
    try{
      return await loadImageAsJpeg(png, 900);
    }catch(_){
      return await loadImageAsJpeg(jpg, 900);
    }
  }

  const watermarkUrl =
    'https://firebasestorage.googleapis.com/v0/b/pdf-creator-f7a8b.firebasestorage.app/o/CREATOR%20BASIC%2Fmaspo%20logo.png?alt=media&token=8fe33ebe-04d2-42cb-acb0-10379dbd7e11';

  async function loadImageAsPng(url){
    if(imageCache.has(url)) return imageCache.get(url);
    const res = await fetch(url);
    if(!res.ok) throw new Error('Image not found');
    const blob = await res.blob();
    const img = new Image();
    img.crossOrigin = 'anonymous';
    const objectUrl = URL.createObjectURL(blob);
    const loadedImg = await new Promise((resolve, reject) => {
      img.onload = () => resolve(img);
      img.onerror = reject;
      img.src = objectUrl;
    });
    const canvas = document.createElement('canvas');
    canvas.width = loadedImg.naturalWidth;
    canvas.height = loadedImg.naturalHeight;
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(loadedImg, 0, 0);
    const png = canvas.toDataURL('image/png');
    URL.revokeObjectURL(objectUrl);
    const result = { dataUrl: png, width: canvas.width, height: canvas.height };
    imageCache.set(url, result);
    return result;
  }

  window.createCatalogPdf = async function(products, options = {}){
    if(!products || !products.length) throw new Error('Brak produktów');
    const { jsPDF } = window.jspdf;
    const pageW = 794;
    const pageH = 1123;
    const margin = 30;
    const gap = 16;
    const cols = 2;
    const rows = 3;
    const cardW = (pageW - margin * 2 - gap * (cols - 1)) / cols;
    const cardH = (pageH - margin * 2 - gap * (rows - 1)) / rows;

    const pdf = new jsPDF({
      orientation: 'p',
      unit: 'px',
      format: [pageW, pageH],
      compress: true
    });

    const colsKeys = Object.keys(products[0] || {});
    const nameKey = findColumn(colsKeys, ['nazwa', 'name']);
    const indexKey = findColumn(colsKeys, ['indeks', 'index', 'id']);
    const eanKey = findColumn(colsKeys, ['kod ean', 'ean']);
    const countryKey = findColumn(colsKeys, ['krajpochodzenia', 'kraj pochodzenia', 'country']);
    const priceMap = options.priceMap || null;
    const currency = options.currency || 'EUR';
    const priceColor = options.priceColor || '#000000';

    const hasCover = !!options.coverDataUrl;
    const totalPages = Math.ceil(products.length / (cols * rows)) + (hasCover ? 1 : 0);

    if(hasCover){
      const coverImg = await loadCoverImage(options.coverDataUrl);
      if(coverImg){
        const scale = Math.min(pageW / coverImg.width, pageH / coverImg.height);
        const cw = coverImg.width * scale;
        const ch = coverImg.height * scale;
        const cx = (pageW - cw) / 2;
        const cy = (pageH - ch) / 2;
        pdf.addImage(coverImg.dataUrl, 'JPEG', cx, cy, cw, ch, undefined, 'FAST');
      }
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(10);
      pdf.setTextColor(120);
      pdf.text(`Strona 1 / ${totalPages}`, pageW - 110, pageH - 12);
      pdf.setTextColor(0);
      pdf.addPage();
    }

    for(let i = 0; i < products.length; i++){
      if(i > 0 && i % (cols * rows) === 0){
        pdf.addPage();
      }

      const p = products[i];
      const pageIndex = i % (cols * rows);
      const pageNumber = Math.floor(i / (cols * rows)) + 1 + (hasCover ? 1 : 0);
      // no global watermark; we will add small logo per card
      const col = pageIndex % cols;
      const row = Math.floor(pageIndex / cols);
      const x = margin + col * (cardW + gap);
      const y = margin + row * (cardH + gap);

      pdf.setDrawColor(220);
      pdf.setFillColor(255, 255, 255);
      pdf.roundedRect(x, y, cardW, cardH, 12, 12, 'FD');

      const name = String(p[nameKey] || '').trim();
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(12);
      const nameLines = pdf.splitTextToSize(name, cardW - 20).slice(0, 2);
      pdf.text(nameLines, x + 10, y + 20);

      const countryVal = String(p[countryKey] || 'Rumunia').trim().toLowerCase();
      const badgeUrl = getCountryBadgeUrl(countryVal);
      if(badgeUrl){
        try{
          const badge = await loadImageAsJpeg(badgeUrl);
          const bw = 68;
          const bh = 44;
          const bx = x + 4;
          const by = y + cardH - bh - 36;
          pdf.addImage(badge.dataUrl, 'JPEG', bx, by, bw, bh, undefined, 'FAST');
        }catch(_){}
      }

      const indexVal = String(p[indexKey] || '').trim();
      let imgData = null;
      if(indexVal){
        try{
          imgData = await getProductImage(indexVal);
        }catch(_){}
      }

      if(imgData){
        const maxW = cardW - 90;
        const maxH = cardH - 150;
        const scale = Math.min(maxW / imgData.width, maxH / imgData.height, 1);
        const iw = imgData.width * scale;
        const ih = imgData.height * scale;
        const ix = x + (cardW - iw) / 2;
        const iy = y + 40 + (maxH - ih) / 2;
        pdf.addImage(imgData.dataUrl, 'JPEG', ix, iy, iw, ih, undefined, 'FAST');

        // small logo near each image
        try{
          const wm = await loadImageAsPng(watermarkUrl);
          const wmW = 48;
          const wmH = (wm.height / wm.width) * wmW;
          const wmX = x + cardW - wmW - 20;
          const wmY = y + cardH - 94;
          pdf.setGState(new pdf.GState({ opacity: 0.5 }));
          pdf.addImage(wm.dataUrl, 'PNG', wmX, wmY, wmW, wmH, undefined, 'FAST');
          pdf.setGState(new pdf.GState({ opacity: 1 }));
        }catch(_){}
      }

      // price from excel
      if(priceMap && indexVal && priceMap.get(indexVal)){
        const priceInfo = priceMap.get(indexVal);
        const priceText = String(priceInfo.price || '').replace(',', '.');
        const unitRaw = String(priceInfo.unit || '').toUpperCase();
        const unit = unitRaw.includes('KG') ? 'KG' : 'SZT.';
        const [main, dec] = formatPriceParts(priceText);
        const symbol = currency === 'PLN' ? 'PLN' : (currency === 'GBP' ? '£' : '€');

        const px = x + cardW - 90;
        const py = y + cardH - 150;

        pdf.setFont('helvetica', 'bold');
        pdf.setFontSize(52);
        const [pr, pg, pb] = hexToRgb(priceColor);
        pdf.setTextColor(pr, pg, pb);
        pdf.text(main, px, py);

        pdf.setFont('helvetica', 'bold');
        pdf.setFontSize(18);
        pdf.text(dec, px + 24, py - 12);

        pdf.setFont('helvetica', 'normal');
        pdf.setFontSize(12);
        pdf.text(`${symbol} / ${unit}`, px + 24, py + 4);
        pdf.setTextColor(0);
      }

      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(12);
      if(indexVal){
        pdf.text(`Indeks: ${indexVal}`, x + 10, y + cardH - 22);
      }
      const eanVal = String(p[eanKey] || '').trim();
      if(eanVal){
        const barcode = buildBarcodeImage(eanVal);
        if(barcode){
          const bw = 220;
          const bh = 60;
          const bx = x + cardW - bw - 12;
          const by = y + cardH - bh - 14;
          pdf.addImage(barcode, 'PNG', bx, by, bw, bh, undefined, 'FAST');
        }
      }

      if(pageIndex === 0){
        pdf.setFont('helvetica', 'normal');
        pdf.setFontSize(10);
        pdf.setTextColor(120);
        pdf.text(`Strona ${pageNumber} / ${totalPages}`, pageW - 110, pageH - 12);
        pdf.setTextColor(0);
      }
    }

    return pdf.output('blob');
  };

  async function loadCoverImage(dataUrl){
    try{
      const img = new Image();
      img.crossOrigin = 'anonymous';
      await new Promise((resolve, reject) => {
        img.onload = resolve;
        img.onerror = reject;
        img.src = dataUrl;
      });
      const canvas = document.createElement('canvas');
      canvas.width = img.naturalWidth;
      canvas.height = img.naturalHeight;
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, 0, 0);
      const jpg = canvas.toDataURL('image/jpeg', 0.9);
      return { dataUrl: jpg, width: canvas.width, height: canvas.height };
    }catch(_){
      return null;
    }
  }

  function buildBarcodeImage(raw){
    const normalized = normalizeEAN(raw);
    if(!normalized) return null;
    const canvas = document.createElement('canvas');
    const format = normalized.length === 8 ? 'EAN8' : 'EAN13';
    try{
      const data = format === 'EAN8' ? normalized.slice(0, 7) : normalized.slice(0, 12);
      JsBarcode(canvas, data, {
        format,
        width: 2,
        height: 50,
        displayValue: true,
        fontSize: 12,
        margin: 6,
        background: '#ffffff',
        lineColor: '#000000'
      });
      return canvas.toDataURL('image/png');
    }catch(_){
      return null;
    }
  }

  function normalizeEAN(raw){
    const e = String(raw || '').replace(/\D/g, '');
    if(e.length >= 13){
      const base = e.slice(0, 12);
      return base + calcEAN13Check(base);
    }
    if(e.length === 12) return e + calcEAN13Check(e);
    if(e.length === 8){
      const base = e.slice(0, 7);
      return base + calcEAN8Check(base);
    }
    if(e.length === 7) return e + calcEAN8Check(e);
    if(e.length < 7) return null;
    return e;
  }

  function calcEAN13Check(code12){
    const digits = code12.split('').map(Number);
    let sum = 0;
    for(let i=0;i<12;i++){
      sum += digits[i] * (i % 2 === 0 ? 1 : 3);
    }
    return String((10 - (sum % 10)) % 10);
  }

  function calcEAN8Check(code7){
    const digits = code7.split('').map(Number);
    let sum = 0;
    for(let i=0;i<7;i++){
      sum += digits[i] * (i % 2 === 0 ? 3 : 1);
    }
    return String((10 - (sum % 10)) % 10);
  }


  function getCountryBadgeUrl(country){
    if(country.includes('rumunia')) return 'https://firebasestorage.googleapis.com/v0/b/pdf-creator-f7a8b.firebasestorage.app/o/CREATOR%20BASIC%2FRumunia.png?alt=media';
    if(country.includes('ukraina')) return 'https://firebasestorage.googleapis.com/v0/b/pdf-creator-f7a8b.firebasestorage.app/o/CREATOR%20BASIC%2FUkraina.png?alt=media';
    if(country.includes('litwa')) return 'https://firebasestorage.googleapis.com/v0/b/pdf-creator-f7a8b.firebasestorage.app/o/CREATOR%20BASIC%2FLitwa.png?alt=media';
    if(country.includes('bulgaria')) return 'https://firebasestorage.googleapis.com/v0/b/pdf-creator-f7a8b.firebasestorage.app/o/CREATOR%20BASIC%2FBulgaria.png?alt=media';
    if(country.includes('polska')) return 'https://firebasestorage.googleapis.com/v0/b/pdf-creator-f7a8b.firebasestorage.app/o/CREATOR%20BASIC%2FPolska.png?alt=media';
    return 'https://firebasestorage.googleapis.com/v0/b/pdf-creator-f7a8b.firebasestorage.app/o/CREATOR%20BASIC%2FRumunia.png?alt=media';
  }

  function formatPriceParts(priceText){
    let val = parseFloat(priceText);
    if(!Number.isFinite(val)) return ['0', '00'];
    const fixed = val.toFixed(2);
    const [main, dec] = fixed.split('.');
    return [main, dec];
  }

  function hexToRgb(hex){
    const h = String(hex || '').replace('#','');
    if(h.length !== 6) return [0,0,0];
    return [parseInt(h.slice(0,2),16), parseInt(h.slice(2,4),16), parseInt(h.slice(4,6),16)];
  }
})();
