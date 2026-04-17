(() => {
  function normalizeCatalogPdfText(value){
    return String(value ?? '')
      .replace(/[\u0000-\u001f\u007f]/g, ' ')
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
      .replace(/\s+/g, ' ')
      .trim();
  }

  function findColumn(cols, variants){
    const lower = cols.map(c => c.toLowerCase());
    for(const v of variants){
      const i = lower.indexOf(v.toLowerCase());
      if(i !== -1) return cols[i];
    }
    return null;
  }

  const imageCache = new Map();
  const PRODUCT_IMAGE_EXTENSIONS = ['webp', 'png', 'jpg', 'jpeg'];

  async function loadImageAsJpeg(url, maxDim){
    if(imageCache.has(url)) return imageCache.get(url);
    const pending = (async () => {
      const res = await fetch(url);
      if(!res.ok) throw new Error('Image not found');
      const blob = await res.blob();
      const img = new Image();
      img.crossOrigin = 'anonymous';
      const objectUrl = URL.createObjectURL(blob);
      try{
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
        return { dataUrl: jpg, width: canvas.width, height: canvas.height };
      }finally{
        URL.revokeObjectURL(objectUrl);
      }
    })().catch(error => {
      imageCache.delete(url);
      throw error;
    });
    imageCache.set(url, pending);
    return pending;
  }

  async function getProductImage(index, baseOverride){
    const base = baseOverride || window.imageBaseUrl || '';
    const normalizedIndex = String(index ?? '').trim();
    const preferredExt = typeof window.getPreferredImageExt === 'function'
      ? window.getPreferredImageExt(normalizedIndex, base)
      : (String(base).includes('Ukraina%2F') ? 'png' : 'webp');
    const extOrder = [preferredExt, ...PRODUCT_IMAGE_EXTENSIONS.filter(ext => ext !== preferredExt)];
    for(const ext of extOrder){
      const url = `${base}${encodeURIComponent(normalizedIndex)}.${ext}?alt=media`;
      try{
        const imageData = await loadImageAsJpeg(url, 900);
        if(typeof window.rememberImageExt === 'function'){
          window.rememberImageExt(normalizedIndex, base, ext);
        }
        return imageData;
      }catch(_){
        // Try the next supported extension.
      }
    }
    throw new Error('Image not found');
  }

  const watermarkUrl =
    'https://firebasestorage.googleapis.com/v0/b/pdf-creator-f7a8b.firebasestorage.app/o/CREATOR%20BASIC%2Fmaspo%20logo.png?alt=media&token=8fe33ebe-04d2-42cb-acb0-10379dbd7e11';

  async function loadImageAsPng(url){
    if(imageCache.has(url)) return imageCache.get(url);
    const pending = (async () => {
      const res = await fetch(url);
      if(!res.ok) throw new Error('Image not found');
      const blob = await res.blob();
      const img = new Image();
      img.crossOrigin = 'anonymous';
      const objectUrl = URL.createObjectURL(blob);
      try{
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
        return { dataUrl: png, width: canvas.width, height: canvas.height };
      }finally{
        URL.revokeObjectURL(objectUrl);
      }
    })().catch(error => {
      imageCache.delete(url);
      throw error;
    });
    imageCache.set(url, pending);
    return pending;
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
    const sectionHeaderH = options.groupByField ? 64 : 0;
    const cardW = (pageW - margin * 2 - gap * (cols - 1)) / cols;
    const cardH = (pageH - margin * 2 - sectionHeaderH - gap * (rows - 1)) / rows;

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
    const groupByField = options.groupByField || '';

    const hasCover = !!options.coverDataUrl;
    const watermarkPromise = loadImageAsPng(watermarkUrl).catch(() => null);
    const groupedSections = groupByField
      ? products.reduce((acc, product) => {
          const groupName = normalizeCatalogPdfText(
            getValueByVariants(product, [groupByField, 'grupa', 'group', 'grupa produktowa']) || product[groupByField] || ''
          );
          if(!groupName) return acc;
          let section = acc.find(item => item.name === groupName);
          if(!section){
            section = { name: groupName, items: [] };
            acc.push(section);
          }
          section.items.push(product);
          return acc;
        }, [])
      : [{ name: '', items: products }];
    const contentPageCount = groupedSections.reduce((sum, section) => sum + Math.ceil(section.items.length / (cols * rows)), 0);
    const totalPages = contentPageCount + (hasCover ? 1 : 0);

    const preloadTasks = [];
    const seenProductImages = new Set();
    const seenBadges = new Set();

    products.forEach(product => {
      const indexVal = String(getValueByVariants(product, ['indeks', 'index', 'id']) || product[indexKey] || '').trim();
      if(indexVal){
        const sourceHint = String(product.__country || product.__source || product[countryKey] || '').toLowerCase();
        const baseOverride = sourceHint.includes('ukraina')
          ? (window.imageBaseUrlUkraina || window.imageBaseUrl || '')
          : (window.imageBaseUrlRumunia || window.imageBaseUrl || '');
        const cacheKey = `${baseOverride}::${indexVal}`;
        if(!seenProductImages.has(cacheKey)){
          seenProductImages.add(cacheKey);
          preloadTasks.push(getProductImage(indexVal, baseOverride).catch(() => null));
        }
      }

      const countryVal = String(product[countryKey] || 'Rumunia').trim().toLowerCase();
      const badgeUrl = getCountryBadgeUrl(countryVal);
      if(badgeUrl && !seenBadges.has(badgeUrl)){
        seenBadges.add(badgeUrl);
        preloadTasks.push(loadImageAsJpeg(badgeUrl).catch(() => null));
      }
    });

    preloadTasks.push(watermarkPromise);
    await Promise.all(preloadTasks);

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

    let pageNumber = hasCover ? 2 : 1;
    let hasRenderedAnyContentPage = false;

    for(const section of groupedSections){
      const sectionItems = section.items || [];
      for(let i = 0; i < sectionItems.length; i++){
        const pageIndex = i % (cols * rows);
        const isNewSectionPage = i === 0;
        const isPageBreakWithinSection = i > 0 && pageIndex === 0;

        if(hasRenderedAnyContentPage && (isNewSectionPage || isPageBreakWithinSection)){
          pdf.addPage();
          pageNumber += 1;
        }
        if(!hasRenderedAnyContentPage){
          hasRenderedAnyContentPage = true;
        }

        const p = sectionItems[i];
        // no global watermark; we will add small logo per card
        const col = pageIndex % cols;
        const row = Math.floor(pageIndex / cols);
        const x = margin + col * (cardW + gap);
        const y = margin + sectionHeaderH + row * (cardH + gap);

        if(pageIndex === 0 && section.name){
          pdf.setFont('helvetica', 'bold');
          pdf.setFontSize(22);
          pdf.setTextColor(15, 27, 45);
          pdf.text(section.name, margin, margin + 22);
          pdf.setDrawColor(22, 138, 63);
          pdf.setLineWidth(2);
          pdf.line(margin, margin + 34, pageW - margin, margin + 34);
          pdf.setTextColor(0);
        }

        pdf.setDrawColor(220);
        pdf.setFillColor(255, 255, 255);
        pdf.roundedRect(x, y, cardW, cardH, 12, 12, 'FD');

        const name = normalizeCatalogPdfText(getValueByVariants(p, ['nazwa', 'name']) || p[nameKey] || '');
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

        const indexVal = String(getValueByVariants(p, ['indeks', 'index', 'id']) || p[indexKey] || '').trim();
        let imgData = null;
        if(indexVal){
          try{
            const sourceHint = String(p.__country || p.__source || p[countryKey] || '').toLowerCase();
            const baseOverride = sourceHint.includes('ukraina')
              ? (window.imageBaseUrlUkraina || window.imageBaseUrl || '')
              : (window.imageBaseUrlRumunia || window.imageBaseUrl || '');
            imgData = await getProductImage(indexVal, baseOverride);
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
          const wm = await watermarkPromise;
          if(wm){
            const wmW = 48;
            const wmH = (wm.height / wm.width) * wmW;
            const wmX = x + cardW - wmW - 20;
            const wmY = y + cardH - 94;
            pdf.setGState(new pdf.GState({ opacity: 0.5 }));
            pdf.addImage(wm.dataUrl, 'PNG', wmX, wmY, wmW, wmH, undefined, 'FAST');
            pdf.setGState(new pdf.GState({ opacity: 1 }));
          }
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
        const eanVal = String(getValueByVariants(p, ['kod ean', 'ean']) || p[eanKey] || '').trim();
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
    }

    return pdf.output('blob');
  };

  function getValueByVariants(obj, variants){
    if(!obj) return '';
    const keys = Object.keys(obj);
    for(const v of variants){
      const key = keys.find(k => k.toLowerCase() === v.toLowerCase());
      if(key && obj[key] !== undefined && obj[key] !== null) return obj[key];
    }
    return '';
  }

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
    try{
      JsBarcode(canvas, normalized.data, {
        format: normalized.format,
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
    const digits = String(raw || '').replace(/\D/g, '');
    if(!digits) return null;
    let e = digits;
    if(e.length > 13){
      e = e.slice(-13);
    }
    if(e.length === 13){
      const base = e.slice(0, 12);
      return { data: base + calcEAN13Check(base), format: 'EAN13' };
    }
    if(e.length === 12){
      return { data: e + calcEAN13Check(e), format: 'EAN13' };
    }
    if(e.length === 8){
      return { data: e, format: 'EAN8' };
    }
    if(e.length === 7){
      return { data: e + calcEAN8Check(e), format: 'EAN8' };
    }
    return null;
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
