const VAT_RATE = 0.12;
const SAMPLE_FILE = 'Заказы с товарами 25-27.05.2026г.xls';
const TEMPLATE_FILES = [
  'ШАБЛОН СЧФ для КОРЗИНКИ-365.xlsx',
  'ШАБЛОН СЧФ для КОРЗИНКИ-365.xls',
  'ШАБЛОН СЧФ для КОРЗИНКИ-365.xlsm',
];

const state = {
  fileName: '',
  rawRows: [],
  stores: [],
  warnings: [],
};

const els = {
  tabs: document.querySelectorAll('.tab'),
  panels: document.querySelectorAll('.panel'),
  fileInput: document.querySelector('#fileInput'),
  dropzone: document.querySelector('.dropzone'),
  statusText: document.querySelector('#statusText'),
  summaryList: document.querySelector('#summaryList'),
  resultsContent: document.querySelector('#resultsContent'),
  invoiceContent: document.querySelector('#invoiceContent'),
  loadSampleBtn: document.querySelector('#loadSampleBtn'),
  clearBtn: document.querySelector('#clearBtn'),
  exportCsvBtn: document.querySelector('#exportCsvBtn'),
  exportExcelBtn: document.querySelector('#exportExcelBtn'),
  exportInvoiceExcelBtn: document.querySelector('#exportInvoiceExcelBtn'),
  printBtn: document.querySelector('#printBtn'),
  buildInvoiceBtn: document.querySelector('#buildInvoiceBtn'),
  invoiceDate: document.querySelector('#invoiceDate'),
  supplierName: document.querySelector('#supplierName'),
  invoicePrefix: document.querySelector('#invoicePrefix'),
  invoiceMode: document.querySelector('#invoiceMode'),
  showVat: document.querySelector('#showVat'),
  groupByPrice: document.querySelector('#groupByPrice'),
};

els.invoiceDate.valueAsDate = new Date();

els.tabs.forEach((tab) => {
  tab.addEventListener('click', () => activateTab(tab.dataset.tab));
});

els.fileInput.addEventListener('change', (event) => {
  const [file] = event.target.files;
  if (file) readFile(file);
});

['dragenter', 'dragover'].forEach((eventName) => {
  els.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    els.dropzone.classList.add('dragover');
  });
});

['dragleave', 'drop'].forEach((eventName) => {
  els.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    els.dropzone.classList.remove('dragover');
  });
});

els.dropzone.addEventListener('drop', (event) => {
  const [file] = event.dataTransfer.files;
  if (file) readFile(file);
});

els.loadSampleBtn.addEventListener('click', async () => {
  try {
    setStatus('Repositorydagi namuna yuklanmoqda...');
    const response = await fetch(encodeURI(SAMPLE_FILE));
    if (!response.ok) throw new Error('Namuna fayl topilmadi. Saytni lokal server orqali oching.');
    const buffer = await response.arrayBuffer();
    processWorkbook(buffer, SAMPLE_FILE);
  } catch (error) {
    showError(error.message);
  }
});

els.clearBtn.addEventListener('click', () => {
  state.fileName = '';
  state.rawRows = [];
  state.stores = [];
  state.warnings = [];
  els.fileInput.value = '';
  setStatus('Hali fayl yuklanmadi.');
  els.summaryList.innerHTML = '';
  els.resultsContent.className = 'empty-state';
  els.resultsContent.textContent = 'Avval Excel fayl yuklang.';
  els.invoiceContent.className = 'invoice-preview empty-state';
  els.invoiceContent.textContent = 'Avval Excel fayl yuklang.';
});

els.exportCsvBtn.addEventListener('click', exportCsv);
els.exportExcelBtn.addEventListener('click', () => exportExcel());
els.exportInvoiceExcelBtn.addEventListener('click', () => exportExcel());
els.printBtn.addEventListener('click', () => {
  activateTab('invoice');
  window.print();
});
els.buildInvoiceBtn.addEventListener('click', renderInvoice);
[els.supplierName, els.invoicePrefix, els.invoiceDate, els.invoiceMode, els.showVat].forEach((el) => {
  el.addEventListener('change', renderInvoice);
  el.addEventListener('input', renderInvoice);
});
els.groupByPrice.addEventListener('change', () => {
  if (state.rawRows.length) {
    const parsed = parseRows(state.rawRows);
    state.stores = parsed.stores;
    state.warnings = parsed.warnings;
    renderSummary();
    renderResults();
  }
  renderInvoice();
});

function activateTab(name) {
  els.tabs.forEach((tab) => tab.classList.toggle('active', tab.dataset.tab === name));
  els.panels.forEach((panel) => panel.classList.toggle('active', panel.id === name));
}

function readFile(file) {
  const reader = new FileReader();
  reader.onload = (event) => processWorkbook(event.target.result, file.name);
  reader.onerror = () => showError('Faylni o‘qishda xatolik yuz berdi.');
  setStatus(`${file.name} o‘qilmoqda...`);
  reader.readAsArrayBuffer(file);
}

function processWorkbook(buffer, fileName) {
  if (!window.XLSX) {
    showError('XLSX kutubxonasi yuklanmadi. Internet ulanishini tekshiring yoki xlsx.full.min.js faylini lokal ulang.');
    return;
  }

  try {
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: false, raw: false });
    const allRows = [];

    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false, blankrows: false });
      rows.forEach((row) => allRows.push(row));
    });

    const parsed = parseRows(allRows);
    state.fileName = fileName;
    state.rawRows = allRows;
    state.stores = parsed.stores;
    state.warnings = parsed.warnings;

    renderSummary();
    renderResults();
    renderInvoice();
    activateTab('results');
  } catch (error) {
    showError(`Excel faylni tahlil qilib bo‘lmadi: ${error.message}`);
  }
}

function parseRows(rows) {
  const flatResult = parseFlatTableRows(rows);
  if (flatResult.stores.length) return flatResult;

  const warnings = [];
  const storeMap = new Map();
  let currentStore = null;
  let currentTotal = 0;
  let productColumns = null;

  rows.forEach((row, rowIndex) => {
    const normalized = row.map(normalizeText);

    const shopHeader = findShopHeader(normalized);
    if (shopHeader) {
      const dataRow = findNextDataRow(rows, rowIndex + 1, shopHeader.nameIndex);
      if (dataRow) {
        const storeName = cleanName(dataRow.row[shopHeader.nameIndex]);
        currentTotal = parseNumber(dataRow.row[shopHeader.totalIndex]);
        if (storeName) currentStore = getStore(storeMap, storeName, currentTotal);
      }
      productColumns = null;
      return;
    }

    const directShopIndex = normalized.findIndex((cell) => cell === 'магазин');
    if (directShopIndex !== -1 && !currentStore) {
      const name = cleanName(row[directShopIndex + 1]) || cleanName(row[directShopIndex]);
      if (name && name.toLowerCase() !== 'магазин') currentStore = getStore(storeMap, name, 0);
    }

    const detectedProductColumns = findProductColumns(normalized);
    if (detectedProductColumns) {
      productColumns = detectedProductColumns;
      return;
    }

    if (!productColumns || !currentStore) return;
    if (isSectionBreak(normalized)) return;

    const productName = cleanName(row[productColumns.product]);
    if (!productName || isIgnoredProductName(productName)) return;

    const quantity = parseNumber(row[productColumns.quantity]);
    const price = parseNumber(row[productColumns.price]);
    const net = parseNumber(row[productColumns.net]) || quantity * price;
    const vatCell = parseNumber(row[productColumns.vat]);
    const vat = vatCell && vatCell !== 12 ? vatCell : net * VAT_RATE;
    const gross = parseNumber(row[productColumns.gross]) || net + vat;

    if (!quantity && !net && !gross) return;

    addProduct(currentStore, {
      name: productName,
      quantity,
      price: price || (quantity ? net / quantity : 0),
      net,
      vat,
      gross,
      sourceRow: rowIndex + 1,
    }, els.groupByPrice.checked);
  });

  const stores = Array.from(storeMap.values()).map((store) => ({
    ...store,
    products: Array.from(store.products.values()).sort((a, b) => a.name.localeCompare(b.name, 'ru')),
    orderNumbers: Array.from(store.orderNumbers || []),
  })).sort((a, b) => a.name.localeCompare(b.name, 'ru'));

  if (!stores.length) warnings.push('Magazin va mahsulot qatorlari topilmadi. Sarlavhalar nomini tekshiring.');
  stores.forEach((store) => {
    if (!store.products.length) warnings.push(`${store.name} uchun mahsulotlar topilmadi.`);
  });

  return { stores, warnings };
}


function parseFlatTableRows(rows) {
  const headerInfo = findFlatHeader(rows);
  if (!headerInfo) return { stores: [], warnings: [] };

  const { headerIndex, columns } = headerInfo;
  const productCol = findColumn(columns, ({ header }) => header === 'товар' || header.includes('товар'));
  const quantityCol = findColumn(columns, ({ header }) => header.includes('кол-во') || header.includes('количество'));
  const priceCol = findColumn(columns, ({ header }) => header.includes('цена за единицу') || header === 'цена');
  const netCol = findColumn(columns, ({ header, full }) => header === 'стоимость' && !full.includes('ндс'));
  const vatRateCol = findColumn(columns, ({ header, parent }) => header === 'ставка' && parent.includes('ндс'));
  const vatAmountCol = findColumn(columns, ({ header, parent }) => header === 'сумма' && parent.includes('ндс'));
  const grossCol = findColumn(columns, ({ full }) => full.includes('стоимость с учетом ндс') || full.includes('стоимость с учётом ндс'));
  const totalCol = findColumn(columns, ({ full }) => full.includes('общая сумма'));
  const orderCol = findColumn(columns, ({ header, parent }) => header === 'номер' && parent.includes('заказ'));
  const storeNameCol = findColumn(columns, ({ header, parent }) => header.includes('название') && parent.includes('магазин'));

  if (storeNameCol === -1 || productCol === -1 || quantityCol === -1) return { stores: [], warnings: [] };

  const storeMap = new Map();
  const orderStorePairs = new Set();

  rows.slice(headerIndex + 1).forEach((row, offset) => {
    const sourceRow = headerIndex + offset + 2;
    const storeName = cleanName(row[storeNameCol]);
    const productName = cleanName(row[productCol]);
    if (!storeName || !productName || isIgnoredProductName(productName)) return;

    const quantity = parseNumber(row[quantityCol]);
    const price = parseNumber(row[priceCol]);
    const net = parseNumber(row[netCol]) || quantity * price;
    const vatRate = parseNumber(row[vatRateCol]) || 12;
    const vat = parseNumber(row[vatAmountCol]) || net * (vatRate / 100 || VAT_RATE);
    const gross = parseNumber(row[grossCol]) || net + vat;
    if (!quantity && !net && !gross) return;

    const orderNumber = cleanName(row[orderCol]) || findTenDigitOrderNumber(row);
    const declaredTotal = parseNumber(row[totalCol]);
    const store = getStore(storeMap, storeName, 0, false);
    const pairKey = `${normalizeText(storeName)}|${orderNumber || sourceRow}`;
    if (!orderStorePairs.has(pairKey)) {
      orderStorePairs.add(pairKey);
      store.parts += 1;
      store.declaredTotal += declaredTotal || 0;
    }

    addOrderNumber(store, orderNumber);

    addProduct(store, {
      name: productName,
      quantity,
      price: price || (quantity ? net / quantity : 0),
      net,
      vat,
      gross,
      sourceRow,
    }, els.groupByPrice.checked);
  });

  const stores = Array.from(storeMap.values()).map((store) => ({
    ...store,
    products: Array.from(store.products.values()).sort((a, b) => a.name.localeCompare(b.name, 'ru')),
    orderNumbers: Array.from(store.orderNumbers || []),
  })).filter((store) => store.products.length).sort((a, b) => a.name.localeCompare(b.name, 'ru'));

  return { stores, warnings: [] };
}

function findFlatHeader(rows) {
  for (let index = 0; index < Math.min(rows.length, 20); index += 1) {
    const headerRow = rows[index] || [];
    const normalized = headerRow.map(normalizeText);
    const hasProduct = normalized.some((cell) => cell === 'товар' || cell.includes('товар'));
    const hasQuantity = normalized.some((cell) => cell.includes('кол-во') || cell.includes('количество'));
    const hasPrice = normalized.some((cell) => cell.includes('цена за единицу'));
    if (!hasProduct || !hasQuantity || !hasPrice) continue;

    const parentRow = rows[Math.max(0, index - 1)] || [];
    const parents = fillMergedParents(parentRow, headerRow.length);
    const columns = headerRow.map((cell, columnIndex) => {
      const header = normalizeText(cell);
      const parent = normalizeText(parents[columnIndex]);
      return {
        index: columnIndex,
        header,
        parent,
        full: `${parent} ${header}`.trim(),
      };
    });
    return { headerIndex: index, columns };
  }
  return null;
}

function fillMergedParents(parentRow, length) {
  const parents = [];
  let current = '';
  for (let index = 0; index < length; index += 1) {
    const value = cleanName(parentRow[index]);
    if (value) current = value;
    parents[index] = current;
  }
  return parents;
}

function findColumn(columns, predicate) {
  const column = columns.find(predicate);
  return column ? column.index : -1;
}

function findShopHeader(normalized) {
  const nameIndex = normalized.findIndex((cell) => cell === 'название' || cell.includes('название'));
  const totalIndex = normalized.findIndex((cell) => cell.includes('общая сумма'));
  const hasShop = normalized.some((cell) => cell.includes('магазин'));
  if (nameIndex !== -1 && (totalIndex !== -1 || hasShop)) {
    return { nameIndex, totalIndex };
  }
  return null;
}

function findNextDataRow(rows, start, nameIndex) {
  for (let i = start; i < Math.min(rows.length, start + 6); i += 1) {
    const row = rows[i] || [];
    const name = cleanName(row[nameIndex]);
    const normalized = row.map(normalizeText);
    if (name && !normalized.includes('название') && !normalized.includes('товар')) return { row, index: i };
  }
  return null;
}

function findProductColumns(normalized) {
  const product = normalized.findIndex((cell) => cell === 'товар' || cell.includes('товар'));
  const quantity = normalized.findIndex((cell) => cell.includes('кол-во') || cell.includes('количество'));
  const price = normalized.findIndex((cell) => cell.includes('цена за единицу') || cell === 'цена');
  const gross = normalized.findIndex((cell) => cell.includes('стоимость с учетом ндс') || cell.includes('стоимость с учётом ндс'));
  const vat = normalized.findIndex((cell, index) => index !== gross && (cell === 'ндс' || cell.includes('ндс')));
  const net = normalized.findIndex((cell, index) => index !== gross && (cell === 'стоимость' || cell.includes('стоимость')));

  if (product !== -1 && quantity !== -1) {
    return { product, quantity, price, net, vat, gross };
  }
  return null;
}

function isSectionBreak(normalized) {
  return normalized.some((cell) => ['итого', 'всего', 'название', 'магазин'].includes(cell));
}

function isIgnoredProductName(name) {
  const value = normalizeText(name);
  return ['товар', 'итого', 'всего'].includes(value) || value.includes('общая сумма');
}

function getStore(storeMap, name, declaredTotal, incrementParts = true) {
  const key = normalizeText(name);
  if (!storeMap.has(key)) {
    storeMap.set(key, { name, declaredTotal: 0, products: new Map(), parts: 0, orderNumbers: new Set() });
  }
  const store = storeMap.get(key);
  if (incrementParts) store.parts += 1;
  store.declaredTotal += declaredTotal || 0;
  return store;
}

function addOrderNumber(store, orderNumber) {
  if (!orderNumber) return;
  if (!store.orderNumbers) store.orderNumbers = new Set();
  store.orderNumbers.add(String(orderNumber));
}

function findTenDigitOrderNumber(row) {
  const match = row.map(cleanName).find((value) => /^\d{10}$/.test(value));
  return match || '';
}

function addProduct(store, product, groupByPrice) {
  const key = groupByPrice ? `${normalizeText(product.name)}|${roundMoney(product.price)}` : normalizeText(product.name);
  if (!store.products.has(key)) {
    store.products.set(key, { ...product, rows: [product.sourceRow] });
    return;
  }
  const existing = store.products.get(key);
  existing.quantity += product.quantity;
  existing.net += product.net;
  existing.vat += product.vat;
  existing.gross += product.gross;
  existing.price = existing.quantity ? existing.net / existing.quantity : product.price;
  existing.rows.push(product.sourceRow);
}

function normalizeText(value) {
  return String(value ?? '')
    .replace(/ё/g, 'е')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function cleanName(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function parseNumber(value) {
  if (typeof value === 'number') return value;
  const text = String(value ?? '').replace(/\s/g, '').replace('%', '').replace(',', '.').replace(/[^\d.-]/g, '');
  const number = Number.parseFloat(text);
  return Number.isFinite(number) ? number : 0;
}

function totals(products) {
  return products.reduce((acc, item) => {
    acc.quantity += item.quantity;
    acc.net += item.net;
    acc.vat += item.vat;
    acc.gross += item.gross;
    return acc;
  }, { quantity: 0, net: 0, vat: 0, gross: 0 });
}

function renderSummary() {
  const productCount = state.stores.reduce((sum, store) => sum + store.products.length, 0);
  const grand = totals(state.stores.flatMap((store) => store.products));
  setStatus(`<b>${state.fileName}</b> muvaffaqiyatli tahlil qilindi.`);
  els.summaryList.innerHTML = [
    `Magazinlar: <b>${state.stores.length}</b>`,
    `Birlashtirilgan mahsulot qatorlari: <b>${productCount}</b>`,
    `Umumiy summa NDS bilan: <b>${formatMoney(grand.gross)}</b>`,
    ...state.warnings.map((warning) => `<span class="error">${escapeHtml(warning)}</span>`),
  ].map((item) => `<li>${item}</li>`).join('');
}

function renderResults() {
  if (!state.stores.length) {
    els.resultsContent.className = 'empty-state';
    els.resultsContent.textContent = 'Tahlil natijasi yo‘q.';
    return;
  }

  els.resultsContent.className = 'store-grid';
  els.resultsContent.innerHTML = state.stores.map((store) => storeTable(store)).join('');
}

function storeTable(store) {
  const total = totals(store.products);
  return `
    <article class="store-card">
      <div class="store-head">
        <div><strong>${escapeHtml(store.name)}</strong><br><span class="badge">${store.parts} ta blok birlashtirildi</span></div>
        <div class="kpis">
          <span class="kpi">Tovar: ${store.products.length}</span>
          <span class="kpi">NDS: ${formatMoney(total.vat)}</span>
          <span class="kpi">Jami: ${formatMoney(total.gross)}</span>
        </div>
      </div>
      <div class="table-wrap">
        <table>
          <thead>${productHeader(true)}</thead>
          <tbody>${store.products.map(productRow).join('')}</tbody>
          <tfoot>${totalRow(total, true)}</tfoot>
        </table>
      </div>
    </article>`;
}

function renderInvoice() {
  if (!state.stores.length) {
    els.invoiceContent.className = 'invoice-preview empty-state';
    els.invoiceContent.textContent = 'Avval Excel fayl yuklang.';
    return;
  }

  const stores = rebuildStoresForCurrentGrouping();
  const mode = els.invoiceMode.value;
  els.invoiceContent.className = 'invoice-preview';
  els.invoiceContent.innerHTML = mode === 'single'
    ? invoiceDocument({ name: 'Umumiy nakladnoy', products: stores.flatMap((store) => store.products) }, 1)
    : stores.map((store, index) => invoiceDocument(store, index + 1)).join('');
}

function rebuildStoresForCurrentGrouping() {
  const storeMap = new Map();
  state.stores.forEach((store) => {
    const target = getStore(storeMap, store.name, store.declaredTotal);
    (store.orderNumbers || []).forEach((orderNumber) => addOrderNumber(target, orderNumber));
    store.products.forEach((product) => addProduct(target, product, els.groupByPrice.checked));
  });
  return Array.from(storeMap.values()).map((store) => ({ ...store, products: Array.from(store.products.values()), orderNumbers: Array.from(store.orderNumbers || []) }));
}

function invoiceDocument(store, number) {
  const total = totals(store.products);
  const date = els.invoiceDate.value || new Date().toISOString().slice(0, 10);
  const showVat = els.showVat.checked;
  return `
    <article class="invoice-document">
      <div class="invoice-title">
        <h2>Nakladnoy № ${escapeHtml(els.invoicePrefix.value || 'NK')}-${String(number).padStart(3, '0')}</h2>
        <p>${formatDate(date)}</p>
      </div>
      <div class="invoice-meta">
        <div><b>Yetkazib beruvchi:</b> ${escapeHtml(els.supplierName.value || '-')}</div>
        <div><b>Qabul qiluvchi / Магазин:</b> ${escapeHtml(store.name)}</div>
        <div><b>Asos fayl:</b> ${escapeHtml(state.fileName)}</div>
        <div><b>12% NDS:</b> ${formatMoney(total.vat)}</div>
      </div>
      <div class="table-wrap">
        <table>
          <thead>${productHeader(showVat)}</thead>
          <tbody>${store.products.map((item, index) => productRow(item, index + 1, showVat)).join('')}</tbody>
          <tfoot>${totalRow(total, showVat)}</tfoot>
        </table>
      </div>
      <div class="signatures">
        <div>Topshirdi</div>
        <div>Qabul qildi</div>
      </div>
    </article>`;
}

function productHeader(showVat) {
  return `<tr>
    <th>№</th><th>Товар / Mahsulot nomi</th><th class="num">Кол-во</th><th class="num">Цена за единицу</th><th class="num">Стоимость</th>
    ${showVat ? '<th class="num">НДС 12%</th><th class="num">Стоимость с учётом НДС</th>' : ''}
  </tr>`;
}

function productRow(item, index = '', showVat = true) {
  return `<tr>
    <td>${index}</td><td>${escapeHtml(item.name)}</td><td class="num">${formatQuantity(item.quantity)}</td><td class="num">${formatMoney(item.price)}</td><td class="num">${formatMoney(item.net)}</td>
    ${showVat ? `<td class="num">${formatMoney(item.vat)}</td><td class="num">${formatMoney(item.gross)}</td>` : ''}
  </tr>`;
}

function totalRow(total, showVat) {
  return `<tr><td colspan="2">Jami</td><td class="num">${formatQuantity(total.quantity)}</td><td></td><td class="num">${formatMoney(total.net)}</td>${showVat ? `<td class="num">${formatMoney(total.vat)}</td><td class="num">${formatMoney(total.gross)}</td>` : ''}</tr>`;
}

function exportCsv() {
  if (!state.stores.length) return;
  const lines = buildFlatExportRows();
  const csv = lines.map((line) => line.map(csvCell).join(';')).join('\n');
  downloadBlob(csv, `nakladnoy-${Date.now()}.csv`, 'text/csv;charset=utf-8');
}

async function exportExcel() {
  if (!state.stores.length || !window.XLSX) return;
  const stores = rebuildStoresForCurrentGrouping();
  const templateWorkbook = await loadTemplateWorkbook();

  if (templateWorkbook) {
    exportTemplateWorkbook(templateWorkbook, stores);
    return;
  }

  const workbook = XLSX.utils.book_new();
  const mode = els.invoiceMode.value;

  if (mode === 'single') {
    appendStoreSheet(workbook, { name: 'Umumiy nakladnoy', products: stores.flatMap((store) => store.products), orderNumbers: stores.flatMap((store) => store.orderNumbers || []) }, 'Nakladnoy');
  } else {
    stores.forEach((store, index) => appendStoreSheet(workbook, store, `${index + 1}-${store.name}`));
  }

  const summarySheet = XLSX.utils.aoa_to_sheet(buildFlatExportRows());
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Tahlil');
  XLSX.writeFile(workbook, `nakladnoy-${Date.now()}.xlsx`);
}

async function loadTemplateWorkbook() {
  for (const templateFile of TEMPLATE_FILES) {
    const response = await fetch(encodeURI(templateFile)).catch(() => null);
    if (!response || !response.ok) continue;
    const buffer = await response.arrayBuffer();
    return XLSX.read(buffer, { type: 'array', cellDates: true, cellStyles: true });
  }
  return null;
}

function exportTemplateWorkbook(templateWorkbook, stores) {
  const workbook = XLSX.utils.book_new();
  const templateSheetName = templateWorkbook.SheetNames[0];
  const templateSheet = templateWorkbook.Sheets[templateSheetName];
  const mode = els.invoiceMode.value;
  const targetStores = mode === 'single'
    ? [{ name: 'Umumiy nakladnoy', products: stores.flatMap((store) => store.products), orderNumbers: stores.flatMap((store) => store.orderNumbers || []) }]
    : stores;

  targetStores.forEach((store, index) => {
    const sheet = cloneSheet(templateSheet);
    fillTemplateSheet(sheet, store);
    XLSX.utils.book_append_sheet(workbook, sheet, safeSheetName(`${index + 1}-${store.name}`));
  });

  XLSX.writeFile(workbook, `nakladnoy-korzinka-365-${Date.now()}.xlsx`);
}

function cloneSheet(sheet) {
  return JSON.parse(JSON.stringify(sheet));
}

function fillTemplateSheet(sheet, store) {
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
  writeNearLabel(sheet, range, ['номер заказ', 'номер заказа', 'номер zakaz'], formatOrderNumbers(store));
  writeNearLabel(sheet, range, ['магазин', 'доставка', 'получатель', 'покупатель'], store.name);

  const table = findTemplateProductTable(sheet, range);
  if (!table) return;

  const productNames = collectTemplateProductNames(sheet, table);
  store.products.forEach((product, index) => {
    const row = table.startRow + index;
    const productName = resolveTemplateProductName(product.name, productNames);
    writeCell(sheet, row, table.productCol, productName);
    if (table.quantityCol !== -1) writeCell(sheet, row, table.quantityCol, roundMoney(product.quantity));
    if (table.priceCol !== -1) writeCell(sheet, row, table.priceCol, roundMoney(product.price));
    if (table.netCol !== -1) writeCell(sheet, row, table.netCol, roundMoney(product.net));
    if (table.vatCol !== -1) writeCell(sheet, row, table.vatCol, roundMoney(product.vat));
    if (table.grossCol !== -1) writeCell(sheet, row, table.grossCol, roundMoney(product.gross));
  });
}

function writeNearLabel(sheet, range, labels, value) {
  if (!value) return;
  const cellAddress = findCellAddressByLabels(sheet, range, labels);
  if (!cellAddress) return;
  const cell = XLSX.utils.decode_cell(cellAddress);
  writeCell(sheet, cell.r + 1, cell.c, value);
}

function findCellAddressByLabels(sheet, range, labels) {
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      const value = normalizeText(sheet[address]?.v);
      if (labels.some((label) => value.includes(label))) return address;
    }
  }
  return '';
}

function findTemplateProductTable(sheet, range) {
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    const headers = [];
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      headers[col] = normalizeText(sheet[XLSX.utils.encode_cell({ r: row, c: col })]?.v);
    }
    const productCol = headers.findIndex((header) => header === 'товар' || header === 'наименование' || header.includes('наименование товара') || header.includes('номенклатура'));
    const quantityCol = headers.findIndex((header) => header.includes('кол-во') || header.includes('количество'));
    if (productCol === -1 || quantityCol === -1) continue;
    return {
      startRow: row + 1,
      productCol,
      quantityCol,
      priceCol: headers.findIndex((header) => header.includes('цена')),
      netCol: headers.findIndex((header) => header === 'стоимость' || header.includes('сумма без ндс')),
      vatCol: headers.findIndex((header) => header.includes('ндс')),
      grossCol: headers.findIndex((header) => header.includes('с учетом ндс') || header.includes('с учётом ндс') || header.includes('итого')),
    };
  }
  return null;
}

function collectTemplateProductNames(sheet, table) {
  const names = [];
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
  for (let row = table.startRow; row <= range.e.r; row += 1) {
    const value = cleanName(sheet[XLSX.utils.encode_cell({ r: row, c: table.productCol })]?.v);
    if (value && !isIgnoredProductName(value)) names.push(value);
  }
  return names;
}

function resolveTemplateProductName(sourceName, templateNames) {
  const source = normalizeText(sourceName);
  const exact = templateNames.find((name) => normalizeText(name) === source);
  if (exact) return exact;
  const partial = templateNames.find((name) => {
    const target = normalizeText(name);
    return target.includes(source) || source.includes(target);
  });
  if (partial) return partial;

  const sourceTokens = productTokens(source);
  let bestName = '';
  let bestScore = 0;
  templateNames.forEach((name) => {
    const targetTokens = productTokens(normalizeText(name));
    const matches = sourceTokens.filter((token) => targetTokens.includes(token)).length;
    const score = matches / Math.max(sourceTokens.length, targetTokens.length, 1);
    if (score > bestScore) {
      bestScore = score;
      bestName = name;
    }
  });
  return bestScore >= 0.45 ? bestName : sourceName;
}

function productTokens(value) {
  return value.split(/[^a-zа-я0-9]+/i).filter((token) => token.length > 2);
}

function writeCell(sheet, row, col, value) {
  if (col < 0) return;
  const address = XLSX.utils.encode_cell({ r: row, c: col });
  const existing = sheet[address] || {};
  sheet[address] = { ...existing, v: value, t: typeof value === 'number' ? 'n' : 's' };
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
  if (row > range.e.r) range.e.r = row;
  if (col > range.e.c) range.e.c = col;
  sheet['!ref'] = XLSX.utils.encode_range(range);
}

function appendStoreSheet(workbook, store, sheetName) {
  const total = totals(store.products);
  const rows = [
    [`Nakladnoy № ${els.invoicePrefix.value || 'NK'}`, '', '', '', '', '', ''],
    ['Sana', formatDate(els.invoiceDate.value || new Date().toISOString().slice(0, 10)), '', 'Magazin', store.name, '', ''],
    ['Номер заказ', formatOrderNumbers(store), '', 'Asos fayl', state.fileName, '', ''],
    ['Yetkazib beruvchi', els.supplierName.value || '-', '', '', '', '', ''],
    [],
    ['№', 'Товар / Mahsulot nomi', 'Кол-во', 'Цена за единицу', 'Стоимость', 'НДС 12%', 'Стоимость с учётом НДС'],
    ...store.products.map((item, index) => [index + 1, item.name, roundMoney(item.quantity), roundMoney(item.price), roundMoney(item.net), roundMoney(item.vat), roundMoney(item.gross)]),
    ['Jami', '', roundMoney(total.quantity), '', roundMoney(total.net), roundMoney(total.vat), roundMoney(total.gross)],
  ];
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  worksheet['!cols'] = [{ wch: 8 }, { wch: 44 }, { wch: 12 }, { wch: 18 }, { wch: 16 }, { wch: 16 }, { wch: 24 }];
  XLSX.utils.book_append_sheet(workbook, worksheet, safeSheetName(sheetName));
}

function buildFlatExportRows() {
  const lines = [['Magazin', 'Номер заказ', 'Tovar', 'Kol-vo', 'Narx', 'Qiymat', 'NDS 12%', 'NDS bilan']];
  state.stores.forEach((store) => {
    store.products.forEach((item) => lines.push([store.name, formatOrderNumbers(store), item.name, item.quantity, item.price, item.net, item.vat, item.gross]));
  });
  return lines;
}

function safeSheetName(name) {
  return String(name || 'Sheet').replace(/[:\\/?*\[\]]/g, ' ').slice(0, 31) || 'Sheet';
}

function csvCell(value) {
  return `"${String(value ?? '').replace(/"/g, '""')}"`;
}

function downloadBlob(content, name, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
}

function setStatus(html) {
  els.statusText.innerHTML = html;
}

function showError(message) {
  setStatus(`<span class="error">${escapeHtml(message)}</span>`);
}

function roundMoney(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function formatMoney(value) {
  return new Intl.NumberFormat('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(roundMoney(value || 0));
}

function formatQuantity(value) {
  return new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 3 }).format(value || 0);
}

function formatOrderNumbers(store) {
  return Array.from(new Set(store?.orderNumbers || [])).filter(Boolean).join(', ');
}

function formatDate(value) {
  const [year, month, day] = value.split('-');
  return day && month && year ? `${day}.${month}.${year}` : value;
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>'"]/g, (char) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;',
  }[char]));
}
