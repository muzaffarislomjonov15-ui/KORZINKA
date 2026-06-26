const VAT_RATE = 0.12;
const SAMPLE_FILE = 'Заказы с товарами - АСОСИЙ.xls';
const INVOICE_TEMPLATE_FILE = 'ШАБЛОН СЧФ для КОРЗИНКИ-365.xlsx';
const JSZIP_CDN_URL = 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js';
const EMBEDDED_INVOICE_TEMPLATE_BASE64 = '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';
const INVOICE_BLOCK_ROWS = 113;
const INVOICE_PRODUCTS = [
  { name: 'Бум. полотенце 365 kun 2-сл 2рул', price: 7200 },
  { name: 'САЛФЕТКА СТОЛ. 24*24 365 kun 100ШТ', price: 2975.9 },
  { name: 'Туалетная бумага 365 kun 2-слой шт', price: 1815.6 },
  { name: 'Бум. полотенце 365 kun BIG SIZE 1шт', price: 10430 },
  { name: 'Салфетки 365 kun для авто короб.70шт', price: 4062.5 },
  { name: 'САЛФЕТКИ КОРОБОЧНЫЕ 365 kun 100ШТ', price: 6205.35 },
  { name: 'САЛФЕТКИ КОРОБ 365 KUN 2СЛ 200ШТ', price: 8571.42 },
  { name: 'ТУАЛЕТНАЯ БУМАГА 365 kun 2СЛ 8ШТ', price: 16950 },
  { name: 'ТУАЛЕТНАЯ БУМАГА 365 kun 2СЛ 4ШТ', price: 10000 },
  { name: 'ТУАЛЕТНАЯ БУМАГА 365 kun 2СЛ 6ШТ', price: 12053.57 },
  { name: 'САЛФЕТКИ 365 kun 27Х27 100ШТ', price: 3426.07 },
  { name: 'Салфетки 365 kun 24*22 50шт', price: 1245.54 },
  { name: 'Салфетки 365 kun 27*27 50шт', price: 1892.86 },
  { name: 'Туалетная бумага Iz Korzinki 8шт', price: 20535.71 },
  { name: 'Туалетная бумага Iz Korzinki Aroma 8шт', price: 22321.4 },
  { name: 'Туалет.бумага Из Корзинки трехслой.8шт', price: 22321.43 },
  { name: 'Туалетная бумага Iz Korzinki 6шт', price: 14285.7 },
  { name: 'Сал.дисп 365 kun целлюлозные 20х20 180шт', price: 7544.64 },
];
const SUPPLIER_DEFAULTS = {
  name: '"NILPAK" MCHJ QK',
  address: 'Toshkent shahri, Yangihayot tumani, Janubiy Sanoat hududi, 28-uy',
  inn: '305124079.',
  vatReg: '326070033725.',
  account: '20208000900811474001.',
  bank: '01042,ТОШКЕНТ Ш., "КАПИТАЛБАНК" АТ БАНКИНИНГ СИРГАЛИ ФИЛ',
};
const BUYER_DEFAULTS = {
  name: '"ANGELESEY FOOD" MCHJ XK',
  address: 'Toshkent shaxri, Chilonzor tumani, Turob Tula ko\'chasi, 57-uy, "Besh Yog\'och" bozori hududida',
  inn: '202099756.',
  vatReg: '326060002860.',
  account: '20208000600578902001.',
  bank: '00450, ТОШКЕНТ Ш., "ТИФ МИЛЛИЙ БАНКИ" АЖ БОШ ОФИСИ',
};

const state = {
  fileName: '',
  rawRows: [],
  stores: [],
  warnings: [],
  selectedStoreKeys: new Set(),
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
  exportSelectedExcelBtn: document.querySelector('#exportSelectedExcelBtn'),
  exportExcelBtn: document.querySelector('#exportExcelBtn'),
  exportInvoiceExcelBtn: document.querySelector('#exportInvoiceExcelBtn'),
  printBtn: document.querySelector('#printBtn'),
  buildInvoiceBtn: document.querySelector('#buildInvoiceBtn'),
  invoiceDate: document.querySelector('#invoiceDate'),
  supplierName: document.querySelector('#supplierName'),
  invoicePrefix: document.querySelector('#invoicePrefix'),
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
    setStatus('Загружается пример из репозитория...');
    const response = await fetch(encodeURI(SAMPLE_FILE));
    if (!response.ok) throw new Error('Файл-пример не найден. Откройте сайт через локальный сервер.');
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
  state.selectedStoreKeys.clear();
  els.fileInput.value = '';
  setStatus('Файл ещё не загружен.');
  els.summaryList.innerHTML = '';
  els.resultsContent.className = 'empty-state';
  els.resultsContent.textContent = 'Сначала загрузите Excel-файл.';
  els.invoiceContent.className = 'invoice-preview empty-state';
  els.invoiceContent.textContent = 'Сначала загрузите Excel-файл.';
});

els.exportCsvBtn.addEventListener('click', exportCsv);
els.exportSelectedExcelBtn.addEventListener('click', exportSelectedQuantitiesExcel);
els.exportExcelBtn.addEventListener('click', exportOrdersExcel);
els.exportInvoiceExcelBtn.addEventListener('click', exportInvoiceTemplateExcel);
els.resultsContent.addEventListener('change', (event) => {
  if (!event.target.matches('.store-select input[type="checkbox"]')) return;
  const key = event.target.dataset.storeKey;
  if (!key) return;
  if (event.target.checked) state.selectedStoreKeys.add(key);
  else state.selectedStoreKeys.delete(key);
});
els.printBtn.addEventListener('click', () => {
  activateTab('invoice');
  window.print();
});
els.buildInvoiceBtn.addEventListener('click', renderInvoice);
[els.supplierName, els.invoicePrefix, els.invoiceDate, els.showVat].forEach((el) => {
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
  reader.onerror = () => showError('Ошибка чтения файла.');
  setStatus(`${file.name} читается...`);
  reader.readAsArrayBuffer(file);
}

function processWorkbook(buffer, fileName) {
  if (!window.XLSX) {
    showError('Библиотека XLSX не загружена. Проверьте подключение к интернету или подключите xlsx-js-style локально.');
    return;
  }

  try {
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: false, raw: true });
    const allRows = [];

    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true, blankrows: false });
      rows.forEach((row) => allRows.push(row));
    });

    const parsed = parseRows(allRows);
    state.fileName = fileName;
    state.rawRows = allRows;
    state.stores = parsed.stores;
    state.warnings = parsed.warnings;
    state.selectedStoreKeys.clear();

    renderSummary();
    renderResults();
    renderInvoice();
    activateTab('results');
  } catch (error) {
    showError(`Не удалось проанализировать Excel-файл: ${error.message}`);
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
        if (storeName) {
          currentStore = getStore(storeMap, storeName, currentTotal);
          addStoreOrderNumber(currentStore, findStoreOrderNumber(rows, rowIndex, dataRow.index));
        }
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
  })).sort((a, b) => a.name.localeCompare(b.name, 'ru'));

  if (!stores.length) warnings.push('Строки магазинов и товаров не найдены. Проверьте названия заголовков.');
  stores.forEach((store) => {
    if (!store.products.length) warnings.push(`${store.name} — товары не найдены.`);
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

    const orderNumber = normalizeOrderNumber(row[orderCol]);
    const declaredTotal = parseNumber(row[totalCol]);
    const store = getStore(storeMap, storeName, 0, false);
    addStoreOrderNumber(store, orderNumber);
    const pairKey = `${normalizeText(storeName)}|${orderNumber || sourceRow}`;
    if (!orderStorePairs.has(pairKey)) {
      orderStorePairs.add(pairKey);
      store.parts += 1;
      store.declaredTotal += declaredTotal || 0;
    }

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
    storeMap.set(key, { name, declaredTotal: 0, products: new Map(), parts: 0, orderNumbers: [] });
  }
  const store = storeMap.get(key);
  if (incrementParts) store.parts += 1;
  store.declaredTotal += declaredTotal || 0;
  return store;
}

function addStoreOrderNumber(store, value) {
  const orderNumber = normalizeOrderNumber(value);
  if (!orderNumber) return;
  store.orderNumbers = store.orderNumbers || [];
  if (!store.orderNumbers.includes(orderNumber)) store.orderNumbers.push(orderNumber);
}

function normalizeOrderNumber(value) {
  const raw = cleanName(value);
  if (!raw) return '';
  const digits = raw.replace(/\D/g, '');
  if (digits.length >= 10) return digits.slice(0, 10);
  return raw;
}

function formatStoreOrderNumbers(store) {
  return (store.orderNumbers || []).join(', ');
}

function findStoreOrderNumber(rows, headerIndex, dataIndex) {
  const headerRow = rows[headerIndex] || [];
  const parentRow = rows[Math.max(0, headerIndex - 1)] || [];
  const parents = fillMergedParents(parentRow, headerRow.length);
  const columns = headerRow.map((cell, columnIndex) => ({
    index: columnIndex,
    header: normalizeText(cell),
    parent: normalizeText(parents[columnIndex]),
  }));
  const orderColumn = findColumn(columns, ({ header, parent }) => (
    header === 'номер' && parent.includes('заказ')
  ));
  return orderColumn === -1 ? '' : normalizeOrderNumber((rows[dataIndex] || [])[orderColumn]);
}

function addProduct(store, product, groupByPrice) {
  const key = groupByPrice ? `${normalizeText(product.name)}|${normalizeNumber(product.price)}` : normalizeText(product.name);
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
  setStatus(`<b>${state.fileName}</b> успешно проанализирован.`);
  els.summaryList.innerHTML = [
    `Магазины: <b>${state.stores.length}</b>`,
    `Объединённые строки товаров: <b>${productCount}</b>`,
    `Общая сумма с НДС: <b>${formatMoney(grand.gross)}</b>`,
    ...state.warnings.map((warning) => `<span class="error">${escapeHtml(warning)}</span>`),
  ].map((item) => `<li>${item}</li>`).join('');
}

function renderResults() {
  if (!state.stores.length) {
    els.resultsContent.className = 'empty-state';
    els.resultsContent.textContent = 'Нет результатов анализа.';
    return;
  }

  const availableKeys = new Set(state.stores.map((store) => storeKey(store)));
  state.selectedStoreKeys.forEach((key) => {
    if (!availableKeys.has(key)) state.selectedStoreKeys.delete(key);
  });

  els.resultsContent.className = 'store-grid';
  els.resultsContent.innerHTML = state.stores.map((store) => storeTable(store)).join('');
}

function storeTable(store) {
  const total = totals(store.products);
  const key = storeKey(store);
  const checked = state.selectedStoreKeys.has(key) ? 'checked' : '';
  return `
    <article class="store-card">
      <div class="store-head">
        <label class="store-select">
          <input type="checkbox" data-store-key="${escapeHtml(key)}" ${checked}>
          <span><strong>${escapeHtml(store.name)}</strong><br><span class="badge">${store.parts} блок(ов) объединено</span></span>
        </label>
        <div class="kpis">
          <span class="kpi">Товар: ${store.products.length}</span>
          <span class="kpi">НДС: ${formatMoney(total.vat)}</span>
          <span class="kpi">Итого: ${formatMoney(total.gross)}</span>
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

function storeKey(store) {
  return normalizeText(store.name);
}

function renderInvoice() {
  if (!state.stores.length) {
    els.invoiceContent.className = 'invoice-preview empty-state';
    els.invoiceContent.textContent = 'Сначала загрузите Excel-файл.';
    return;
  }

  const stores = rebuildStoresForCurrentGrouping();
  els.invoiceContent.className = 'invoice-preview';
  els.invoiceContent.innerHTML = stores
    .map((store, index) => invoiceDocument({ ...store, products: buildInvoiceProducts(store.products) }, index + 1))
    .join('');
}

function rebuildStoresForCurrentGrouping() {
  const storeMap = new Map();
  state.stores.forEach((store) => {
    const target = getStore(storeMap, store.name, store.declaredTotal);
    (store.orderNumbers || []).forEach((orderNumber) => addStoreOrderNumber(target, orderNumber));
    store.products.forEach((product) => addProduct(target, product, els.groupByPrice.checked));
  });
  return Array.from(storeMap.values()).map((store) => ({ ...store, products: Array.from(store.products.values()) }));
}

function invoiceDocument(store, number) {
  const total = totals(store.products);
  const date = els.invoiceDate.value || new Date().toISOString().slice(0, 10);
  const showVat = els.showVat.checked;
  return `
    <article class="invoice-document">
      <div class="invoice-title">
        <h2>Счёт-фактура №</h2>
        <p>${formatDate(date)}</p>
      </div>
      <div class="invoice-meta">
        <div><b>Поставщик:</b> ${escapeHtml(els.supplierName.value || '-')}</div>
        <div><b>Покупатель / Магазин:</b> ${escapeHtml(store.name)}</div>
        <div><b>Исходный файл:</b> ${escapeHtml(state.fileName)}</div>
        <div><b>12% НДС:</b> ${formatMoney(total.vat)}</div>
      </div>
      <div class="table-wrap">
        <table>
          <thead>${productHeader(showVat)}</thead>
          <tbody>${store.products.map((item, index) => productRow(item, index + 1, showVat)).join('')}</tbody>
          <tfoot>${totalRow(total, showVat)}</tfoot>
        </table>
      </div>
      <div class="signatures">
        <div>Отпустил</div>
        <div>Получил</div>
      </div>
    </article>`;
}

function productHeader(showVat) {
  return `<tr>
    <th>№</th><th>Наименование товаров</th><th class="num">Кол-во</th><th class="num">Цена за единицу</th><th class="num">Стоимость</th>
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
  return `<tr><td colspan="2">Итого</td><td class="num">${formatQuantity(total.quantity)}</td><td></td><td class="num">${formatMoney(total.net)}</td>${showVat ? `<td class="num">${formatMoney(total.vat)}</td><td class="num">${formatMoney(total.gross)}</td>` : ''}</tr>`;
}

function exportCsv() {
  if (!state.stores.length) {
    showError('Сначала загрузите и разберите файл заказов.');
    return;
  }
  const lines = buildFlatExportRows();
  const csv = lines.map((line) => line.map(csvCell).join(';')).join('\n');
  downloadBlob(csv, `nakladnoy-${Date.now()}.csv`, 'text/csv;charset=utf-8');
}

function exportOrdersExcel() {
  if (!state.stores.length) {
    showError('Сначала загрузите и разберите файл заказов.');
    return;
  }
  if (!window.XLSX) {
    showError('Библиотека XLSX не загружена. Проверьте подключение к интернету.');
    return;
  }

  try {
    const stores = rebuildStoresForCurrentGrouping();
    const workbook = XLSX.utils.book_new();
    const summaryRows = buildOrdersSummaryRows(stores);
    const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
    applyOrdersSheetLayout(summarySheet, summaryRows);
    XLSX.utils.book_append_sheet(workbook, summarySheet, 'Все накладные');

    stores.forEach((store) => {
      const rows = buildStoreExportRows(store);
      const sheet = XLSX.utils.aoa_to_sheet(rows);
      applyOrdersSheetLayout(sheet, rows);
      XLSX.utils.book_append_sheet(workbook, sheet, uniqueSheetName(workbook, safeSheetName(store.name)));
    });

    XLSX.writeFile(workbook, `nakladnoy-${Date.now()}.xlsx`);
    setStatus(`<b>${state.fileName}</b>: ${stores.length} магазинов выгружено в Excel.`);
  } catch (error) {
    showError(`Не удалось выгрузить накладные в Excel: ${error.message}`);
  }
}

function exportSelectedQuantitiesExcel() {
  if (!state.stores.length) {
    showError('Сначала загрузите и разберите файл заказов.');
    return;
  }
  if (!state.selectedStoreKeys.size) {
    showError('Сначала отметьте магазины во вкладке результатов.');
    return;
  }
  if (!window.XLSX) {
    showError('Библиотека XLSX не загружена. Проверьте подключение к интернету.');
    return;
  }

  try {
    const selectedStores = rebuildStoresForCurrentGrouping()
      .filter((store) => state.selectedStoreKeys.has(storeKey(store)));
    if (!selectedStores.length) {
      showError('Отмеченные магазины не найдены. Проверьте выбранные галочки.');
      return;
    }

    const workbook = XLSX.utils.book_new();
    const rows = buildSelectedQuantityRows(selectedStores);
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    applySelectedQuantitySheetLayout(sheet, rows);
    XLSX.utils.book_append_sheet(workbook, sheet, 'Танланган жами');
    XLSX.writeFile(workbook, `tanlangan-jami-${Date.now()}.xlsx`);
    setStatus(`${selectedStores.length} отмеченных магазинов выгружено в Excel без цен.`);
  } catch (error) {
    showError(`Не удалось выгрузить выбранные магазины в Excel: ${error.message}`);
  }
}

async function exportInvoiceTemplateExcel() {
  if (!state.stores.length) {
    showError('Сначала загрузите и разберите файл заказов.');
    return;
  }
  if (!window.XLSX) {
    showError('Библиотека XLSX не загружена. Проверьте подключение к интернету.');
    return;
  }
  const buttons = [els.exportExcelBtn, els.exportInvoiceExcelBtn];
  buttons.forEach((button) => { button.disabled = true; });

  try {
    setStatus('Шаблон загружается, магазины раскладываются по отдельным листам...');
    const templateBuffer = await loadInvoiceTemplateBuffer();
    const stores = rebuildStoresForCurrentGrouping().map((store) => ({
      ...store,
      products: buildInvoiceProducts(store.products),
    }));
    const output = await buildCombinedInvoicePackage(templateBuffer, stores);
    downloadBlob(
      output,
      `schf-korzinka-${Date.now()}.xlsx`,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    setStatus(`<b>${state.fileName}</b>: ${stores.length} магазинов выгружено по отдельным листам Excel.`);
  } catch (error) {
    showError(`Не удалось создать Excel по шаблону: ${error.message}`);
  } finally {
    buttons.forEach((button) => { button.disabled = false; });
  }
}


async function loadInvoiceTemplateBuffer() {
  try {
    const response = await fetch(encodeURI(INVOICE_TEMPLATE_FILE));
    if (response.ok) return response.arrayBuffer();
  } catch (error) {
    console.warn('Не удалось загрузить файл шаблона из репозитория, используется встроенная копия:', error);
  }

  return base64ToArrayBuffer(EMBEDDED_INVOICE_TEMPLATE_BASE64);
}

function base64ToArrayBuffer(value) {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes.buffer;
}

function buildInvoiceProducts(products) {
  const productTotals = new Map();
  products.forEach((product) => {
    const key = normalizeProductKey(product.name);
    const quantity = normalizeNumber(product.quantity);
    const price = normalizeNumber(product.price);
    const net = normalizeNumber(product.net || quantity * price);
    const vat = normalizeNumber(product.vat || net * VAT_RATE);
    const gross = normalizeNumber(product.gross || net + vat);
    const current = productTotals.get(key) || { quantity: 0, net: 0, vat: 0, gross: 0 };
    current.quantity = normalizeNumber(current.quantity + quantity);
    current.net = normalizeNumber(current.net + net);
    current.vat = normalizeNumber(current.vat + vat);
    current.gross = normalizeNumber(current.gross + gross);
    productTotals.set(key, current);
  });

  return INVOICE_PRODUCTS.map((templateProduct) => {
    const total = productTotals.get(normalizeProductKey(templateProduct.name));
    if (!total || !total.quantity) {
      return {
        name: templateProduct.name,
        quantity: 0,
        price: '',
        net: 0,
        vat: 0,
        gross: 0,
      };
    }

    return {
      name: templateProduct.name,
      quantity: total.quantity,
      price: normalizeNumber(total.net / total.quantity),
      net: total.net,
      vat: total.vat,
      gross: total.gross,
    };
  });
}

function normalizeProductKey(value) {
  return normalizeText(value)
    .replace(/[×х]/g, 'x')
    .replace(/[^a-zа-я0-9]+/gi, '');
}

async function buildCombinedInvoicePackage(templateBuffer, stores) {
  if (!window.JSZip) {
    try {
      await loadExternalScript(JSZIP_CDN_URL);
    } catch (error) {
      console.warn('JSZip не загрузился, используется резервная XLSX-выгрузка:', error);
    }
  }

  if (window.JSZip) {
    try {
      return await buildCombinedInvoicePackageWithZip(templateBuffer, stores);
    } catch (error) {
      console.warn('Не удалось собрать шаблон через JSZip, используется резервная XLSX-выгрузка:', error);
    }
  }

  return buildCombinedInvoicePackageWithXlsx(templateBuffer, stores);
}

async function buildCombinedInvoicePackageWithZip(templateBuffer, stores) {
  const archive = await JSZip.loadAsync(templateBuffer);
  const sheetFile = findArchiveFile(archive, 'xl/worksheets/sheet1.xml');
  if (!sheetFile) throw new Error('В шаблоне не найден первый лист.');

  const parser = new DOMParser();
  const serializer = new XMLSerializer();
  const worksheetXml = await sheetFile.async('string');
  const sharedStrings = await loadArchiveSharedStrings(archive, parser);
  const templateSheetRelsFile = findArchiveFile(archive, 'xl/worksheets/_rels/sheet1.xml.rels');
  const templateSheetRelsXml = templateSheetRelsFile ? await templateSheetRelsFile.async('string') : '';

  for (const fileName of Object.keys(archive.files)) {
    const normalized = fileName.replace(/\\/g, '/').toLowerCase();
    if (normalized.startsWith('xl/worksheets/sheet') || normalized.startsWith('xl/worksheets/_rels/sheet')) {
      archive.remove(fileName);
    }
  }

  stores.forEach((store, storeIndex) => {
    const documentXml = parser.parseFromString(worksheetXml, 'application/xml');
    if (documentXml.getElementsByTagName('parsererror').length) {
      throw new Error('Не удалось прочитать структуру листа шаблона.');
    }

    const namespace = documentXml.documentElement.namespaceURI;
    const sheetData = documentXml.getElementsByTagNameNS(namespace, 'sheetData')[0];
    const dimension = documentXml.getElementsByTagNameNS(namespace, 'dimension')[0];
    const cellMap = mapWorksheetCells(sheetData);

    fillInvoiceXmlCopy(documentXml, cellMap, 0, store, storeIndex + 1, sharedStrings);
    fillInvoiceXmlCopy(documentXml, cellMap, 57, store, storeIndex + 1, sharedStrings);
    if (dimension) dimension.setAttribute('ref', `A1:V${INVOICE_BLOCK_ROWS}`);

    const sheetNumber = storeIndex + 1;
    archive.file(`xl/worksheets/sheet${sheetNumber}.xml`, serializer.serializeToString(documentXml));
    if (templateSheetRelsXml) {
      archive.file(`xl/worksheets/_rels/sheet${sheetNumber}.xml.rels`, templateSheetRelsXml);
    }
  });

  await updateWorkbookSheets(archive, parser, serializer, stores);
  await updateWorkbookRelationships(archive, parser, serializer, stores.length);
  await updateWorkbookContentTypes(archive, parser, serializer, stores.length);
  await setWorkbookToFullCalculation(archive, parser, serializer);
  await removeCalculationChain(archive, parser, serializer);

  return archive.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

function buildCombinedInvoicePackageWithXlsx(templateBuffer, stores) {
  const workbook = XLSX.read(templateBuffer, {
    type: 'array',
    cellDates: false,
    cellFormula: true,
    cellNF: true,
    cellStyles: true,
    sheetStubs: true,
  });
  const [sheetName] = workbook.SheetNames;
  const templateSheet = workbook.Sheets[sheetName];
  if (!templateSheet) throw new Error('В шаблоне не найден первый лист.');

  const templateRange = XLSX.utils.decode_range(templateSheet['!ref'] || `A1:V${INVOICE_BLOCK_ROWS}`);
  const templateMerges = (templateSheet['!merges'] || []).map((merge) => ({
    s: { ...merge.s },
    e: { ...merge.e },
  }));
  workbook.SheetNames = [];
  workbook.Sheets = {};

  stores.forEach((store, storeIndex) => {
    const outputSheet = cloneSheetMetadata(templateSheet);
    appendTemplateSheetBlock(templateSheet, outputSheet, templateRange, 0);
    appendTemplateSheetMerges(outputSheet, templateMerges, 0);
    fillInvoiceSheetCopy(outputSheet, 0, store, storeIndex + 1);
    fillInvoiceSheetCopy(outputSheet, 57, store, storeIndex + 1);
    outputSheet['!ref'] = `A1:V${INVOICE_BLOCK_ROWS}`;
    XLSX.utils.book_append_sheet(workbook, outputSheet, uniqueSheetName(workbook, safeSheetName(store.name)));
  });

  workbook.Workbook = workbook.Workbook || {};
  workbook.Workbook.CalcPr = {
    ...(workbook.Workbook.CalcPr || {}),
    calcMode: 'auto',
    fullCalcOnLoad: '1',
    forceFullCalc: '1',
  };

  return XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array',
    cellStyles: true,
    bookSST: true,
    compression: true,
  });
}

function cloneSheetMetadata(templateSheet) {
  const metadata = {};
  [
    '!cols',
    '!margins',
    '!pageSetup',
    '!printHeader',
    '!printFooter',
    '!protect',
    '!outline',
    '!type',
  ].forEach((key) => {
    if (templateSheet[key] !== undefined) metadata[key] = clonePlainObject(templateSheet[key]);
  });
  metadata['!merges'] = [];
  metadata['!rows'] = [];
  return metadata;
}

function appendTemplateSheetBlock(templateSheet, outputSheet, templateRange, rowOffset) {
  for (let row = templateRange.s.r; row <= templateRange.e.r; row += 1) {
    const targetRow = row + rowOffset;
    if (templateSheet['!rows']?.[row]) {
      outputSheet['!rows'][targetRow] = clonePlainObject(templateSheet['!rows'][row]);
    }

    for (let column = templateRange.s.c; column <= templateRange.e.c; column += 1) {
      const sourceAddress = XLSX.utils.encode_cell({ r: row, c: column });
      const sourceCell = templateSheet[sourceAddress];
      if (!sourceCell) continue;

      const targetAddress = XLSX.utils.encode_cell({ r: targetRow, c: column });
      const targetCell = clonePlainObject(sourceCell);
      if (targetCell.f) targetCell.f = shiftFormulaRows(targetCell.f, rowOffset);
      if (targetCell.F) targetCell.F = shiftRangeRows(targetCell.F, rowOffset);
      outputSheet[targetAddress] = targetCell;
    }
  }
}

function appendTemplateSheetMerges(outputSheet, templateMerges, rowOffset) {
  templateMerges.forEach((merge) => {
    outputSheet['!merges'].push({
      s: { r: merge.s.r + rowOffset, c: merge.s.c },
      e: { r: merge.e.r + rowOffset, c: merge.e.c },
    });
  });
}

function fillInvoiceSheetCopy(sheet, offset, store, invoiceNumber) {
  const titleRow = 1 + offset;
  const contractRow = 2 + offset;
  const infoStart = 3 + offset;
  const headerRow = 10 + offset;
  const vatSubHeaderRow = 11 + offset;
  const firstProductRow = 12 + offset;
  const totalRowIndex = 30 + offset;
  const signRow = 34 + offset;
  const stampRow = 36 + offset;
  const total = totals(store.products);
  const date = els.invoiceDate.value || new Date().toISOString().slice(0, 10);
  const invoiceNo = '';

  const monthText = formatRussianInvoiceMonth(date);
  const orderNumberText = formatStoreOrderNumbers(store);
  const orderNumberAddress = findSheetCellAddressByText(sheet, offset, 'номер заказ');
  setSheetCell(sheet, `A${titleRow}`, store.name);
  setSheetCell(sheet, `H${titleRow}`, `СЧЕТ-ФАКТУРА № ${invoiceNo}                         от           ${monthText}г.`);
  setSheetCell(sheet, `A${contractRow}`, '      к Договору № 15/365 от «28» ноября 2022г.');
  if (orderNumberText && orderNumberAddress) setSheetCell(sheet, orderNumberAddress, orderNumberText);

  setSheetCell(sheet, `A${infoStart}`, 'Поставщик:');
  setSheetCell(sheet, `C${infoStart}`, els.supplierName.value || SUPPLIER_DEFAULTS.name);
  setSheetCell(sheet, `M${infoStart}`, 'Покупатель:');
  setSheetCell(sheet, `O${infoStart}`, BUYER_DEFAULTS.name);
  setSheetCell(sheet, `A${infoStart + 1}`, 'Адрес:');
  setSheetCell(sheet, `C${infoStart + 1}`, SUPPLIER_DEFAULTS.address);
  setSheetCell(sheet, `M${infoStart + 1}`, 'Адрес:');
  setSheetCell(sheet, `O${infoStart + 1}`, BUYER_DEFAULTS.address);
  setSheetCell(sheet, `A${infoStart + 2}`, 'ИНН:');
  setSheetCell(sheet, `C${infoStart + 2}`, SUPPLIER_DEFAULTS.inn);
  setSheetCell(sheet, `M${infoStart + 2}`, 'ИНН:');
  setSheetCell(sheet, `O${infoStart + 2}`, BUYER_DEFAULTS.inn);
  setSheetCell(sheet, `A${infoStart + 3}`, 'РКП НДС:');
  setSheetCell(sheet, `C${infoStart + 3}`, SUPPLIER_DEFAULTS.vatReg);
  setSheetCell(sheet, `M${infoStart + 3}`, 'РКП НДС:');
  setSheetCell(sheet, `O${infoStart + 3}`, BUYER_DEFAULTS.vatReg);
  setSheetCell(sheet, `A${infoStart + 4}`, 'Банковский счет:');
  setSheetCell(sheet, `C${infoStart + 4}`, SUPPLIER_DEFAULTS.account);
  setSheetCell(sheet, `M${infoStart + 4}`, 'Банковский счет:');
  setSheetCell(sheet, `O${infoStart + 4}`, BUYER_DEFAULTS.account);
  setSheetCell(sheet, `A${infoStart + 5}`, 'МФО банка:');
  setSheetCell(sheet, `C${infoStart + 5}`, SUPPLIER_DEFAULTS.bank);
  setSheetCell(sheet, `M${infoStart + 5}`, 'МФО банка:');
  setSheetCell(sheet, `O${infoStart + 5}`, BUYER_DEFAULTS.bank);

  setSheetCell(sheet, `A${headerRow}`, '№');
  setSheetCell(sheet, `B${headerRow}`, 'Наименование товаров ');
  setSheetCell(sheet, `I${headerRow}`, 'Ед.');
  setSheetCell(sheet, `J${headerRow}`, 'Кол-во');
  setSheetCell(sheet, `M${headerRow}`, 'Цена');
  setSheetCell(sheet, `N${headerRow}`, 'Стоимость поставки');
  setSheetCell(sheet, `Q${headerRow}`, 'НДС');
  setSheetCell(sheet, `S${headerRow}`, 'Стоим. поставки с учетом НДС');
  setSheetCell(sheet, `Q${vatSubHeaderRow}`, 'Ставка');
  setSheetCell(sheet, `R${vatSubHeaderRow}`, 'Сумма');

  store.products.forEach((item, index) => {
    const row = firstProductRow + index;
    setSheetCell(sheet, `A${row}`, index + 1);
    setSheetCell(sheet, `B${row}`, item.name);
    setSheetCell(sheet, `I${row}`, 'пачка');
    setSheetCell(sheet, `J${row}`, item.quantity ? normalizeNumber(item.quantity) : '');
    setSheetCell(sheet, `M${row}`, item.price === '' ? '' : normalizeNumber(item.price));
    setSheetCell(sheet, `N${row}`, normalizeNumber(item.net), `M${row}*J${row}`);
    setSheetCell(sheet, `Q${row}`, 12);
    setSheetCell(sheet, `R${row}`, normalizeNumber(item.vat), `N${row}*12/100`);
    setSheetCell(sheet, `S${row}`, normalizeNumber(item.gross), `N${row}+R${row}`);
  });

  setSheetCell(sheet, `A${totalRowIndex}`, 'Итого:');
  setSheetCell(sheet, `N${totalRowIndex}`, normalizeNumber(total.net), `SUM(N${firstProductRow}:P${firstProductRow + 17})`);
  setSheetCell(sheet, `R${totalRowIndex}`, normalizeNumber(total.vat), `SUM(R${firstProductRow}:R${firstProductRow + 17})`);
  setSheetCell(sheet, `S${totalRowIndex}`, normalizeNumber(total.gross), `SUM(S${firstProductRow}:U${firstProductRow + 17})`);
  setSheetCell(sheet, `A${signRow}`, 'Товар отпустил:');
  setSheetCell(sheet, `P${signRow}`, 'Получил:');
  setSheetCell(sheet, `B${stampRow}`, 'М.П.');
}

function setSheetCell(sheet, address, value, formula = '') {
  const existing = sheet[address] || {};
  const nextCell = {};
  ['s', 'z', 'l', 'c'].forEach((key) => {
    if (existing[key] !== undefined) nextCell[key] = clonePlainObject(existing[key]);
  });

  if (value === '') {
    nextCell.t = 'z';
    delete nextCell.v;
    delete nextCell.f;
    sheet[address] = nextCell;
    return;
  }

  if (typeof value === 'number') {
    nextCell.t = 'n';
    nextCell.v = value;
    if (formula) nextCell.f = formula;
    sheet[address] = nextCell;
    return;
  }

  nextCell.t = 's';
  nextCell.v = String(value);
  delete nextCell.f;
  sheet[address] = nextCell;
}

function clonePlainObject(value) {
  if (value === undefined || value === null || typeof value !== 'object') return value;
  if (Array.isArray(value)) return value.map(clonePlainObject);
  return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, clonePlainObject(item)]));
}

function findArchiveFile(archive, path) {
  const normalized = path.replace(/\\/g, '/').toLowerCase();
  const fileName = Object.keys(archive.files).find((item) => (
    !archive.files[item].dir && item.replace(/\\/g, '/').toLowerCase().endsWith(normalized)
  ));
  return fileName ? archive.files[fileName] : null;
}


async function loadArchiveSharedStrings(archive, parser) {
  const sharedStringsFile = findArchiveFile(archive, 'xl/sharedStrings.xml');
  if (!sharedStringsFile) return [];
  const sharedStringsXml = parser.parseFromString(await sharedStringsFile.async('string'), 'application/xml');
  return Array.from(sharedStringsXml.getElementsByTagName('si')).map((item) => (
    Array.from(item.getElementsByTagName('t')).map((text) => text.textContent || '').join('')
  ));
}

function findSheetCellAddressByText(sheet, offset, needle) {
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:V113');
  const normalizedNeedle = normalizeText(needle);
  const startRow = offset;
  const endRow = Math.min(range.e.r, offset + 8);
  for (let row = startRow; row <= endRow; row += 1) {
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: column });
      const cell = sheet[address];
      if (cell && normalizeText(cell.v) === normalizedNeedle) return address;
    }
  }
  return '';
}

function findXmlCellAddressByText(cellMap, offset, needle, sharedStrings = []) {
  const normalizedNeedle = normalizeText(needle);
  for (const [address, cell] of cellMap.entries()) {
    const row = Number(String(address).replace(/\D/g, ''));
    if (row < offset + 1 || row > offset + 9) continue;
    if (normalizeText(getXmlCellText(cell, sharedStrings)) === normalizedNeedle) return address;
  }
  return '';
}

function getXmlCellText(cell, sharedStrings = []) {
  if (cell.getAttribute('t') === 's') {
    const valueNode = Array.from(cell.children).find((child) => child.localName === 'v');
    const index = valueNode ? Number(valueNode.textContent) : -1;
    return sharedStrings[index] || '';
  }
  if (cell.getAttribute('t') === 'inlineStr') {
    return Array.from(cell.getElementsByTagName('t')).map((item) => item.textContent || '').join('');
  }
  const valueNode = Array.from(cell.children).find((child) => child.localName === 'v');
  return valueNode ? valueNode.textContent || '' : '';
}

function mapWorksheetCells(sheetData) {
  const cellMap = new Map();
  Array.from(sheetData.children).forEach((row) => {
    Array.from(row.children).forEach((cell) => {
      if (cell.localName === 'c') cellMap.set(cell.getAttribute('r'), cell);
    });
  });
  return cellMap;
}

async function updateWorkbookSheets(archive, parser, serializer, stores) {
  const workbookFile = findArchiveFile(archive, 'xl/workbook.xml');
  if (!workbookFile) throw new Error('В шаблоне не найден файл workbook.xml.');

  const workbookXml = parser.parseFromString(await workbookFile.async('string'), 'application/xml');
  const namespace = workbookXml.documentElement.namespaceURI;
  const relationshipsNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const sheets = workbookXml.getElementsByTagNameNS(namespace, 'sheets')[0];
  if (!sheets) throw new Error('В шаблоне не найден список листов.');

  sheets.replaceChildren();
  const usedNames = new Set();
  stores.forEach((store, index) => {
    const sheet = workbookXml.createElementNS(namespace, 'sheet');
    const sheetName = uniqueXmlSheetName(usedNames, store.name || `Лист ${index + 1}`);
    sheet.setAttribute('name', sheetName);
    sheet.setAttribute('sheetId', String(index + 1));
    sheet.setAttributeNS(relationshipsNamespace, 'r:id', `rIdSheet${index + 1}`);
    sheets.appendChild(sheet);
  });

  archive.file(workbookFile.name, serializer.serializeToString(workbookXml));
}

async function updateWorkbookRelationships(archive, parser, serializer, sheetCount) {
  const workbookRelsFile = findArchiveFile(archive, 'xl/_rels/workbook.xml.rels');
  if (!workbookRelsFile) throw new Error('В шаблоне не найден файл workbook.xml.rels.');

  const relsXml = parser.parseFromString(await workbookRelsFile.async('string'), 'application/xml');
  const rels = relsXml.documentElement;
  Array.from(rels.getElementsByTagName('Relationship')).forEach((relationship) => {
    const type = relationship.getAttribute('Type') || '';
    const target = relationship.getAttribute('Target') || '';
    if (type.includes('/worksheet') || target.replace(/\\/g, '/').startsWith('worksheets/sheet')) {
      rels.removeChild(relationship);
    }
  });

  const namespace = rels.namespaceURI;
  for (let index = 1; index <= sheetCount; index += 1) {
    const relationship = relsXml.createElementNS(namespace, 'Relationship');
    relationship.setAttribute('Id', `rIdSheet${index}`);
    relationship.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet');
    relationship.setAttribute('Target', `worksheets/sheet${index}.xml`);
    rels.appendChild(relationship);
  }

  archive.file(workbookRelsFile.name, serializer.serializeToString(relsXml));
}

async function updateWorkbookContentTypes(archive, parser, serializer, sheetCount) {
  const contentTypesFile = findArchiveFile(archive, '[Content_Types].xml');
  if (!contentTypesFile) throw new Error('В шаблоне не найден файл [Content_Types].xml.');

  const contentTypesXml = parser.parseFromString(await contentTypesFile.async('string'), 'application/xml');
  const types = contentTypesXml.documentElement;
  Array.from(types.getElementsByTagName('Override')).forEach((override) => {
    const partName = (override.getAttribute('PartName') || '').toLowerCase();
    if (partName.startsWith('/xl/worksheets/sheet')) types.removeChild(override);
  });

  const namespace = types.namespaceURI;
  for (let index = 1; index <= sheetCount; index += 1) {
    const override = contentTypesXml.createElementNS(namespace, 'Override');
    override.setAttribute('PartName', `/xl/worksheets/sheet${index}.xml`);
    override.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
    types.appendChild(override);
  }

  archive.file(contentTypesFile.name, serializer.serializeToString(contentTypesXml));
}

function uniqueXmlSheetName(usedNames, name) {
  const baseName = safeSheetName(name);
  let candidate = baseName;
  for (let index = 2; usedNames.has(candidate.toLowerCase()); index += 1) {
    const suffix = ` ${index}`;
    candidate = safeSheetName(`${baseName.slice(0, 31 - suffix.length)}${suffix}`);
  }
  usedNames.add(candidate.toLowerCase());
  return candidate;
}

function appendTemplateXmlBlock(documentXml, sheetData, templateRows, rowOffset) {
  const cellMap = new Map();
  templateRows.forEach((templateRow) => {
    const row = templateRow.cloneNode(true);
    const sourceRowNumber = Number(row.getAttribute('r'));
    row.setAttribute('r', String(sourceRowNumber + rowOffset));

    Array.from(row.children).forEach((cell) => {
      if (cell.localName !== 'c') return;
      const shiftedAddress = shiftCellAddress(cell.getAttribute('r'), rowOffset);
      cell.setAttribute('r', shiftedAddress);
      const formula = Array.from(cell.children).find((child) => child.localName === 'f');
      if (formula && formula.textContent) formula.textContent = shiftFormulaRows(formula.textContent, rowOffset);
      cellMap.set(shiftedAddress, cell);
    });

    sheetData.appendChild(row);
  });
  return cellMap;
}

function appendTemplateXmlMerges(mergeCells, templateMerges, rowOffset) {
  if (!mergeCells) return;
  templateMerges.forEach((templateMerge) => {
    const merge = templateMerge.cloneNode(true);
    merge.setAttribute('ref', shiftRangeRows(merge.getAttribute('ref'), rowOffset));
    mergeCells.appendChild(merge);
  });
}

function shiftCellAddress(address, rowOffset) {
  return String(address).replace(/(\$?[A-Z]{1,3})(\$?)(\d+)/, (match, column, absolute, row) => (
    `${column}${absolute}${Number(row) + rowOffset}`
  ));
}

function shiftRangeRows(range, rowOffset) {
  return String(range).split(':').map((address) => shiftCellAddress(address, rowOffset)).join(':');
}

function shiftFormulaRows(formula, rowOffset) {
  return String(formula).replace(/(\$?[A-Z]{1,3})(\$?)(\d+)/g, (match, column, absolute, row) => {
    if (absolute === '$') return match;
    return `${column}${Number(row) + rowOffset}`;
  });
}

function fillInvoiceXmlCopy(documentXml, cellMap, offset, store, invoiceNumber, sharedStrings = []) {
  const titleRow = 1 + offset;
  const contractRow = 2 + offset;
  const infoStart = 3 + offset;
  const headerRow = 10 + offset;
  const vatSubHeaderRow = 11 + offset;
  const firstProductRow = 12 + offset;
  const totalRowIndex = 30 + offset;
  const signRow = 34 + offset;
  const stampRow = 36 + offset;
  const total = totals(store.products);
  const date = els.invoiceDate.value || new Date().toISOString().slice(0, 10);
  const invoiceNo = '';

  const monthText = formatRussianInvoiceMonth(date);
  const orderNumberText = formatStoreOrderNumbers(store);
  const orderNumberAddress = findXmlCellAddressByText(cellMap, offset, 'номер заказ', sharedStrings);
  setXmlCell(documentXml, cellMap, `A${titleRow}`, store.name);
  setXmlCell(documentXml, cellMap, `H${titleRow}`, `СЧЕТ-ФАКТУРА № ${invoiceNo}                         от           ${monthText}г.`);
  setXmlCell(documentXml, cellMap, `A${contractRow}`, '      к Договору № 15/365 от «28» ноября 2022г.');
  if (orderNumberText && orderNumberAddress) setXmlCell(documentXml, cellMap, orderNumberAddress, orderNumberText);

  setXmlCell(documentXml, cellMap, `A${infoStart}`, 'Поставщик:');
  setXmlCell(documentXml, cellMap, `C${infoStart}`, els.supplierName.value || SUPPLIER_DEFAULTS.name);
  setXmlCell(documentXml, cellMap, `M${infoStart}`, 'Покупатель:');
  setXmlCell(documentXml, cellMap, `O${infoStart}`, BUYER_DEFAULTS.name);
  setXmlCell(documentXml, cellMap, `A${infoStart + 1}`, 'Адрес:');
  setXmlCell(documentXml, cellMap, `C${infoStart + 1}`, SUPPLIER_DEFAULTS.address);
  setXmlCell(documentXml, cellMap, `M${infoStart + 1}`, 'Адрес:');
  setXmlCell(documentXml, cellMap, `O${infoStart + 1}`, BUYER_DEFAULTS.address);
  setXmlCell(documentXml, cellMap, `A${infoStart + 2}`, 'ИНН:');
  setXmlCell(documentXml, cellMap, `C${infoStart + 2}`, SUPPLIER_DEFAULTS.inn);
  setXmlCell(documentXml, cellMap, `M${infoStart + 2}`, 'ИНН:');
  setXmlCell(documentXml, cellMap, `O${infoStart + 2}`, BUYER_DEFAULTS.inn);
  setXmlCell(documentXml, cellMap, `A${infoStart + 3}`, 'РКП НДС:');
  setXmlCell(documentXml, cellMap, `C${infoStart + 3}`, SUPPLIER_DEFAULTS.vatReg);
  setXmlCell(documentXml, cellMap, `M${infoStart + 3}`, 'РКП НДС:');
  setXmlCell(documentXml, cellMap, `O${infoStart + 3}`, BUYER_DEFAULTS.vatReg);
  setXmlCell(documentXml, cellMap, `A${infoStart + 4}`, 'Банковский счет:');
  setXmlCell(documentXml, cellMap, `C${infoStart + 4}`, SUPPLIER_DEFAULTS.account);
  setXmlCell(documentXml, cellMap, `M${infoStart + 4}`, 'Банковский счет:');
  setXmlCell(documentXml, cellMap, `O${infoStart + 4}`, BUYER_DEFAULTS.account);
  setXmlCell(documentXml, cellMap, `A${infoStart + 5}`, 'МФО банка:');
  setXmlCell(documentXml, cellMap, `C${infoStart + 5}`, SUPPLIER_DEFAULTS.bank);
  setXmlCell(documentXml, cellMap, `M${infoStart + 5}`, 'МФО банка:');
  setXmlCell(documentXml, cellMap, `O${infoStart + 5}`, BUYER_DEFAULTS.bank);

  setXmlCell(documentXml, cellMap, `A${headerRow}`, '№');
  setXmlCell(documentXml, cellMap, `B${headerRow}`, 'Наименование товаров ');
  setXmlCell(documentXml, cellMap, `I${headerRow}`, 'Ед.');
  setXmlCell(documentXml, cellMap, `J${headerRow}`, 'Кол-во');
  setXmlCell(documentXml, cellMap, `M${headerRow}`, 'Цена');
  setXmlCell(documentXml, cellMap, `N${headerRow}`, 'Стоимость поставки');
  setXmlCell(documentXml, cellMap, `Q${headerRow}`, 'НДС');
  setXmlCell(documentXml, cellMap, `S${headerRow}`, 'Стоим. поставки с учетом НДС');
  setXmlCell(documentXml, cellMap, `Q${vatSubHeaderRow}`, 'Ставка');
  setXmlCell(documentXml, cellMap, `R${vatSubHeaderRow}`, 'Сумма');

  store.products.forEach((item, index) => {
    const row = firstProductRow + index;
    setXmlCell(documentXml, cellMap, `A${row}`, index + 1);
    setXmlCell(documentXml, cellMap, `B${row}`, item.name);
    setXmlCell(documentXml, cellMap, `I${row}`, 'пачка');
    setXmlCell(documentXml, cellMap, `J${row}`, item.quantity ? normalizeNumber(item.quantity) : '');
    setXmlCell(documentXml, cellMap, `M${row}`, item.price === '' ? '' : normalizeNumber(item.price));
    setXmlCell(documentXml, cellMap, `N${row}`, normalizeNumber(item.net), `M${row}*J${row}`);
    setXmlCell(documentXml, cellMap, `Q${row}`, 12);
    setXmlCell(documentXml, cellMap, `R${row}`, normalizeNumber(item.vat), `N${row}*12/100`);
    setXmlCell(documentXml, cellMap, `S${row}`, normalizeNumber(item.gross), `N${row}+R${row}`);
  });

  setXmlCell(documentXml, cellMap, `A${totalRowIndex}`, 'Итого:');
  setXmlCell(documentXml, cellMap, `N${totalRowIndex}`, normalizeNumber(total.net), `SUM(N${firstProductRow}:P${firstProductRow + 17})`);
  setXmlCell(documentXml, cellMap, `R${totalRowIndex}`, normalizeNumber(total.vat), `SUM(R${firstProductRow}:R${firstProductRow + 17})`);
  setXmlCell(documentXml, cellMap, `S${totalRowIndex}`, normalizeNumber(total.gross), `SUM(S${firstProductRow}:U${firstProductRow + 17})`);
  setXmlCell(documentXml, cellMap, `A${signRow}`, 'Товар отпустил:');
  setXmlCell(documentXml, cellMap, `P${signRow}`, 'Получил:');
  setXmlCell(documentXml, cellMap, `B${stampRow}`, 'М.П.');
}

function setXmlCell(documentXml, cellMap, address, value, formula = '') {
  const namespace = documentXml.documentElement.namespaceURI;
  const cell = cellMap.get(address) || createXmlCell(documentXml, cellMap, address);
  Array.from(cell.children).forEach((child) => cell.removeChild(child));
  cell.removeAttribute('t');

  if (value === '') return;

  if (typeof value === 'number') {
    if (formula) {
      const formulaNode = documentXml.createElementNS(namespace, 'f');
      formulaNode.textContent = formula;
      cell.appendChild(formulaNode);
    }
    const valueNode = documentXml.createElementNS(namespace, 'v');
    valueNode.textContent = String(value);
    cell.appendChild(valueNode);
    return;
  }

  cell.setAttribute('t', 'inlineStr');
  const inlineString = documentXml.createElementNS(namespace, 'is');
  const text = documentXml.createElementNS(namespace, 't');
  if (/^\s|\s$/.test(value)) text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
  text.textContent = String(value);
  inlineString.appendChild(text);
  cell.appendChild(inlineString);
}

function createXmlCell(documentXml, cellMap, address) {
  const namespace = documentXml.documentElement.namespaceURI;
  const [, column, rowText] = String(address).match(/^([A-Z]+)(\d+)$/) || [];
  if (!column || !rowText) throw new Error(`Некорректный адрес ячейки ${address}.`);

  const sheetData = documentXml.getElementsByTagNameNS(namespace, 'sheetData')[0];
  const rowNumber = Number(rowText);
  let row = Array.from(sheetData.children).find((item) => (
    item.localName === 'row' && Number(item.getAttribute('r')) === rowNumber
  ));

  if (!row) {
    row = documentXml.createElementNS(namespace, 'row');
    row.setAttribute('r', String(rowNumber));
    const nextRow = Array.from(sheetData.children).find((item) => (
      item.localName === 'row' && Number(item.getAttribute('r')) > rowNumber
    ));
    if (nextRow) sheetData.insertBefore(row, nextRow);
    else sheetData.appendChild(row);
  }

  const cell = documentXml.createElementNS(namespace, 'c');
  cell.setAttribute('r', address);
  const targetColumnIndex = columnToIndex(column);
  const nextCell = Array.from(row.children).find((item) => (
    item.localName === 'c' && columnToIndex(String(item.getAttribute('r')).replace(/\d+/g, '')) > targetColumnIndex
  ));
  if (nextCell) row.insertBefore(cell, nextCell);
  else row.appendChild(cell);
  cellMap.set(address, cell);
  return cell;
}

function columnToIndex(column) {
  return String(column).split('').reduce((sum, char) => (sum * 26) + char.charCodeAt(0) - 64, 0);
}

async function setWorkbookToFullCalculation(archive, parser, serializer) {
  const workbookFile = findArchiveFile(archive, 'xl/workbook.xml');
  if (!workbookFile) return;
  const workbookXml = parser.parseFromString(
    await workbookFile.async('string'),
    'application/xml',
  );
  const namespace = workbookXml.documentElement.namespaceURI;
  let calcPr = workbookXml.getElementsByTagNameNS(namespace, 'calcPr')[0];
  if (!calcPr) {
    calcPr = workbookXml.createElementNS(namespace, 'calcPr');
    workbookXml.documentElement.appendChild(calcPr);
  }
  calcPr.setAttribute('calcMode', 'auto');
  calcPr.setAttribute('calcId', '0');
  calcPr.setAttribute('fullCalcOnLoad', '1');
  calcPr.setAttribute('forceFullCalc', '1');
  archive.file(workbookFile.name, serializer.serializeToString(workbookXml));
}

async function removeCalculationChain(archive, parser, serializer) {
  const calcChainFile = findArchiveFile(archive, 'xl/calcChain.xml');
  if (calcChainFile) archive.remove(calcChainFile.name);

  const workbookRelsFile = findArchiveFile(archive, 'xl/_rels/workbook.xml.rels');
  if (workbookRelsFile) {
    const relsXml = parser.parseFromString(await workbookRelsFile.async('string'), 'application/xml');
    Array.from(relsXml.getElementsByTagName('Relationship')).forEach((relationship) => {
      const type = relationship.getAttribute('Type') || '';
      const target = relationship.getAttribute('Target') || '';
      if (type.includes('/calcChain') || target.replace(/\\/g, '/').endsWith('calcChain.xml')) {
        relationship.parentNode.removeChild(relationship);
      }
    });
    archive.file(workbookRelsFile.name, serializer.serializeToString(relsXml));
  }

  const contentTypesFile = findArchiveFile(archive, '[Content_Types].xml');
  if (contentTypesFile) {
    const contentTypesXml = parser.parseFromString(await contentTypesFile.async('string'), 'application/xml');
    Array.from(contentTypesXml.getElementsByTagName('Override')).forEach((override) => {
      if ((override.getAttribute('PartName') || '').toLowerCase() === '/xl/calcchain.xml') {
        override.parentNode.removeChild(override);
      }
    });
    archive.file(contentTypesFile.name, serializer.serializeToString(contentTypesXml));
  }
}

function formatRussianDate(value) {
  const [year, month, day] = value.split('-');
  const monthNames = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  const monthIndex = Number(month) - 1;
  if (!day || monthIndex < 0 || monthIndex > 11 || !year) return value;
  return `${Number(day)} ${monthNames[monthIndex]} ${year}`;
}

function formatRussianInvoiceMonth(value) {
  const [year, month] = value.split('-');
  const monthNames = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  const monthIndex = Number(month) - 1;
  if (monthIndex < 0 || monthIndex > 11 || !year) return value;
  return `${monthNames[monthIndex]} ${year}`;
}

function buildFlatExportRows() {
  const lines = [['Магазин', 'Товар', 'Кол-во', 'Цена', 'Стоимость', 'НДС 12%', 'Стоимость с НДС']];
  state.stores.forEach((store) => {
    store.products.forEach((item) => lines.push([store.name, item.name, item.quantity, item.price, item.net, item.vat, item.gross]));
  });
  return lines;
}

function buildSelectedQuantityRows(stores) {
  const products = new Map();
  stores.forEach((store) => {
    store.products.forEach((item) => {
      const key = normalizeText(item.name);
      const current = products.get(key) || { name: item.name, quantity: 0 };
      current.quantity = normalizeNumber(current.quantity + Number(item.quantity || 0));
      products.set(key, current);
    });
  });

  const productRows = Array.from(products.values())
    .sort((left, right) => left.name.localeCompare(right.name, 'ru'))
    .map((item, index) => [index + 1, item.name, normalizeNumber(item.quantity)]);
  const totalQuantity = productRows.reduce((sum, row) => sum + Number(row[2] || 0), 0);

  return [
    ['Отмеченные магазины', stores.map((store) => store.name).join(', ')],
    ['Количество магазинов', stores.length],
    [],
    ['№', 'Товар', 'Жами кол-во'],
    ...productRows,
    ['Итого', '', normalizeNumber(totalQuantity)],
  ];
}

function buildOrdersSummaryRows(stores) {
  const rows = [['Магазин', 'Товар', 'Кол-во', 'Цена', 'Стоимость', 'НДС 12%', 'Стоимость с НДС']];
  stores.forEach((store) => {
    store.products.forEach((item) => rows.push([
      store.name,
      item.name,
      normalizeNumber(item.quantity),
      roundMoney(item.price),
      roundMoney(item.net),
      roundMoney(item.vat),
      roundMoney(item.gross),
    ]));
    const total = totals(store.products);
    rows.push([`${store.name} — Итого`, '', normalizeNumber(total.quantity), '', roundMoney(total.net), roundMoney(total.vat), roundMoney(total.gross)]);
    rows.push([]);
  });
  return rows;
}

function buildStoreExportRows(store) {
  const total = totals(store.products);
  return [
    ['Магазин', store.name],
    ['Блоков объединено', store.parts],
    [],
    ['№', 'Товар', 'Кол-во', 'Цена', 'Стоимость', 'НДС 12%', 'Стоимость с НДС'],
    ...store.products.map((item, index) => [
      index + 1,
      item.name,
      normalizeNumber(item.quantity),
      roundMoney(item.price),
      roundMoney(item.net),
      roundMoney(item.vat),
      roundMoney(item.gross),
    ]),
    ['Итого', '', normalizeNumber(total.quantity), '', roundMoney(total.net), roundMoney(total.vat), roundMoney(total.gross)],
  ];
}

function applySelectedQuantitySheetLayout(sheet, rows) {
  const range = XLSX.utils.decode_range(sheet['!ref']);
  sheet['!cols'] = [
    { wch: 8 },
    { wch: 64 },
    { wch: 14 },
  ];
  sheet['!autofilter'] = { ref: `A4:C${Math.max(4, rows.length)}` };

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex <= range.e.c; columnIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      const cell = sheet[address];
      if (!cell) continue;
      const isHeader = rowIndex === 0 || rowIndex === 1 || rowIndex === 3;
      const isTotal = normalizeText(rows[rowIndex]?.[0]).includes('итого');
      cell.s = {
        font: { bold: isHeader || isTotal },
        alignment: { vertical: 'center', wrapText: true },
        border: {
          top: { style: 'thin', color: { rgb: 'D9E2EF' } },
          bottom: { style: 'thin', color: { rgb: 'D9E2EF' } },
          left: { style: 'thin', color: { rgb: 'D9E2EF' } },
          right: { style: 'thin', color: { rgb: 'D9E2EF' } },
        },
      };
      if (isHeader || isTotal) {
        cell.s.fill = { fgColor: { rgb: isTotal ? 'F8FAFC' : 'EAF2FF' } };
      }
      if (columnIndex === 2 && typeof cell.v === 'number') cell.z = '#,##0.00';
    }
  }
}

function applyOrdersSheetLayout(sheet, rows) {
  const range = XLSX.utils.decode_range(sheet['!ref']);
  sheet['!cols'] = [
    { wch: 28 },
    { wch: 52 },
    { wch: 12 },
    { wch: 14 },
    { wch: 16 },
    { wch: 16 },
    { wch: 18 },
  ];
  sheet['!autofilter'] = { ref: XLSX.utils.encode_range(range) };

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex <= range.e.c; columnIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      const cell = sheet[address];
      if (!cell) continue;
      const isHeader = rows[rowIndex]?.some((value) => ['Магазин', '№'].includes(value)) && rowIndex <= 3;
      const isTotal = normalizeText(rows[rowIndex]?.[0]).includes('итого');
      cell.s = {
        font: { bold: isHeader || isTotal },
        alignment: { vertical: 'center', wrapText: true },
        border: {
          top: { style: 'thin', color: { rgb: 'D9E2EF' } },
          bottom: { style: 'thin', color: { rgb: 'D9E2EF' } },
          left: { style: 'thin', color: { rgb: 'D9E2EF' } },
          right: { style: 'thin', color: { rgb: 'D9E2EF' } },
        },
      };
      if (isHeader || isTotal) {
        cell.s.fill = { fgColor: { rgb: isTotal ? 'F8FAFC' : 'EAF2FF' } };
      }
      if (columnIndex >= 2 && typeof cell.v === 'number') cell.z = '#,##0.00';
    }
  }
}

function uniqueSheetName(workbook, name) {
  const existingNames = new Set(workbook.SheetNames.map((sheetName) => sheetName.toLowerCase()));
  const baseName = safeSheetName(name);
  if (!existingNames.has(baseName.toLowerCase())) return baseName;

  for (let index = 2; index < 1000; index += 1) {
    const suffix = ` ${index}`;
    const candidate = safeSheetName(`${baseName.slice(0, 31 - suffix.length)}${suffix}`);
    if (!existingNames.has(candidate.toLowerCase())) return candidate;
  }
  return safeSheetName(`${Date.now()}`);
}

function safeSheetName(name) {
  return String(name || 'Sheet').replace(/[:\\/?*\[\]]/g, ' ').slice(0, 31) || 'Sheet';
}

function csvCell(value) {
  return `"${String(value ?? '').replace(/"/g, '""')}"`;
}


function loadExternalScript(src) {
  return new Promise((resolve, reject) => {
    const existingScript = document.querySelector(`script[src="${src}"]`);
    if (existingScript) {
      if (existingScript.dataset.loaded === 'true') {
        resolve();
        return;
      }
      if (existingScript.dataset.loadFailed === 'true') {
        reject(new Error('проверьте подключение к интернету'));
        return;
      }
      existingScript.addEventListener('load', resolve, { once: true });
      existingScript.addEventListener('error', () => reject(new Error('проверьте подключение к интернету')), { once: true });
      return;
    }

    const script = document.createElement('script');
    script.src = src;
    script.onload = () => {
      script.dataset.loaded = 'true';
      resolve();
    };
    script.onerror = () => {
      script.dataset.loadFailed = 'true';
      reject(new Error('проверьте подключение к интернету'));
    };
    document.head.appendChild(script);
  });
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
  return normalizeNumber(value, 2);
}

function normalizeNumber(value, digits = 6) {
  const number = Number(value || 0);
  if (!Number.isFinite(number)) return 0;
  return Math.round((number + Number.EPSILON) * (10 ** digits)) / (10 ** digits);
}

function formatMoney(value) {
  return new Intl.NumberFormat('ru-RU', { minimumFractionDigits: 0, maximumFractionDigits: 6 }).format(normalizeNumber(value));
}

function formatQuantity(value) {
  return new Intl.NumberFormat('ru-RU', { minimumFractionDigits: 0, maximumFractionDigits: 6 }).format(normalizeNumber(value));
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
