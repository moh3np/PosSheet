function showToylandPricingDialog() {
  var html = HtmlService.createTemplateFromFile('toyland-pricing-auth')
    .evaluate()
    .setWidth(320)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'ورود رمز دسترسی');
}

function toylandPricingCompleteAuthorization(pinDigits) {
  var ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (err) {
    throw new Error('امکان دسترسی به رابط کاربری برای ورود رمز وجود ندارد.');
  }
  if (!ui) {
    throw new Error('امکان دسترسی به رابط کاربری برای ورود رمز وجود ندارد.');
  }

  var expected = '1350';
  var joined = Array.isArray(pinDigits) ? pinDigits.join('') : pinDigits;
  var input = normalizeDigits(joined);
  if (input === expected) {
    toylandPricingShowMainDialog_();
    return { success: true };
  }
  return {
    success: false,
    message: 'رمز وارد شده صحیح نیست.'
  };
}

function toylandPricingShowMainDialog_() {
  var template = HtmlService.createTemplateFromFile('toyland-pricing');
  var pricingData = getToylandPricingData();
  template.pricingJson = JSON.stringify(pricingData).replace(/</g, '\\u003c');
  var html = template.evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'قیمت محصولات تویلند');
}

function getToylandPricingData() {
  var result = {
    items: [],
    error: null
  };

  var sourceSpreadsheet;
  try {
    sourceSpreadsheet = SpreadsheetApp.openById(Config.sheets.tolyandSheet);
  } catch (err) {
    result.error = 'دسترسی به فایل محصولات تویلند مقدور نیست.';
    return result;
  }
  if (!sourceSpreadsheet) {
    result.error = 'فایل محصولات تویلند یافت نشد.';
    return result;
  }

  var productsRangeName = Config.namedRanges.toylandProducts;
  var productsRange = sourceSpreadsheet.getRangeByName(productsRangeName);
  if (!productsRange) {
    result.error = 'رِینج ' + productsRangeName + ' تعریف نشده است.';
    return result;
  }

  var numRows = productsRange.getNumRows();
  if (numRows < 1) {
    result.generatedAt = new Date().toISOString();
    return result;
  }

  var numCols = productsRange.getNumColumns();
  var requiredIndex = 9; // zero-based index we rely on
  if (numCols <= requiredIndex) {
    result.error = 'ستون‌های موردنیاز در رِینج ' + productsRangeName + ' موجود نیست.';
    return result;
  }

  var values = productsRange.getValues();
  var displayValues = productsRange.getDisplayValues();
  var timezone = 'Asia/Tehran';
  var items = [];
  var sheet = productsRange.getSheet();
  var frozenRows = sheet ? sheet.getFrozenRows() : 0;
  var startRow = productsRange.getRow();

  for (var i = 0; i < numRows; i++) {
    var sheetRow = startRow + i;
    if (frozenRows && sheetRow <= frozenRows) {
      continue;
    }

    var rawRow = values[i];
    var displayRow = displayValues[i];

    var skuDisplay = toylandPricingSanitizeField(displayRow, 0);
    var skuRaw = rawRow[0];
    var skuLabel = skuDisplay || (skuRaw == null ? '' : String(skuRaw));
    var normalizedSku = normalizeDigits(skuLabel || skuRaw);
    if (!normalizedSku) {
      continue;
    }

    var tlSku = toylandPricingSanitizeField(displayRow, 1);

    var name = toylandPricingSanitizeField(displayRow, 2);
    if (!name) {
      continue;
    }

    var currentPriceValue = toylandPricingParseNumber_(rawRow[9]);
    var currentPriceDisplay = toylandPricingSanitizeField(displayRow, 9);
    if (!currentPriceDisplay && currentPriceValue != null) {
      currentPriceDisplay = String(currentPriceValue);
    }

    var purchasePriceValue = toylandPricingParseNumber_(rawRow[4]);

    var purchaseDateDisplay = toylandPricingSanitizeField(displayRow, 6);
    if (!purchaseDateDisplay) {
      var purchaseDateRaw = rawRow[6];
      if (purchaseDateRaw instanceof Date) {
        purchaseDateDisplay = Utilities.formatDate(purchaseDateRaw, timezone, 'yyyy/MM/dd');
      }
    }

    var lastUpdatedDisplay = toylandPricingSanitizeField(displayRow, 8);

    items.push({
      name: name,
      sku: skuLabel,
      normalizedSku: normalizedSku,
      tlSku: tlSku || skuLabel,
      currentPriceDisplay: currentPriceDisplay,
      currentPriceValue: currentPriceValue,
      latestDate: purchaseDateDisplay,
      purchasePriceValue: purchasePriceValue,
      lastUpdatedValue: lastUpdatedDisplay,
      lastUpdatedDisplay: lastUpdatedDisplay
    });
  }

  items.sort(function(a, b) {
    var nameA = a.name || '';
    var nameB = b.name || '';
    var cmp = nameA.localeCompare(nameB, 'fa');
    if (cmp !== 0) {
      return cmp;
    }
    return (a.sku || '').localeCompare(b.sku || '', 'en');
  });

  result.items = items;
  result.generatedAt = new Date().toISOString();
  return result;
}

function toylandPricingSanitizeField(row, index) {
  if (!row || row.length <= index) {
    return '';
  }
  var value = row[index];
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function toylandPricingParseNumber_(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  if (typeof value === 'number') {
    return isFinite(value) ? value : null;
  }
  var normalized = normalizeDigits(value);
  if (!normalized) {
    return null;
  }
  var numeric = Number(normalized.replace(/[^\d.]/g, ''));
  return numeric === numeric ? numeric : null;
}

function applyToylandPricingChanges(changes) {
  var prepared = toylandPricingPrepareChanges_(changes);
  if (!prepared.length) {
    return { updated: 0, missing: [] };
  }

  var destSpreadsheet = SpreadsheetApp.openById(Config.sheets.tolyandSheet);
  if (!destSpreadsheet) {
    throw new Error('فایل مقصد برای ثبت قیمت‌ها یافت نشد.');
  }

  var productsRangeName = Config.namedRanges.toylandProducts;
  var productsRange = destSpreadsheet.getRangeByName(productsRangeName);
  if (!productsRange) {
    throw new Error('رِینج ' + productsRangeName + ' در فایل مقصد تعریف نشده است.');
  }

  var numRows = productsRange.getNumRows();
  if (numRows < 1) {
    return { updated: 0, missing: prepared.map(function(item) { return item.label; }) };
  }

  var skuValues = productsRange.offset(0, 0, numRows, 1).getValues();
  var skuIndexMap = {};
  for (var i = 0; i < skuValues.length; i++) {
    var normalizedSku = normalizeDigits(String(skuValues[i][0] == null ? '' : skuValues[i][0]));
    if (!normalizedSku) {
      continue;
    }
    if (skuIndexMap[normalizedSku] == null) {
      skuIndexMap[normalizedSku] = i;
    }
  }

  var priceColumnOffset = 7; // eighth column inside the Products range
  var timestampColumnOffset = 8; // ninth column inside the Products range
  var numCols = productsRange.getNumColumns();
  if (numCols <= priceColumnOffset || numCols <= timestampColumnOffset) {
    throw new Error('ستون‌های قیمت یا زمان تغییر در رِینج ' + productsRangeName + ' موجود نیست.');
  }

  var priceValues = productsRange.offset(0, priceColumnOffset, numRows, 1).getValues();
  var timestampValues = productsRange.offset(0, timestampColumnOffset, numRows, 1).getValues();

  var timestamp = typeof getPersianDate === 'function'
    ? getPersianDate()
    : (typeof getPersianDateTime === 'function'
      ? String(getPersianDateTime()).split(' ')[0]
      : Utilities.formatDate(new Date(), 'Asia/Tehran', 'yyyy/MM/dd'));

  var updated = 0;
  var missing = [];
  var updatedSkuList = [];
  prepared.forEach(function(change) {
    var idx = skuIndexMap[change.sku];
    if (idx == null) {
      missing.push(change.label);
      return;
    }
    priceValues[idx][0] = change.price;
    timestampValues[idx][0] = timestamp;
    updated++;
    updatedSkuList.push(change.sku);
  });

  if (updated) {
    productsRange.offset(0, priceColumnOffset, numRows, 1).setValues(priceValues);
    productsRange.offset(0, timestampColumnOffset, numRows, 1).setValues(timestampValues);
    toylandPricingMarkInventoryFlags_(destSpreadsheet, updatedSkuList);
  }

  return {
    updated: updated,
    missing: missing,
    timestamp: timestamp
  };
}

function toylandPricingPrepareChanges_(changes) {
  if (!Array.isArray(changes)) {
    return [];
  }
  var result = [];
  for (var i = 0; i < changes.length; i++) {
    var item = changes[i];
    if (!item) {
      continue;
    }
    var normalizedSku = normalizeDigits(String(item.sku == null ? '' : item.sku));
    if (!normalizedSku) {
      continue;
    }
    var priceValue = Number(item.price);
    if (!isFinite(priceValue) || priceValue <= 0) {
      continue;
    }
    result.push({
      sku: normalizedSku,
      price: Math.round(priceValue),
      label: item.displaySku ? String(item.displaySku) : (item.name ? String(item.name) : normalizedSku)
    });
  }
  return result;
}

function toylandPricingMarkInventoryFlags_(spreadsheet, skuList) {
  if (!spreadsheet || !skuList || !skuList.length) {
    return;
  }
  var normalizedSet = {};
  skuList.forEach(function(sku) {
    var normalized = normalizeDigits(String(sku == null ? '' : sku));
    if (normalized) {
      normalizedSet[normalized] = true;
    }
  });
  var keys = Object.keys(normalizedSet);
  if (!keys.length) {
    return;
  }
  var toylandInventoryRangeName = Config.namedRanges.toylandInventory;
  var range = spreadsheet.getRangeByName(toylandInventoryRangeName);
  if (!range) {
    return;
  }
  var numRows = range.getNumRows();
  if (numRows < 1) {
    return;
  }
  var numCols = range.getNumColumns();
  var skuCol = 2; // third column (0-based index)
  var flagCol = 8; // ninth column (0-based index)
  if (numCols <= skuCol || numCols <= flagCol) {
    return;
  }
  var values = range.getValues();
  var flagColumnValues = range.offset(0, flagCol, numRows, 1).getValues();
  var changed = false;
  for (var i = 0; i < numRows; i++) {
    var skuRaw = values[i][skuCol];
    var normalizedSku = normalizeDigits(String(skuRaw == null ? '' : skuRaw));
    if (!normalizedSku) {
      continue;
    }
    if (normalizedSet[normalizedSku]) {
      if (flagColumnValues[i][0] !== false) {
        flagColumnValues[i][0] = false;
        changed = true;
      }
    }
  }
  if (changed) {
    range.offset(0, flagCol, numRows, 1).setValues(flagColumnValues);
  }
}
