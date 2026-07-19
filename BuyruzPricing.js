function showBuyruzPricingDialog() {
  var html = HtmlService.createTemplateFromFile('buyruz-pricing-auth')
    .evaluate()
    .setWidth(320)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'ورود رمز دسترسی');
}

function buyruzPricingCompleteAuthorization(pinDigits) {
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
    buyruzPricingShowMainDialog_();
    return { success: true };
  }
  return {
    success: false,
    message: 'رمز وارد شده صحیح نیست.'
  };
}

function buyruzPricingShowMainDialog_() {
  var template = HtmlService.createTemplateFromFile('buyruz-pricing');
  var pricingData = getBuyruzPricingData();
  template.pricingJson = JSON.stringify(pricingData).replace(/</g, '\\u003c');
  var html = template.evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'قیمت محصولات بایروز');
}

function getBuyruzPricingData() {
  var result = {
    items: [],
    error: null
  };

  var spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(Config.sheets.buyruzSheet);
  } catch (err) {
    result.error = 'دسترسی به فایل محصولات بایروز مقدور نیست.';
    return result;
  }
  if (!spreadsheet) {
    result.error = 'فایل محصولات بایروز یافت نشد.';
    return result;
  }

  var ranges = buyruzPricingLoadProductRanges_(spreadsheet);
  if (ranges.error) {
    result.error = ranges.error;
    return result;
  }

  var factorLookup = buyruzPricingBuildFactorLookup_(spreadsheet);
  var numRows = ranges.count;
  var items = [];
  var frozenRows = ranges.sheet.getFrozenRows();
  var startRow = ranges.nameRange.getRow();
  for (var i = 0; i < numRows; i++) {
    var sheetRow = startRow + i;
    if (frozenRows && sheetRow <= frozenRows) {
      continue;
    }

    var stockValue = ranges.stockValues && ranges.stockValues[i] ? buyruzPricingParseNumber_(ranges.stockValues[i][0]) : null;
    if (stockValue == null || stockValue <= 0) {
      continue;
    }

    var name = buyruzPricingSanitizeDisplay_(ranges.nameDisplay, i);
    if (!name) {
      continue;
    }

    var skuLabel = buyruzPricingSanitizeDisplay_(ranges.skuDisplay, i);
    var skuRaw = ranges.skuValues[i][0];
    var normalizedSku = normalizeDigits(skuLabel || skuRaw);
    if (!normalizedSku) {
      continue;
    }

    var currentPriceValue = buyruzPricingParseNumber_(ranges.priceValues[i][0]);
    var currentPriceDisplay = buyruzPricingSanitizeDisplay_(ranges.priceDisplay, i);
    if (!currentPriceDisplay && currentPriceValue != null) {
      currentPriceDisplay = String(currentPriceValue);
    }

    var lastUpdatedDisplay = buyruzPricingSanitizeDisplay_(ranges.updatedDisplay, i);
    if (!lastUpdatedDisplay) {
      var rawUpdated = ranges.updatedValues[i][0];
      if (rawUpdated instanceof Date) {
        lastUpdatedDisplay = Utilities.formatDate(rawUpdated, 'Asia/Tehran', 'yyyy/MM/dd');
      } else if (rawUpdated) {
        lastUpdatedDisplay = String(rawUpdated);
      }
    }

    var factor = factorLookup[normalizedSku];
    var purchaseDateDisplay = factor ? factor.dateDisplay : '';
    var purchasePriceValue = factor ? factor.maxPrice : null;

    items.push({
      name: name,
      sku: skuLabel || (skuRaw == null ? '' : String(skuRaw)),
      normalizedSku: normalizedSku,
      tlSku: skuLabel || (skuRaw == null ? '' : String(skuRaw)),
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
    return (a.sku || '').localeCompare(b.sku || '', 'fa');
  });

  result.items = items;
  result.generatedAt = new Date().toISOString();
  return result;
}

function buyruzPricingLoadProductRanges_(spreadsheet) {
  var nameRange = spreadsheet.getRangeByName(Config.namedRanges.buyruzProductName);
  var skuRange = spreadsheet.getRangeByName(Config.namedRanges.buyruzProductSku);
  var priceRange = spreadsheet.getRangeByName(Config.namedRanges.buyruzProductsPrice);
  var updatedRange = spreadsheet.getRangeByName(Config.namedRanges.buyruzProductPriceUpdateTime);
  var stockRange = spreadsheet.getRangeByName(Config.namedRanges.buyruzProductsStock);

  if (!nameRange || !skuRange || !priceRange || !updatedRange || !stockRange) {
    return {
      error: 'یکی از رِینج‌های محصولات بایروز تعریف نشده است.'
    };
  }

  var skuValuesAll = skuRange.getValues();
  var skuDisplayAll = skuRange.getDisplayValues();
  var numRows = buyruzPricingGetDataRowCount_(skuValuesAll);
  if (numRows < 1) {
    return { error: 'هیچ داده‌ای برای محصولات بایروز یافت نشد.' };
  }

  var nameDisplay = buyruzPricingGetAlignedDisplayValues_(nameRange, numRows);
  var priceValues = buyruzPricingGetAlignedValues_(priceRange, numRows);
  var priceDisplay = buyruzPricingGetAlignedDisplayValues_(priceRange, numRows);
  var updatedValues = buyruzPricingGetAlignedValues_(updatedRange, numRows);
  var updatedDisplay = buyruzPricingGetAlignedDisplayValues_(updatedRange, numRows);
  var stockValues = buyruzPricingGetAlignedValues_(stockRange, numRows);

  return {
    count: numRows,
    nameRange: nameRange,
    sheet: nameRange.getSheet(),
    nameDisplay: nameDisplay,
    skuDisplay: skuDisplayAll.slice(0, numRows),
    skuValues: skuValuesAll.slice(0, numRows),
    priceValues: priceValues,
    priceDisplay: priceDisplay,
    updatedValues: updatedValues,
    updatedDisplay: updatedDisplay,
    stockValues: stockValues
  };
}

function buyruzPricingBuildFactorLookup_(spreadsheet) {
  var map = {};
  var skuRange = spreadsheet.getRangeByName(Config.namedRanges.buyruzFactorSku);
  var dateRange = spreadsheet.getRangeByName(Config.namedRanges.buyruzFactorPurchaseDate);
  var netPriceRange = spreadsheet.getRangeByName(Config.namedRanges.buyruzFactorNetUnitPrice);
  if (!skuRange || !dateRange || !netPriceRange) {
    return map;
  }
  var numRows = Math.min(skuRange.getNumRows(), dateRange.getNumRows(), netPriceRange.getNumRows());
  if (numRows < 1) {
    return map;
  }

  var skuValuesAll = skuRange.getValues();
  var effectiveRows = buyruzPricingGetDataRowCount_(skuValuesAll);
  if (effectiveRows < 1) {
    return map;
  }

  var dateValues = buyruzPricingGetAlignedValues_(dateRange, effectiveRows);
  var dateDisplay = buyruzPricingGetAlignedDisplayValues_(dateRange, effectiveRows);
  var netValues = buyruzPricingGetAlignedValues_(netPriceRange, effectiveRows);
  var skuValues = skuValuesAll.slice(0, effectiveRows);

  for (var i = 0; i < effectiveRows; i++) {
    var sku = normalizeDigits(skuValues[i][0]);
    if (!sku) {
      continue;
    }
    var entry = map[sku];
    if (!entry) {
      entry = { dateToken: null, dateDisplay: '', maxPrice: null };
      map[sku] = entry;
    }

    var displayDate = buyruzPricingSanitizeDisplay_(dateDisplay, i);
    var comparable = buyruzPricingMakeDateToken_(dateValues[i][0], displayDate);
    if (comparable != null) {
      if (entry.dateToken == null || comparable > entry.dateToken) {
        entry.dateToken = comparable;
        entry.dateDisplay = displayDate || buyruzPricingFormatFallbackDate_(dateValues[i][0]);
      }
    }

    var netPrice = buyruzPricingParseNumber_(netValues[i][0]);
    if (netPrice != null) {
      if (entry.maxPrice == null || netPrice > entry.maxPrice) {
        entry.maxPrice = netPrice;
      }
    }
  }

  return map;
}

function buyruzPricingMakeDateToken_(rawValue, displayValue) {
  if (displayValue) {
    var normalized = normalizeDigits(displayValue).replace(/[^\d]/g, '');
    if (normalized.length >= 8) {
      return Number(normalized.slice(0, 8));
    }
  }
  if (rawValue instanceof Date) {
    var formatted = Utilities.formatDate(rawValue, 'Asia/Tehran', 'yyyyMMdd');
    return Number(formatted);
  }
  if (rawValue) {
    var rawNormalized = normalizeDigits(String(rawValue)).replace(/[^\d]/g, '');
    if (rawNormalized.length >= 8) {
      return Number(rawNormalized.slice(0, 8));
    }
  }
  return null;
}

function buyruzPricingFormatFallbackDate_(rawValue) {
  if (rawValue instanceof Date) {
    return Utilities.formatDate(rawValue, 'Asia/Tehran', 'yyyy/MM/dd');
  }
  return rawValue == null ? '' : String(rawValue);
}

function buyruzPricingSanitizeDisplay_(twoDimensionalArray, index) {
  if (!twoDimensionalArray || twoDimensionalArray.length <= index) {
    return '';
  }
  var row = twoDimensionalArray[index];
  if (!row || !row.length) {
    return '';
  }
  var value = row[0];
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function buyruzPricingParseNumber_(value) {
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

function applyBuyruzPricingChanges(changes) {
  var prepared = buyruzPricingPrepareChanges_(changes);
  if (!prepared.length) {
    return { updated: 0, missing: [] };
  }

  var spreadsheet = SpreadsheetApp.openById(Config.sheets.buyruzSheet);
  if (!spreadsheet) {
    throw new Error('فایل مقصد برای ثبت قیمت‌های بایروز یافت نشد.');
  }

  var skuRange = spreadsheet.getRangeByName(Config.namedRanges.buyruzProductSku);
  var priceRange = spreadsheet.getRangeByName(Config.namedRanges.buyruzProductsPrice);
  var updatedRange = spreadsheet.getRangeByName(Config.namedRanges.buyruzProductPriceUpdateTime);
  if (!skuRange || !priceRange || !updatedRange) {
    throw new Error('رِینج‌های محصولات بایروز برای ثبت قیمت کامل نیست.');
  }

  var numRows = skuRange.getNumRows();
  if (priceRange.getNumRows() !== numRows || updatedRange.getNumRows() !== numRows) {
    throw new Error('تعداد ردیف ستون‌های محصولات بایروز با هم برابر نیست.');
  }
  if (numRows < 1) {
    return { updated: 0, missing: prepared.map(function(item) { return item.label; }) };
  }

  var skuValues = skuRange.getValues();
  var skuIndexMap = {};
  var dataRowCount = buyruzPricingGetDataRowCount_(skuValues);
  if (dataRowCount < 1) {
    return { updated: 0, missing: prepared.map(function(item) { return item.label; }) };
  }
  for (var i = 0; i < dataRowCount; i++) {
    var normalized = normalizeDigits(skuValues[i][0]);
    if (!normalized) {
      continue;
    }
    if (skuIndexMap[normalized] == null) {
      skuIndexMap[normalized] = i;
    }
  }

  var timestamp = typeof getPersianDate === 'function'
    ? getPersianDate()
    : Utilities.formatDate(new Date(), 'Asia/Tehran', 'yyyy/MM/dd');

  var updated = 0;
  var missing = [];
  var updatedSkuList = [];
  var priceUpdates = [];

  prepared.forEach(function(change) {
    var idx = skuIndexMap[change.sku];
    if (idx == null) {
      missing.push(change.label);
      return;
    }
    updated++;
    updatedSkuList.push(change.sku);
    priceUpdates.push({ index: idx, price: change.price });
  });

  if (updated) {
    var priceSheet = priceRange.getSheet();
    var priceStartRow = priceRange.getRow();
    var priceColumn = priceRange.getColumn();
    var updatedSheet = updatedRange.getSheet();
    var updatedStartRow = updatedRange.getRow();
    var updatedColumn = updatedRange.getColumn();
    priceUpdates.forEach(function(entry) {
      priceSheet.getRange(priceStartRow + entry.index, priceColumn).setValue(entry.price);
      updatedSheet.getRange(updatedStartRow + entry.index, updatedColumn).setValue(timestamp);
    });
    buyruzPricingMarkInventoryFlags_(spreadsheet, updatedSkuList);
  }

  return {
    updated: updated,
    missing: missing,
    timestamp: timestamp
  };
}

function buyruzPricingPrepareChanges_(changes) {
  if (!Array.isArray(changes)) {
    return [];
  }
  var result = [];
  for (var i = 0; i < changes.length; i++) {
    var item = changes[i];
    if (!item) {
      continue;
    }
    var normalizedSku = normalizeDigits(item.sku);
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

function buyruzPricingMarkInventoryFlags_(spreadsheet, skuList) {
  if (!spreadsheet || !skuList || !skuList.length) {
    return;
  }
  var normalizedSet = {};
  skuList.forEach(function(sku) {
    var normalized = normalizeDigits(sku);
    if (normalized) {
      normalizedSet[normalized] = true;
    }
  });
  var keys = Object.keys(normalizedSet);
  if (!keys.length) {
    return;
  }

  var skuRange = spreadsheet.getRangeByName(Config.namedRanges.buyruzInventorySku);
  var locationRange = spreadsheet.getRangeByName(Config.namedRanges.buyruzInventoryLocation);
  var labelRange = spreadsheet.getRangeByName(Config.namedRanges.buyruzInventoryLabelPrinted);
  if (!skuRange || !locationRange || !labelRange) {
    return;
  }

  var skuValuesAll = skuRange.getValues();
  var numRows = buyruzPricingGetDataRowCount_(skuValuesAll);
  if (numRows < 1) {
    return;
  }

  var locationValues = buyruzPricingGetAlignedValues_(locationRange, numRows);
  var labelValues = buyruzPricingGetAlignedValues_(labelRange, numRows);

  var rowsToUpdate = [];
  for (var i = 0; i < numRows; i++) {
    var normalizedSku = normalizeDigits(skuValuesAll[i][0]);
    if (!normalizedSku || !normalizedSet[normalizedSku]) {
      continue;
    }
    var locationValue = locationValues[i][0];
    var locationText = locationValue == null ? '' : String(locationValue).trim().toUpperCase();
    if (locationText === 'STORE' && labelValues[i][0] !== false) {
      rowsToUpdate.push(i);
    }
  }

  if (rowsToUpdate.length) {
    var labelSheet = labelRange.getSheet();
    var labelStartRow = labelRange.getRow();
    var labelColumn = labelRange.getColumn();
    rowsToUpdate.forEach(function(idx) {
      labelSheet.getRange(labelStartRow + idx, labelColumn).setValue(false);
    });
  }
}

function buyruzPricingGetDataRowCount_(values) {
  if (!values || !values.length) {
    return 0;
  }
  for (var i = values.length - 1; i >= 0; i--) {
    var row = values[i];
    if (!row || !row.length) {
      continue;
    }
    for (var j = 0; j < row.length; j++) {
      var cell = row[j];
      if (cell !== '' && cell !== null) {
        return i + 1;
      }
    }
  }
  return 0;
}

function buyruzPricingGetAlignedValues_(range, length) {
  return buyruzPricingReadRangeRows_(range, length, false);
}

function buyruzPricingGetAlignedDisplayValues_(range, length) {
  return buyruzPricingReadRangeRows_(range, length, true);
}

function buyruzPricingReadRangeRows_(range, length, useDisplay) {
  if (!range || length < 1) {
    return [];
  }
  var sheet = range.getSheet();
  var startRow = range.getRow();
  var maxRows = sheet.getMaxRows();
  var available = Math.max(0, maxRows - startRow + 1);
  var readLength = Math.min(length, available);
  var data = [];
  if (readLength > 0) {
    var method = useDisplay ? 'getDisplayValues' : 'getValues';
    data = sheet.getRange(startRow, range.getColumn(), readLength, range.getNumColumns())[method]();
  }
  if (data.length < length) {
    data = buyruzPricingPadRows_(data, length, range.getNumColumns());
  }
  return data;
}

function buyruzPricingPadRows_(values, targetLength, numCols) {
  var result = Array.isArray(values) ? values.slice() : [];
  var columnCount = numCols || (result.length && result[0] ? result[0].length : 1);
  while (result.length < targetLength) {
    var row = [];
    for (var i = 0; i < columnCount; i++) {
      row.push('');
    }
    result.push(row);
  }
  return result;
}
