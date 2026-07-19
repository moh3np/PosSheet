// Removed logging and timing utilities to simplify code and avoid side effects.

function buildPosMenus() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('فروش')
    .addItem('خرده فروشی', 'showSaleDialog')
    .addItem('لغو سفارش', 'showCancelDialog')
    .addToUi();

  ui.createMenu('محصولات')
    .addItem('محصولات بی‌لیبل', 'showUnlabeledProductsDialog')
    .addItem('قیمت محصولات تویلند\u200c', 'showToylandPricingDialog')
    .addItem('قیمت محصولات بایروز', 'showBuyruzPricingDialog')
    .addItem('انبارگردانی تویلند', 'showInventoryCountingDialog')
    .addToUi();

  var analysisMenu = ui.createMenu('تحلیل فروش');
  analysisMenu.addItem('تحلیل فروش (Sidebar)', 'PosSheetInsights_showSidebar');
  analysisMenu.addItem('نمایش گسترده', 'PosSheetInsights_showDialog');
  analysisMenu.addItem('به‌روزرسانی داده‌ها', 'PosSheetInsights_refreshCache');
  analysisMenu.addItem('تنظیم زمان‌بندی', 'PosSheetInsights_setupTrigger');
  analysisMenu.addToUi();
}

function onOpen() {
  buildPosMenus();
}

function showSaleDialog() {
  var data = getInventoryData();
  var template = HtmlService.createTemplateFromFile('sale');
  template.inventoryData = data;
  var html = template.evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'فروش محصول');
}

function getInventoryData() {
  var ss = SpreadsheetApp.getActive();
  var posInventoryRangeName = Config.namedRanges.posInventory;
  var invRange = ss.getRangeByName(posInventoryRangeName);
  var empty = {names:[], skus:[], sns:[], persianSNS:[], locations:[], prices:[], uniqueCodes:[], brands:[], sellers:[]};
  if (!invRange) return empty;
  var sheet = invRange.getSheet();
  var startRow = getDataStartRow(invRange);
  var lastRow = getLastDataRow(invRange);
  var numRows = lastRow - startRow + 1;
  if (numRows < 1) return empty;
  var values = sheet.getRange(startRow, invRange.getColumn(), numRows, invRange.getNumColumns()).getValues();
  var names = [], brands = [], uniqueCodes = [], sns = [], sellers = [], prices = [], locations = [], skus = [], persianSns = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    names.push(r[0]);
    brands.push(r[1]);
    uniqueCodes.push(r[3]);
    sns.push(r[4]);
    sellers.push(r[5]);
    prices.push(r[6]);
    locations.push(r[7]);
    skus.push(r[8]);
    persianSns.push(r[9]);
  }
  return {names:names, skus:skus, sns:sns, persianSNS:persianSns, locations:locations, prices:prices, uniqueCodes:uniqueCodes, brands:brands, sellers:sellers};
}









function submitOrder(items) {
  if (!items || !items.length) {
    return;
  }
  ensureSerialAvailability(items);
  var dateStr = getPersianDateTime();
  var orderId = getNextOrderId();
  handleExternalOrders(dateStr, items, orderId);
}

function ensureSerialAvailability(items) {
  var ss = SpreadsheetApp.getActive();
  var posOrdersRange = ss.getRangeByName('PosOrders');
  if (!posOrdersRange) return;
  var sheet = posOrdersRange.getSheet();
  var serialCol = posOrdersRange.getColumn() + 3;
  var dataStart = getDataStartRow(posOrdersRange);
  var lastRow = getLastDataRow(posOrdersRange);
  var dataRows = lastRow >= dataStart ? lastRow - dataStart + 1 : 0;
  var existing = {};
  if (dataRows > 0) {
    var vals = sheet.getRange(dataStart, serialCol, dataRows, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      var s = String(vals[i][0]).trim();
      if (s) existing[s] = true;
    }
  }
  var seen = {};
  var duplicates = [];
  items.forEach(function(it) {
    var s = String(it.serial).trim();
    if (!s) return;
    if (existing[s] || seen[s]) duplicates.push(s);
    seen[s] = true;
  });
  if (duplicates.length) {
    throw new Error('سریال ' + duplicates.join(', ') + ' قبلا به فروش رفته و موجود نیست');
  }
}

function handleExternalOrders(dateStr, items, orderId) {
  var tlItems = items.filter(function(it){ return it.sku && it.sku.indexOf('TL') === 0; });
  if (tlItems.length) {
    processExternalOrder({
      spreadsheetId: Config.sheets.tolyandSheet,
      ordersRange: Config.namedRanges.toylandOrders,
      inventoryRange: Config.namedRanges.toylandInventory
    }, tlItems, dateStr, orderId);
  }
  var brItems = items.filter(function(it){ return it.sku && it.sku.indexOf('BR') === 0; });
  if (brItems.length) {
    processExternalOrder({
      spreadsheetId: Config.sheets.buyruzSheet,
      ordersRange: Config.namedRanges.buyruzOrders,
      inventoryRange: Config.namedRanges.buyruzInventory
    }, brItems, dateStr, orderId);
  }
}

function processExternalOrder(cfg, items, dateStr, orderId) {
  var ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  var ordersRange = ss.getRangeByName(cfg.ordersRange);
  if (!ordersRange) return;
  var sheet = ordersRange.getSheet();
  var baseCol = ordersRange.getColumn();
  var numCols = ordersRange.getNumColumns();
  var col = function(idx){ return numCols > idx ? baseCol + idx : null; };
  var idCol = col(0);
  var nameCol = col(1);
  var skuCol = col(2);
  var snCol = col(3);
  var dateCol = col(4);
  var priceCol = col(5);
  var paidCol = col(6);
  var locationCol = col(7);
  var sellerCol = col(8);
  var brandCol = col(9);
  var uniqueCodeCol = col(10);

  var dataStart = getDataStartRow(ordersRange);
  var lastOrderRow = getLastDataRow(ordersRange);
  var dataRows = lastOrderRow >= dataStart ? lastOrderRow - dataStart + 1 : 0;
  var idValues = dataRows > 0 ? sheet.getRange(dataStart, idCol, dataRows, 1).getValues().map(function(r){ return r[0]; }) : [];
  var nextIndex = 0;
  while (nextIndex < idValues.length && idValues[nextIndex]) {
    nextIndex++;
  }
  var nextRow = dataStart + nextIndex;

  var rows = items.map(function(it) {
    var row = new Array(numCols);
    row[idCol - baseCol] = orderId;
    row[nameCol - baseCol] = it.name;
    row[skuCol - baseCol] = it.sku ? it.sku.replace(/\D/g, '') : '';
    row[snCol - baseCol] = it.serial;
    row[dateCol - baseCol] = dateStr;
    row[priceCol - baseCol] = it.price;
    row[paidCol - baseCol] = it.paid;
    if (locationCol != null) row[locationCol - baseCol] = it.location;
    if (sellerCol != null) row[sellerCol - baseCol] = it.seller;
    if (brandCol != null) row[brandCol - baseCol] = it.brand;
    if (uniqueCodeCol != null) row[uniqueCodeCol - baseCol] = it.uniqueCode;
    return row;
  });
  sheet.getRange(nextRow, baseCol, rows.length, numCols).setValues(rows);

  var invRange = ss.getRangeByName(cfg.inventoryRange);
  if (invRange) {
    if (cfg.inventoryRange === Config.namedRanges.buyruzInventory) {
      processBuyruzInventorySale_(ss, items);
      return;
    }
    var invSheet = invRange.getSheet();
    var invDataStart = getDataStartRow(invRange);
    var numInvCols = invRange.getNumColumns();
    var invLastRow = getLastDataRow(invRange);
    var invDataRows = invLastRow >= invDataStart ? invLastRow - invDataStart + 1 : 0;
    var invValues = invDataRows > 0 ? invSheet.getRange(invDataStart, invRange.getColumn(), invDataRows, numInvCols).getValues() : [];
    var removeSet = items.map(function(it){ return String(it.serial).trim(); });
    var filtered = invValues.filter(function(r){ return removeSet.indexOf(String(r[4]).trim()) === -1; });
    var skipToylandColumn = Config.inventory.skipToylandColumn; // skip column 7 (0-based index) for Toyland inventory updates
    var isToylandInventory = cfg.inventoryRange === Config.namedRanges.toylandInventory && numInvCols > skipToylandColumn;
    if (invDataRows > 0) {
      if (isToylandInventory) {
        var leftCols = Math.min(skipToylandColumn, numInvCols);
        if (leftCols > 0) {
          var leftClearRange = invSheet.getRange(invDataStart, invRange.getColumn(), invDataRows, leftCols);
          leftClearRange.clearContent();
          leftClearRange.clearDataValidations();
        }
        var rightCols = numInvCols - (skipToylandColumn + 1);
        if (rightCols > 0) {
          var rightClearRange = invSheet.getRange(invDataStart, invRange.getColumn() + skipToylandColumn + 1, invDataRows, rightCols);
          rightClearRange.clearContent();
          rightClearRange.clearDataValidations();
        }
      } else {
        var clearRange = invSheet.getRange(invDataStart, invRange.getColumn(), invDataRows, numInvCols);
        clearRange.clearContent();
        clearRange.clearDataValidations();
      }
    }
    if (filtered.length) {
      if (isToylandInventory) {
        var leftCols = Math.min(skipToylandColumn, numInvCols);
        if (leftCols > 0) {
          var leftRange = invSheet.getRange(invDataStart, invRange.getColumn(), filtered.length, leftCols);
          var leftValues = filtered.map(function(r) { return r.slice(0, skipToylandColumn); });
          leftRange.setValues(leftValues);
        }
        var rightCols = numInvCols - (skipToylandColumn + 1);
        if (rightCols > 0) {
          var rightRange = invSheet.getRange(invDataStart, invRange.getColumn() + skipToylandColumn + 1, filtered.length, rightCols);
          var rightValues = filtered.map(function(r) { return r.slice(skipToylandColumn + 1); });
          rightRange.setValues(rightValues);
        }
      } else {
        var targetRange = invSheet.getRange(invDataStart, invRange.getColumn(), filtered.length, numInvCols);
        targetRange.setValues(filtered);
      }
      invSheet.getRange(invDataStart, invRange.getColumn() + 8, filtered.length, 1).insertCheckboxes();
    }
  }
}

function processBuyruzInventorySale_(ss, items) {
  var ranges = getBuyruzInventorySaleRanges_(ss);
  if (!ranges) return;

  var startRow = getDataStartRow(ranges.sn);
  var lastRow = getLastDataRow(ranges.sn);
  if (lastRow < startRow) return;

  var numRows = lastRow - startRow + 1;
  var removeSet = {};
  items.forEach(function(it) {
    removeSet[String(it.serial).trim()] = true;
  });

  var columns = [
    { key: 'sku', range: ranges.sku },
    { key: 'uniqueCode', range: ranges.uniqueCode },
    { key: 'sn', range: ranges.sn },
    { key: 'location', range: ranges.location },
    { key: 'labelPrinted', range: ranges.labelPrinted }
  ];
  var valuesByKey = {};
  columns.forEach(function(col) {
    valuesByKey[col.key] = readBuyruzInventoryColumn_(col.range, startRow, numRows);
  });

  var kept = {};
  columns.forEach(function(col) {
    kept[col.key] = [];
  });

  for (var i = 0; i < numRows; i++) {
    var serial = valuesByKey.sn[i] ? String(valuesByKey.sn[i][0]).trim() : '';
    if (removeSet[serial]) {
      continue;
    }
    columns.forEach(function(col) {
      kept[col.key].push(valuesByKey[col.key][i]);
    });
  }

  columns.forEach(function(col) {
    var updateRange = getBuyruzInventoryColumnRange_(col.range, startRow, numRows);
    updateRange.clearContent();
    updateRange.clearDataValidations();
    if (kept[col.key].length) {
      getBuyruzInventoryColumnRange_(col.range, startRow, kept[col.key].length).setValues(kept[col.key]);
    }
  });

  if (kept.labelPrinted.length) {
    getBuyruzInventoryColumnRange_(ranges.labelPrinted, startRow, kept.labelPrinted.length).insertCheckboxes();
  }
}

function getBuyruzInventorySaleRanges_(ss) {
  var namedRanges = Config.namedRanges;
  var ranges = {
    productName: ss.getRangeByName(namedRanges.buyruzInventoryProductName),
    productBrand: ss.getRangeByName(namedRanges.buyruzInventoryProductBrand),
    sku: ss.getRangeByName(namedRanges.buyruzInventorySku),
    uniqueCode: ss.getRangeByName(namedRanges.buyruzInventoryUniqueCode),
    sn: ss.getRangeByName(namedRanges.buyruzInventorySN),
    supplier: ss.getRangeByName(namedRanges.buyruzInventorySupplier),
    price: ss.getRangeByName(namedRanges.buyruzInventoryPrice),
    location: ss.getRangeByName(namedRanges.buyruzInventoryLocation),
    labelPrinted: ss.getRangeByName(namedRanges.buyruzInventoryLabelPrinted)
  };
  for (var key in ranges) {
    if (!ranges[key]) {
      return null;
    }
  }
  return ranges;
}

function readBuyruzInventoryColumn_(range, startRow, numRows) {
  return getBuyruzInventoryColumnRange_(range, startRow, numRows).getValues();
}

function getBuyruzInventoryColumnRange_(range, startRow, numRows) {
  return range.getSheet().getRange(startRow, range.getColumn(), numRows, 1);
}

function getNextOrderId() {
  var ss = SpreadsheetApp.getActive();
  var posOrdersRangeName = Config.namedRanges.posOrders;
  var posOrdersRange = ss.getRangeByName(posOrdersRangeName);
  if (!posOrdersRange) return 1;
  var sheet = posOrdersRange.getSheet();
  var idCol = posOrdersRange.getColumn();
  var dataStart = getDataStartRow(posOrdersRange);
  var lastRow = getLastDataRow(posOrdersRange);
  var dataRows = lastRow >= dataStart ? lastRow - dataStart + 1 : 0;
  if (dataRows < 1) return 1;
  var values = sheet.getRange(dataStart, idCol, dataRows, 1).getValues();
  var lastId = 0;
  values.forEach(function(r) {
    var num = parseInt(String(r[0]).replace(/\D/g, ''), 10);
    if (!isNaN(num) && num > lastId) {
      lastId = num;
    }
  });
  return lastId + 1;
}
