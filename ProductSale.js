// Removed logging and timing utilities to simplify code and avoid side effects.

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('فروش')
    .addItem('خرده فروشی', 'showSaleDialog')
    .addItem('لغو سفارش', 'showCancelDialog')
    .addToUi();
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
  var invRange = ss.getRangeByName('PosInventory');
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
  var dateStr = getPersianDateTime();
  var orderId = getNextOrderId();
  handleExternalOrders(dateStr, items, orderId);
}

function handleExternalOrders(dateStr, items, orderId) {
  var tlItems = items.filter(function(it){ return it.sku && it.sku.indexOf('TL') === 0; });
  if (tlItems.length) {
    processExternalOrder({
      spreadsheetId: '1LIR_q1xrpdzcqoBJmNXTO0UJ9dksoBjS7h3Me4PRB1s',
      ordersRange: 'ToylandOrders',
      inventoryRange: 'ToylandInventory'
    }, tlItems, dateStr, orderId);
  }
  var brItems = items.filter(function(it){ return it.sku && it.sku.indexOf('BR') === 0; });
  if (brItems.length) {
    processExternalOrder({
      spreadsheetId: '12-Khe_IZ9S7z_VN_LZQCHdcKEIgKDquviar8cSR_wG8',
      ordersRange: 'BuyruzPosOrders',
      inventoryRange: 'BuyruzInventory'
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
    var invSheet = invRange.getSheet();
    var invDataStart = getDataStartRow(invRange);
    var numInvCols = invRange.getNumColumns();
    var invLastRow = getLastDataRow(invRange);
    var invDataRows = invLastRow >= invDataStart ? invLastRow - invDataStart + 1 : 0;
    var invValues = invDataRows > 0 ? invSheet.getRange(invDataStart, invRange.getColumn(), invDataRows, numInvCols).getValues() : [];
    var removeSet = items.map(function(it){ return String(it.serial).trim(); });
    var filtered = invValues.filter(function(r){ return removeSet.indexOf(String(r[4]).trim()) === -1; });
    if (invDataRows > 0) {
      var clearRange = invSheet.getRange(invDataStart, invRange.getColumn(), invDataRows, numInvCols);
      clearRange.clearContent();
      clearRange.clearDataValidations();
    }
    if (filtered.length) {
      var targetRange = invSheet.getRange(invDataStart, invRange.getColumn(), filtered.length, numInvCols);
      targetRange.setValues(filtered);
      invSheet.getRange(invDataStart, invRange.getColumn() + 8, filtered.length, 1).insertCheckboxes();
    }
  }
}

function getNextOrderId() {
  var ss = SpreadsheetApp.getActive();
  var posOrdersRange = ss.getRangeByName('PosOrders');
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

