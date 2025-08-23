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
  var lastRow = getLastDataRow(invRange);
  var numRows = lastRow - invRange.getRow();
  if (numRows < 1) return empty;
  var values = sheet.getRange(invRange.getRow() + 1, invRange.getColumn(), numRows, invRange.getNumColumns()).getValues();
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
  handleExternalOrders(dateStr, items);
}

function handleExternalOrders(dateStr, items) {
  var tlItems = items.filter(function(it){ return it.sku && it.sku.indexOf('TL') === 0; });
  if (tlItems.length) {
    processExternalOrder({
      spreadsheetId: '1LIR_q1xrpdzcqoBJmNXTO0UJ9dksoBjS7h3Me4PRB1s',
      ordersRange: 'ToylandOrders',
      inventoryRange: 'ToylandInventory'
    }, tlItems, dateStr);
  }
  var brItems = items.filter(function(it){ return it.sku && it.sku.indexOf('BR') === 0; });
  if (brItems.length) {
    processExternalOrder({
      spreadsheetId: '12-Khe_IZ9S7z_VN_LZQCHdcKEIgKDquviar8cSR_wG8',
      ordersRange: 'BuyruzPosOrders',
      inventoryRange: 'BuyruzInventory'
    }, brItems, dateStr);
  }
}

function processExternalOrder(cfg, items, dateStr) {
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

  var headerRow = ordersRange.getRow();
  var idValuesRange = sheet.getRange(headerRow, idCol, sheet.getLastRow() - headerRow + 1, 1);
  var idValues = idValuesRange.getValues().map(function(r){ return r[0]; });
  var lastId = 0;
  idValues.forEach(function(v){
    var num = parseInt(String(v).replace(/\D/g, ''), 10);
    if (!isNaN(num) && num > lastId) {
      lastId = num;
    }
  });
  var orderId = lastId + 1;
  var nextIndex = 0;
  while (nextIndex < idValues.length && idValues[nextIndex]) {
    nextIndex++;
  }
  var nextRow = headerRow + nextIndex;

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
    var startCol = invRange.getColumn();
    var dataStart = invRange.getRow() + 1;
    var numInvCols = invRange.getNumColumns();
    var dataRows = invSheet.getLastRow() - invRange.getRow();
    var invValues = dataRows > 0 ? invSheet.getRange(dataStart, startCol, dataRows, numInvCols).getValues() : [];
    var removeSet = items.map(function(it){ return String(it.serial).trim(); });
    var filtered = invValues.filter(function(r){ return removeSet.indexOf(String(r[4]).trim()) === -1; });
    if (cfg.inventoryRange === 'BuyruzInventory') {
      filtered.forEach(function(r){ r[6] = ''; });
    }
    if (dataRows > 0) {
      invSheet.getRange(dataStart, startCol, dataRows, numInvCols).clearContent();
      invSheet.getRange(dataStart, startCol + 8, dataRows, 1).clearDataValidations();
    }
    if (filtered.length) {
      var targetRange = invSheet.getRange(dataStart, startCol, filtered.length, numInvCols);
      targetRange.setValues(filtered);
      var cbRange = invSheet.getRange(dataStart, startCol + 8, filtered.length, 1);
      cbRange.insertCheckboxes();
    }
    if (cfg.inventoryRange === 'BuyruzInventory') {
      var formula = "=ARRAYFORMULA(IFERROR(XLOOKUP(VALUE(C2:C), VALUE('قیمت محصولات'!Z:Z), 'قیمت محصولات'!C:C, \"\"), \"\"))";
      invSheet.getRange(dataStart, startCol + 6).setFormula(formula);
    }
  }
}

