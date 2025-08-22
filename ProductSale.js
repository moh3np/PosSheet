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

  function getLastDataRow(range) {
    var sheet = range.getSheet();
    var startRow = range.getRow() + 1;
    var col = range.getColumn();
    var lastRow = sheet.getLastRow();
  var numRows = lastRow - range.getRow();
  if (numRows < 1) return range.getRow();
  var values = sheet.getRange(startRow, col, numRows, 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    var val = values[i][0];
    if (val !== '' && val !== null) {
      return startRow + i;
    }
  }
  return range.getRow();
}

function getInventoryData() {
  var ss = SpreadsheetApp.getActive();
  var invRange = ss.getRangeByName('Inventory');
  if (!invRange) {
    return {names:[], skus:[], sns:[], persianSNS:[], locations:[], prices:[], uniqueCodes:[], brands:[], sellers:[]};
  }
  var sheet = invRange.getSheet();
  var lastRow = getLastDataRow(invRange);
  var numRows = lastRow - invRange.getRow();
  if (numRows < 1) {
    return {names:[], skus:[], sns:[], persianSNS:[], locations:[], prices:[], uniqueCodes:[], brands:[], sellers:[]};
  }
  var values = sheet.getRange(invRange.getRow() + 1, invRange.getColumn(), numRows, invRange.getNumColumns()).getValues();
  var names = [], brands = [], uniqueCodes = [], sns = [], sellers = [], prices = [], locations = [], skus = [], persianSns = [];
  values.forEach(function(r){
    names.push(r[0]);
    brands.push(r[1]);
    uniqueCodes.push(r[3]);
    sns.push(r[4]);
    sellers.push(r[5]);
    prices.push(r[6]);
    locations.push(r[7]);
    skus.push(r[8]);
    persianSns.push(r[9]);
  });
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
      ordersRange: 'Orders',
      inventoryRange: 'Inventory'
    }, tlItems, dateStr);
  }
  var brItems = items.filter(function(it){ return it.sku && it.sku.indexOf('BR') === 0; });
  if (brItems.length) {
    processExternalOrder({
      spreadsheetId: '12-Khe_IZ9S7z_VN_LZQCHdcKEIgKDquviar8cSR_wG8',
      ordersRange: 'StoreOrders',
      inventoryRange: 'Inventory'
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

  var ids = [], names = [], skus = [], sns = [], dates = [], prices = [], paid = [], uniqueCodes = [], locations = [], sellers = [], brands = [];
  items.forEach(function(it) {
    ids.push([orderId]);
    names.push([it.name]);
    skus.push([it.sku ? it.sku.replace(/\D/g, '') : '']);
    sns.push([it.serial]);
    dates.push([dateStr]);
    prices.push([it.price]);
    paid.push([it.paid]);
    uniqueCodes.push([it.uniqueCode]);
    locations.push([it.location]);
    sellers.push([it.seller]);
    brands.push([it.brand]);
  });
  sheet.getRange(nextRow, idCol, items.length, 1).setValues(ids);
  sheet.getRange(nextRow, nameCol, items.length, 1).setValues(names);
  sheet.getRange(nextRow, skuCol, items.length, 1).setValues(skus);
  sheet.getRange(nextRow, snCol, items.length, 1).setValues(sns);
  sheet.getRange(nextRow, dateCol, items.length, 1).setValues(dates);
  sheet.getRange(nextRow, priceCol, items.length, 1).setValues(prices);
  sheet.getRange(nextRow, paidCol, items.length, 1).setValues(paid);
  if (uniqueCodeCol != null) sheet.getRange(nextRow, uniqueCodeCol, items.length, 1).setValues(uniqueCodes);
  if (locationCol != null) sheet.getRange(nextRow, locationCol, items.length, 1).setValues(locations);
  if (sellerCol != null) sheet.getRange(nextRow, sellerCol, items.length, 1).setValues(sellers);
  if (brandCol != null) sheet.getRange(nextRow, brandCol, items.length, 1).setValues(brands);

  var invRange = ss.getRangeByName(cfg.inventoryRange);
  var invSheet = invRange ? invRange.getSheet() : null;
  var invValues = invRange ? invRange.getValues() : [];
  items.forEach(function(it) {
    if (!invRange) return;
    var targetSn = String(it.serial).trim();
    for (var i = 0; i < invValues.length; i++) {
      var sn = String(invValues[i][4]).trim();
      if (sn === targetSn) {
        invSheet.deleteRow(invRange.getRow() + i);
        invValues.splice(i, 1);
        break;
      }
    }
  });
}

function getPersianDateTime() {
  var parts = Utilities.formatDate(new Date(), 'Asia/Tehran', 'yyyy-M-d-HH:mm:ss').split('-');
  var gYear = Number(parts[0]);
  var gMonth = Number(parts[1]);
  var gDay = Number(parts[2]);
  var time = parts[3];
  var j = gregorianToJalali(gYear, gMonth, gDay);
  var jy = j[0];
  var jm = ('0' + j[1]).slice(-2);
  var jd = ('0' + j[2]).slice(-2);
  return jy + '/' + jm + '/' + jd + ' ' + time;
}

function gregorianToJalali(gy, gm, gd) {
  var g_d_m = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334];
  var jy;
  if (gy > 1600) {
    jy = 979;
    gy -= 1600;
  } else {
    jy = 0;
    gy -= 621;
  }
  var gy2 = gm > 2 ? gy + 1 : gy;
  var days = (365 * gy) + Math.floor((gy2 + 3) / 4) - Math.floor((gy2 + 99) / 100) + Math.floor((gy2 + 399) / 400) - 80 + gd + g_d_m[gm - 1];
  jy += 33 * Math.floor(days / 12053);
  days %= 12053;
  jy += 4 * Math.floor(days / 1461);
  days %= 1461;
  if (days > 365) {
    jy += Math.floor((days - 1) / 365);
    days = (days - 1) % 365;
  }
  var jm = (days < 186) ? 1 + Math.floor(days / 31) : 7 + Math.floor((days - 186) / 30);
  var jd = 1 + ((days < 186) ? (days % 31) : ((days - 186) % 30));
  return [jy, jm, jd];
}

