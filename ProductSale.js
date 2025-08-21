var global = this;

function startTimer(name){
  var start = Date.now();
  return function(){
    console.log(name + ' took ' + (Date.now() - start) + 'ms');
  };
}

function addTiming(names){
  names.forEach(function(name){
    var fn = global[name];
    if (typeof fn === 'function'){
      global[name] = function(){
        var end = startTimer(name);
        try {
          return fn.apply(this, arguments);
        } finally {
          end();
        }
      };
    }
  });
}

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

function getColumnValues(rangeName, sheet, lastRow) {
  var range = SpreadsheetApp.getActive().getRangeByName(rangeName);
  if (!range) return [];
  sheet = sheet || range.getSheet();
  if (lastRow === undefined) {
    lastRow = sheet.getLastRow();
  }
  var startRow = range.getRow() + 1;
  var col = range.getColumn();
  if (lastRow < startRow) return [];
  return sheet
    .getRange(startRow, col, lastRow - startRow + 1, 1)
    .getValues()
    .map(function(r){return r[0];});
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
  var snRange = ss.getRangeByName('InventorySN');
  if (!snRange) {
    return {names:[], skus:[], sns:[], persianSNS:[], locations:[], prices:[], uniqueCodes:[], brands:[], sellers:[]};
  }
  var sheet = snRange.getSheet();
  var lastRow = getLastDataRow(snRange);
  var names = getColumnValues('InventoryName', sheet, lastRow);
  var sns = getColumnValues('InventorySN', sheet, lastRow);
  var persianSns = getColumnValues('InventoryPersianSN', sheet, lastRow);
  var skus = getColumnValues('InventorySKU', sheet, lastRow);
  var locations = getColumnValues('InventoryLocation', sheet, lastRow);
  var prices = getColumnValues('InventoryPrice', sheet, lastRow);
  var uniqueCodes = getColumnValues('InventoryUniqueCode', sheet, lastRow);
  var brands = getColumnValues('InventoryBrand', sheet, lastRow);
  var sellers = getColumnValues('InventorySeller', sheet, lastRow);
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
      rangeNames: {
        id: 'OrderID',
        name: 'OrderName',
        sku: 'OrderSKU',
        sn: 'OrderSN',
        date: 'OrderDate',
        price: 'OrderPrice',
        paid: 'OrderPaidPrice',
        uniqueCode: 'OrderUniqueCode',
        location: 'OrderLocation',
        seller: 'OrderSeller',
        brand: 'OrderBrand'
      },
      inventoryRange: 'InventorySN'
    }, tlItems, dateStr);
  }
  var brItems = items.filter(function(it){ return it.sku && it.sku.indexOf('BR') === 0; });
  if (brItems.length) {
    processExternalOrder({
      spreadsheetId: '12-Khe_IZ9S7z_VN_LZQCHdcKEIgKDquviar8cSR_wG8',
      rangeNames: {
        id: 'StoreOrderID',
        name: 'StoreOrderName',
        sku: 'StoreOrderSKU',
        sn: 'StoreOrderSN',
        date: 'StoreOrderDate',
        price: 'StoreOrderPrice',
        paid: 'StoreOrderPaidPrice',
        uniqueCode: 'StoreOrderUniqueCode',
        location: 'StoreOrderLocation',
        seller: 'StoreOrderSeller',
        brand: 'StoreOrderBrand'
      },
      inventoryRange: 'InventorySN'
    }, brItems, dateStr);
  }
}

function processExternalOrder(cfg, items, dateStr) {
  var ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  var idRange = ss.getRangeByName(cfg.rangeNames.id);
  if (!idRange) return;
  var sheet = idRange.getSheet();
  var idCol = idRange.getColumn();
  var nameCol = ss.getRangeByName(cfg.rangeNames.name).getColumn();
  var skuCol = ss.getRangeByName(cfg.rangeNames.sku).getColumn();
  var snCol = ss.getRangeByName(cfg.rangeNames.sn).getColumn();
  var dateCol = ss.getRangeByName(cfg.rangeNames.date).getColumn();
  var priceCol = ss.getRangeByName(cfg.rangeNames.price).getColumn();
  var paidCol = ss.getRangeByName(cfg.rangeNames.paid).getColumn();
  var uniqueCodeRange = ss.getRangeByName(cfg.rangeNames.uniqueCode);
  var locationRange = ss.getRangeByName(cfg.rangeNames.location);
  var sellerRange = ss.getRangeByName(cfg.rangeNames.seller);
  var brandRange = ss.getRangeByName(cfg.rangeNames.brand);
  var uniqueCodeCol = uniqueCodeRange ? uniqueCodeRange.getColumn() : null;
  var locationCol = locationRange ? locationRange.getColumn() : null;
  var sellerCol = sellerRange ? sellerRange.getColumn() : null;
  var brandCol = brandRange ? brandRange.getColumn() : null;

  var idValuesRange = sheet.getRange(idRange.getRow(), idCol, sheet.getLastRow() - idRange.getRow() + 1, 1);
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
  var nextRow = idRange.getRow() + nextIndex;

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
  if (uniqueCodeCol) sheet.getRange(nextRow, uniqueCodeCol, items.length, 1).setValues(uniqueCodes);
  if (locationCol) sheet.getRange(nextRow, locationCol, items.length, 1).setValues(locations);
  if (sellerCol) sheet.getRange(nextRow, sellerCol, items.length, 1).setValues(sellers);
  if (brandCol) sheet.getRange(nextRow, brandCol, items.length, 1).setValues(brands);

  var invRange = ss.getRangeByName(cfg.inventoryRange);
  var invSheet = invRange ? invRange.getSheet() : null;
  var invValues = invRange ? invRange.getValues().map(function(r){
    return String(r[0]).trim();
  }) : [];
  items.forEach(function(it) {
    if (!invRange) return;
    var targetSn = String(it.serial).trim();
    var idx = invValues.indexOf(targetSn);
    if (idx > -1) {
      invSheet.deleteRow(invRange.getRow() + idx);
      invValues.splice(idx, 1);
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

addTiming([
  'onOpen',
  'showSaleDialog',
  'getColumnValues',
  'getLastDataRow',
  'getInventoryData',
  'submitOrder',
  'handleExternalOrders',
  'processExternalOrder',
  'getPersianDateTime',
  'gregorianToJalali'
]);
