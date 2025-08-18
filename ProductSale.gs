function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('توسعه')
    .addSubMenu(ui.createMenu('فروش').addItem('فروش محصول', 'showSaleDialog'))
    .addToUi();
}

function showSaleDialog() {
  var ss = SpreadsheetApp.getActive();
  var names = ss.getRangeByName('InventoryName').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;});
  var sns = ss.getRangeByName('InventorySN').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;});
  var persianSns = ss.getRangeByName('InventoryPersianSN').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;});
  var skus = ss.getRangeByName('InventorySKU').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;});
  var locations = ss.getRangeByName('InventoryLocation').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;});
  var prices = ss.getRangeByName('InventoryPrice').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;});
  var template = HtmlService.createTemplateFromFile('sale');
  template.inventoryData = {names:names, skus:skus, sns:sns, persianSNS:persianSns, locations:locations, prices:prices};
  var html = template.evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'فروش محصول');
}

function submitOrder(items) {
  if (!items || !items.length) {
    return;
  }
  var ss = SpreadsheetApp.getActive();
  var idRange = ss.getRangeByName('OrderID');
  var sheet = idRange.getSheet();
  var idCol = idRange.getColumn();
  var nameCol = ss.getRangeByName('OrderName').getColumn();
  var skuCol = ss.getRangeByName('OrderSKU').getColumn();
  var snCol = ss.getRangeByName('OrderSN').getColumn();
  var dateCol = ss.getRangeByName('OrderDate').getColumn();
  var priceCol = ss.getRangeByName('OrderPrice').getColumn();
  var paidCol = ss.getRangeByName('OrderPaidPrice').getColumn();
  var nextRow = sheet.getLastRow() + 1;
  var lastId = sheet.getRange(nextRow - 1, idCol).getValue();
  var orderId = lastId ? Number(lastId) + 1 : 1;
  var ids = [], names = [], skus = [], sns = [], dates = [], prices = [], paid = [];
  var dateStr = getPersianDateTime();
  items.forEach(function(it) {
    ids.push([orderId]);
    names.push([it.name]);
    skus.push([it.sku]);
    sns.push([it.serial]);
    dates.push([dateStr]);
    prices.push([it.price]);
    paid.push([it.paid]);
  });
  sheet.getRange(nextRow, idCol, items.length, 1).setValues(ids);
  sheet.getRange(nextRow, nameCol, items.length, 1).setValues(names);
  sheet.getRange(nextRow, skuCol, items.length, 1).setValues(skus);
  sheet.getRange(nextRow, snCol, items.length, 1).setValues(sns);
  sheet.getRange(nextRow, dateCol, items.length, 1).setValues(dates);
  sheet.getRange(nextRow, priceCol, items.length, 1).setValues(prices);
  sheet.getRange(nextRow, paidCol, items.length, 1).setValues(paid);
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
