function getLastRowInRange(sheet, startCol, endCol) {
  var lastRow = sheet.getLastRow();
  var values = sheet
    .getRange(1, startCol, lastRow, endCol - startCol + 1)
    .getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    var row = values[i];
    for (var j = 0; j < row.length; j++) {
      if (row[j] !== '' && row[j] !== null) {
        return i + 1;
      }
    }
  }
  return 0;
}

function showCancelDialog() {
  var ss = SpreadsheetApp.getActive();
  var baseRange = ss.getRangeByName('OrderID');
  if (!baseRange) return;
  var sheet = baseRange.getSheet();
  var lastRow = getLastRowInRange(sheet, 1, 4);
  var getValuesByName = function(name) {
    var range = ss.getRangeByName(name);
    if (!range) return [];
    var startRow = range.getRow() + 1;
    var col = range.getColumn();
    if (lastRow < startRow) return [];
    return sheet
      .getRange(startRow, col, lastRow - startRow + 1, 1)
      .getValues()
      .map(function(r){return r[0];});
  };
  var ids = getValuesByName('OrderID');
  var len = ids.length;
  var slice = function(arr){ return arr.slice(0, len); };
  var orderData = {
    ids: ids,
    persianIds: slice(getValuesByName('OrderPersianID')),
    names: slice(getValuesByName('OrderName')),
    dates: slice(getValuesByName('OrderDate')),
    persianDates: slice(getValuesByName('OrderPersianDate')),
    sns: slice(getValuesByName('OrderSN')),
    persianSNS: slice(getValuesByName('OrderPersianSN')),
    prices: slice(getValuesByName('OrderPrice')),
    paidPrices: slice(getValuesByName('OrderPaidPrice')),
    skus: slice(getValuesByName('OrderSKU')),
    uniqueCodes: slice(getValuesByName('OrderUniqueCode')),
    sellers: slice(getValuesByName('OrderSeller')),
    locations: slice(getValuesByName('OrderLocation')),
    brands: slice(getValuesByName('OrderBrand')),
    cancellations: slice(getValuesByName('OrderCancellation'))
  };
  var template = HtmlService.createTemplateFromFile('cancel');
  template.orderData = orderData;
  var html = template.evaluate().setWidth(1200).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'لغو سفارش');
}

function cancelOrders(orderIds) {
  Logger.log('شروع تابع لغو سفارش با شناسه‌ها: %s', orderIds);
  if (!orderIds || !orderIds.length) {
    Logger.log('هیچ شناسه‌ای برای لغو سفارش ارسال نشده است.');
    return;
  }
  var tlSs = SpreadsheetApp.openById('1LIR_q1xrpdzcqoBJmNXTO0UJ9dksoBjS7h3Me4PRB1s');
  Logger.log('شیت TL با موفقیت باز شد.');
  var lastRows = {};
  var getValues = function(ss, name){
    Logger.log('در حال خواندن رنج %s', name);
    var range = ss.getRangeByName(name);
    if (!range) {
      Logger.log('رنج %s یافت نشد', name);
      return [];
    }
    var sheet = range.getSheet();
    var sheetId = sheet.getSheetId();
    var lastRow = lastRows[sheetId];
    if (!lastRow) {
      lastRow = getLastRowInRange(sheet, 1, 4);
      lastRows[sheetId] = lastRow;
      Logger.log('آخرین سطر برای شیت %s محاسبه شد: %s', sheet.getName(), lastRow);
    }
    var startRow = range.getRow() + 1;
    var col = range.getColumn();
    if (lastRow < startRow) {
      Logger.log('رنج %s داده‌ای برای خواندن ندارد', name);
      return [];
    }
    var vals = sheet
      .getRange(startRow, col, lastRow - startRow + 1, 1)
      .getValues()
      .map(function(r){return r[0];});
    Logger.log('%s مقدار از رنج %s خوانده شد', vals.length, name);
    return vals;
  };
  var ids = getValues(tlSs, 'OrderID');
  var len = ids.length;
  Logger.log('تعداد سفارشات موجود: %s', len);
  var skus = getValues(tlSs, 'OrderSKU').slice(0, len);
  var locations = getValues(tlSs, 'OrderLocation').slice(0, len);
  var names = getValues(tlSs, 'OrderName').slice(0, len);
  // price values are intentionally ignored when returning cancelled items to inventory
  var sellers = getValues(tlSs, 'OrderSeller').slice(0, len);
  var sns = getValues(tlSs, 'OrderSN').slice(0, len);
  var uniques = getValues(tlSs, 'OrderUniqueCode').slice(0, len);
  var brands = getValues(tlSs, 'OrderBrand').slice(0, len);
  var cancelRange = tlSs.getRangeByName('OrderCancellation');
  Logger.log('رنج OrderCancellation دریافت شد.');

  orderIds.forEach(function(id){
    Logger.log('--- بررسی شناسه %s ---', id);
    var idx = ids.indexOf(id);
    if (idx < 0) {
      Logger.log('شناسه %s در TL یافت نشد', id);
      return;
    }
    Logger.log('شناسه در سطر %s یافت شد', idx + 2);
    var sku = skus[idx] || '';
    Logger.log('SKU مربوطه: %s', sku);
    if (sku.slice(0,2).toUpperCase() === 'BR') {
      Logger.log('نوع سفارش BR است، اجرای handleBR');
      handleBR(sku);
    } else if (sku.slice(0,2).toUpperCase() === 'TL') {
      Logger.log('نوع سفارش TL است، اجرای handleTL');
      handleTL(idx);
    } else {
      Logger.log('پیشوند SKU ناشناخته است: %s', sku);
    }
  });

  function handleTL(idx){
    Logger.log('شروع handleTL برای سطر %s', idx + 2);
    try {
      cancelRange.getCell(idx + 2, 1).setValue(true);
      Logger.log('لغو سفارش در TL به true تنظیم شد');
    } catch(e) {
      Logger.log('خطا در تنظیم لغو سفارش TL: %s', e);
    }
    var data = {
      location: locations[idx],
      name: names[idx],
      seller: sellers[idx],
      sku: skus[idx],
      sn: sns[idx],
      unique: uniques[idx],
      brand: brands[idx]
    };
    Logger.log('داده‌های ارسال به موجودی: %s', JSON.stringify(data));
    appendToInventory(tlSs, data, false);
  }

  function handleBR(sku){
    Logger.log('شروع handleBR برای SKU: %s', sku);
    var brSs = SpreadsheetApp.openById('12-Khe_IZ9S7z_VN_LZQCHdcKEIgKDquviar8cSR_wG8');
    Logger.log('شیت BR با موفقیت باز شد');
    var bSkus = getValues(brSs, 'StoreOrderSKU');
    Logger.log('تعداد SKU های BR: %s', bSkus.length);
    var idx = bSkus.indexOf(sku);
    if (idx < 0) {
      Logger.log('SKU %s در BR یافت نشد', sku);
      return;
    }
    try {
      brSs.getRangeByName('StoreOrderCancellation').getCell(idx + 2, 1).setValue(true);
      Logger.log('لغو سفارش BR در سطر %s ثبت شد', idx + 2);
    } catch(e) {
      Logger.log('خطا در تنظیم لغو سفارش BR: %s', e);
    }
    var data = {
      location: brSs.getRangeByName('StoreOrderLocation').getCell(idx + 2, 1).getValue(),
      name: brSs.getRangeByName('StoreOrderName').getCell(idx + 2, 1).getValue(),
      seller: brSs.getRangeByName('StoreOrderSeller').getCell(idx + 2, 1).getValue(),
      sku: sku,
      sn: brSs.getRangeByName('StoreOrderSN').getCell(idx + 2, 1).getValue(),
      unique: brSs.getRangeByName('StoreOrderUniqueCode').getCell(idx + 2, 1).getValue(),
      brand: brSs.getRangeByName('StoreOrderBrand').getCell(idx + 2, 1).getValue()
    };
    Logger.log('داده‌های ارسال به موجودی BR: %s', JSON.stringify(data));
    appendToInventory(brSs, data, true);
  }

  function appendToInventory(ss, data, isStore){
    Logger.log('افزودن به موجودی، isStore=%s , sku=%s', isStore, data.sku);
    var locRange = ss.getRangeByName('InventoryLocation');
    var sheet = locRange.getSheet();
    var row = sheet.getLastRow() + 1;
    Logger.log('سطر جدید موجودی: %s', row);
    sheet.getRange(row, locRange.getColumn()).setValue(data.location);
    sheet.getRange(row, ss.getRangeByName(isStore ? 'InventoryProductName' : 'InventoryName').getColumn()).setValue(data.name);
    sheet.getRange(row, ss.getRangeByName('InventorySupplier').getColumn()).setValue(data.seller);
    sheet.getRange(row, ss.getRangeByName('InventorySKU').getColumn()).setValue(data.sku);
    sheet.getRange(row, ss.getRangeByName('InventorySN').getColumn()).setValue(data.sn);
    sheet.getRange(row, ss.getRangeByName('InventoryUniqueCode').getColumn()).setValue(data.unique);
    sheet.getRange(row, ss.getRangeByName(isStore ? 'InventoryProductBrand' : 'InventoryBrand').getColumn()).setValue(data.brand);
    var lblRange = ss.getRangeByName('InventoryLablePrinted');
    var cell = sheet.getRange(row, lblRange.getColumn());
    cell.insertCheckboxes();
    cell.setValue(false);
    Logger.log('آیتم با موفقیت به موجودی افزوده شد');
  }
}
