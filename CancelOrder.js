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
    sns: slice(getValuesByName('OrderSN')).map(String),
    persianSNS: slice(getValuesByName('OrderPersianSN')),
    prices: slice(getValuesByName('OrderPrice')),
    paidPrices: slice(getValuesByName('OrderPaidPrice')),
    skus: slice(getValuesByName('OrderSKU')).map(function(s){return s != null ? String(s) : ''; }),
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

function cancelOrders(items) {
  Logger.log('شروع تابع لغو سفارش با داده‌ها: %s', JSON.stringify(items));
  if (!items || !items.length) {
    Logger.log('هیچ آیتمی برای لغو سفارش ارسال نشده است.');
    return;
  }
  // اطمینان از اینکه سریال و SKU به صورت رشته هستند
  items = items.map(function(it){
    return {
      sn: String(it.sn),
      sku: it.sku != null ? String(it.sku) : ''
    };
  });
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
  var sns = getValues(tlSs, 'OrderSN').map(function(s){ return s != null ? String(s) : ''; });
  var len = sns.length;
  Logger.log('تعداد سریال‌های TL: %s', len);
  Logger.log('لیست سریال‌های TL: %s', JSON.stringify(sns));

  var skus = getValues(tlSs, 'OrderSKU')
    .map(function(s){ return s != null ? String(s) : ''; })
    .slice(0, len);
  Logger.log('لیست SKU ها: %s', JSON.stringify(skus));

  var locations = getValues(tlSs, 'OrderLocation').slice(0, len);
  Logger.log('لیست موقعیت‌ها: %s', JSON.stringify(locations));
  var names = getValues(tlSs, 'OrderName').slice(0, len);
  Logger.log('لیست نام‌ها: %s', JSON.stringify(names));
  // price values are intentionally ignored when returning cancelled items to inventory
  var sellers = getValues(tlSs, 'OrderSeller').slice(0, len);
  Logger.log('لیست فروشنده‌ها: %s', JSON.stringify(sellers));
  var uniques = getValues(tlSs, 'OrderUniqueCode').slice(0, len);
  Logger.log('لیست کدهای یکتا: %s', JSON.stringify(uniques));
  var brands = getValues(tlSs, 'OrderBrand').slice(0, len);
  Logger.log('لیست برندها: %s', JSON.stringify(brands));
  var cancelRange = tlSs.getRangeByName('OrderCancellation');
  Logger.log('رنج OrderCancellation دریافت شد.');

  // دریافت داده‌های شیت BR برای پشتیبانی از لغو سفارشات با پیشوند BR
  var brSs = SpreadsheetApp.openById('12-Khe_IZ9S7z_VN_LZQCHdcKEIgKDquviar8cSR_wG8');
  Logger.log('شیت BR با موفقیت باز شد.');
  var brSns = getValues(brSs, 'StoreOrderSN').map(function(s){ return s != null ? String(s) : ''; });
  Logger.log('تعداد سریال‌های BR: %s', brSns.length);
  var brSkus = getValues(brSs, 'StoreOrderSKU')
    .map(function(s){ return s != null ? String(s) : ''; })
    .slice(0, brSns.length);
  var brLocations = getValues(brSs, 'StoreOrderLocation').slice(0, brSns.length);
  var brNames = getValues(brSs, 'StoreOrderName').slice(0, brSns.length);
  var brSellers = getValues(brSs, 'StoreOrderSeller').slice(0, brSns.length);
  var brUniques = getValues(brSs, 'StoreOrderUniqueCode').slice(0, brSns.length);
  var brBrands = getValues(brSs, 'StoreOrderBrand').slice(0, brSns.length);
  var brCancelRange = brSs.getRangeByName('StoreOrderCancellation');
  Logger.log('رنج StoreOrderCancellation دریافت شد.');

  items.forEach(function(item){
    Logger.log('--- بررسی سریال %s ---', item.sn);
    var prefix = item.sku.slice(0,2).toUpperCase();
    if (prefix === 'BR') {
      var brIdx = brSns.indexOf(item.sn);
      if (brIdx >= 0) {
        Logger.log('سریال در BR و در سطر %s یافت شد', brIdx + 2);
        handleBR(brIdx);
      } else {
        Logger.log('سریال %s در BR یافت نشد', item.sn);
      }
    } else {
      var idx = sns.indexOf(item.sn);
      if (idx >= 0) {
        Logger.log('سریال در TL و در سطر %s یافت شد', idx + 2);
        Logger.log('اطلاعات ردیف انتخاب‌شده: location=%s, name=%s, seller=%s, sn=%s, unique=%s, brand=%s',
                   locations[idx], names[idx], sellers[idx], sns[idx], uniques[idx], brands[idx]);
        handleTL(idx);
      } else {
        Logger.log('سریال %s در TL یافت نشد', item.sn);
      }
    }
  });

  function handleTL(idx){
    Logger.log('شروع handleTL برای سطر %s', idx + 2);
    try {
      var cell = cancelRange.getCell(idx + 2, 1);
      Logger.log('مقدار فعلی لغو سفارش: %s', cell.getValue());
      cell.setValue(true);
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

  function handleBR(idx){
    if (idx < 0) {
      Logger.log('شاخص نامعتبر برای BR: %s', idx);
      return;
    }
    Logger.log('شروع handleBR برای سطر %s', idx + 2);
    try {
      brCancelRange.getCell(idx + 2, 1).setValue(true);
      Logger.log('لغو سفارش BR در سطر %s ثبت شد', idx + 2);
    } catch(e) {
      Logger.log('خطا در تنظیم لغو سفارش BR: %s', e);
    }
    var data = {
      location: brLocations[idx],
      name: brNames[idx],
      seller: brSellers[idx],
      sku: brSkus[idx],
      sn: brSns[idx],
      unique: brUniques[idx],
      brand: brBrands[idx]
    };
    Logger.log('داده‌های ارسال به موجودی BR: %s', JSON.stringify(data));
    appendToInventory(brSs, data, true);
  }

  function appendToInventory(ss, data, isStore){
    Logger.log('افزودن به موجودی، isStore=%s , sku=%s', isStore, data.sku);
    Logger.log('داده‌های افزوده‌شونده به موجودی: %s', JSON.stringify(data));
    var locRange = ss.getRangeByName('InventoryLocation');
    var sheet = locRange.getSheet();
    var row = sheet.getLastRow() + 1;
    Logger.log('سطر جدید موجودی: %s', row);
    var locationValue = data.location === 'مغازه' ? 'STORE' : data.location;
    sheet.getRange(row, locRange.getColumn()).setValue(locationValue);
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
