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
  if (!orderIds || !orderIds.length) return;
  var tlSs = SpreadsheetApp.openById('1LIR_q1xrpdzcqoBJmNXTO0UJ9dksoBjS7h3Me4PRB1s');
  var lastRows = {};
  var getValues = function(ss, name){
    var range = ss.getRangeByName(name);
    if (!range) return [];
    var sheet = range.getSheet();
    var sheetId = sheet.getSheetId();
    var lastRow = lastRows[sheetId];
    if (!lastRow) {
      lastRow = getLastRowInRange(sheet, 1, 4);
      lastRows[sheetId] = lastRow;
    }
    var startRow = range.getRow() + 1;
    var col = range.getColumn();
    if (lastRow < startRow) return [];
    return sheet
      .getRange(startRow, col, lastRow - startRow + 1, 1)
      .getValues()
      .map(function(r){return r[0];});
  };
  var ids = getValues(tlSs, 'OrderID');
  var len = ids.length;
  var skus = getValues(tlSs, 'OrderSKU').slice(0, len);
  var locations = getValues(tlSs, 'OrderLocation').slice(0, len);
  var names = getValues(tlSs, 'OrderName').slice(0, len);
  // price values are intentionally ignored when returning cancelled items to inventory
  var sellers = getValues(tlSs, 'OrderSeller').slice(0, len);
  var sns = getValues(tlSs, 'OrderSN').slice(0, len);
  var uniques = getValues(tlSs, 'OrderUniqueCode').slice(0, len);
  var brands = getValues(tlSs, 'OrderBrand').slice(0, len);
  var cancelRange = tlSs.getRangeByName('OrderCancellation');

  orderIds.forEach(function(id){
    var idx = ids.indexOf(id);
    if (idx < 0) return;
    var sku = skus[idx] || '';
    if (sku.slice(0,2).toUpperCase() === 'BR') {
      handleBR(sku);
    } else if (sku.slice(0,2).toUpperCase() === 'TL') {
      handleTL(idx);
    }
  });

  function handleTL(idx){
    cancelRange.getCell(idx + 2, 1).setValue(true);
    appendToInventory(tlSs, {
      location: locations[idx],
      name: names[idx],
      seller: sellers[idx],
      sku: skus[idx],
      sn: sns[idx],
      unique: uniques[idx],
      brand: brands[idx]
    }, false);
  }

  function handleBR(sku){
    var brSs = SpreadsheetApp.openById('12-Khe_IZ9S7z_VN_LZQCHdcKEIgKDquviar8cSR_wG8');
    var bSkus = getValues(brSs, 'StoreOrderSKU');
    var idx = bSkus.indexOf(sku);
    if (idx < 0) return;
    brSs.getRangeByName('StoreOrderCancellation').getCell(idx + 2, 1).setValue(true);
    var data = {
      location: brSs.getRangeByName('StoreOrderLocation').getCell(idx + 2, 1).getValue(),
      name: brSs.getRangeByName('StoreOrderName').getCell(idx + 2, 1).getValue(),
      seller: brSs.getRangeByName('StoreOrderSeller').getCell(idx + 2, 1).getValue(),
      sku: sku,
      sn: brSs.getRangeByName('StoreOrderSN').getCell(idx + 2, 1).getValue(),
      unique: brSs.getRangeByName('StoreOrderUniqueCode').getCell(idx + 2, 1).getValue(),
      brand: brSs.getRangeByName('StoreOrderBrand').getCell(idx + 2, 1).getValue()
    };
    appendToInventory(brSs, data, true);
  }

  function appendToInventory(ss, data, isStore){
    var locRange = ss.getRangeByName('InventoryLocation');
    var sheet = locRange.getSheet();
    var row = sheet.getLastRow() + 1;
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
  }
}
