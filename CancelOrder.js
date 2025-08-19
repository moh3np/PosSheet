function showCancelDialog() {
  var ss = SpreadsheetApp.getActive();
  var getValuesByName = function(name) {
    var range = ss.getRangeByName(name);
    return range ? range.getValues().map(function(r){return r[0];}) : [];
  };
  var ids = getValuesByName('OrderID').slice(1).filter(String);
  var len = ids.length;
  var slice = function(arr){ return arr.slice(1, len + 1); };
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
  var getValues = function(ss, name){
    return ss.getRangeByName(name).getValues().map(function(r){ return r[0]; });
  };
  var ids = getValues(tlSs, 'OrderID').slice(1);
  var len = ids.length;
  var skus = getValues(tlSs, 'OrderSKU').slice(1, len + 1);
  var locations = getValues(tlSs, 'OrderLocation').slice(1, len + 1);
  var names = getValues(tlSs, 'OrderName').slice(1, len + 1);
  // price values are intentionally ignored when returning cancelled items to inventory
  var sellers = getValues(tlSs, 'OrderSeller').slice(1, len + 1);
  var sns = getValues(tlSs, 'OrderSN').slice(1, len + 1);
  var uniques = getValues(tlSs, 'OrderUniqueCode').slice(1, len + 1);
  var brands = getValues(tlSs, 'OrderBrand').slice(1, len + 1);
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
    var bSkus = getValues(brSs, 'StoreOrderSKU').slice(1);
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
  CacheService.getDocumentCache().remove('inventoryData');
}
