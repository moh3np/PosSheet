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

function showCancelDialog() {
  var ss = SpreadsheetApp.getActive();
  var snRange = ss.getRangeByName('OrderSN');
  if (!snRange) return;
  var sheet = snRange.getSheet();
  var lastRow = getLastDataRow(snRange);
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
  var end = startTimer('cancelOrders');
  try {
    if (!items || !items.length) {
      return;
    }
    items = items.map(function(it){
      return { sn: String(it.sn), sku: it.sku != null ? String(it.sku) : '' };
    });
    var tlSs = SpreadsheetApp.openById('1LIR_q1xrpdzcqoBJmNXTO0UJ9dksoBjS7h3Me4PRB1s');
    var tlSnRange = tlSs.getRangeByName('OrderSN');
    var brSs = SpreadsheetApp.openById('12-Khe_IZ9S7z_VN_LZQCHdcKEIgKDquviar8cSR_wG8');
    var brSnRange = brSs.getRangeByName('StoreOrderSN');
    var lastRows = {};
    lastRows[tlSnRange.getSheet().getSheetId()] = getLastDataRow(tlSnRange);
    lastRows[brSnRange.getSheet().getSheetId()] = getLastDataRow(brSnRange);
    var getValues = function(ss, name){
      var range = ss.getRangeByName(name);
      if (!range) return [];
      var sheet = range.getSheet();
      var sheetId = sheet.getSheetId();
      var lastRow = lastRows[sheetId];
      var startRow = range.getRow() + 1;
      var col = range.getColumn();
      if (lastRow < startRow) return [];
      return sheet
        .getRange(startRow, col, lastRow - startRow + 1, 1)
        .getValues()
        .map(function(r){return r[0];});
    };
    var sns = getValues(tlSs, 'OrderSN').map(function(s){ return s != null ? String(s) : ''; });
    var len = sns.length;
    var skus = getValues(tlSs, 'OrderSKU').map(function(s){ return s != null ? String(s) : ''; }).slice(0, len);
    var locations = getValues(tlSs, 'OrderLocation').slice(0, len);
    var names = getValues(tlSs, 'OrderName').slice(0, len);
    var sellers = getValues(tlSs, 'OrderSeller').slice(0, len);
    var uniques = getValues(tlSs, 'OrderUniqueCode').slice(0, len);
    var brands = getValues(tlSs, 'OrderBrand').slice(0, len);
    var cancelRange = tlSs.getRangeByName('OrderCancellation');

    var brSns = getValues(brSs, 'StoreOrderSN').map(function(s){ return s != null ? String(s) : ''; });
    var brSkus = getValues(brSs, 'StoreOrderSKU').map(function(s){ return s != null ? String(s) : ''; }).slice(0, brSns.length);
    var brLocations = getValues(brSs, 'StoreOrderLocation').slice(0, brSns.length);
    var brNames = getValues(brSs, 'StoreOrderName').slice(0, brSns.length);
    var brSellers = getValues(brSs, 'StoreOrderSeller').slice(0, brSns.length);
    var brUniques = getValues(brSs, 'StoreOrderUniqueCode').slice(0, brSns.length);
    var brBrands = getValues(brSs, 'StoreOrderBrand').slice(0, brSns.length);
    var brCancelRange = brSs.getRangeByName('StoreOrderCancellation');

    items.forEach(function(item){
      var prefix = item.sku.slice(0,2).toUpperCase();
      if (prefix === 'BR') {
        var brIdx = brSns.indexOf(item.sn);
        if (brIdx >= 0) {
          handleBR(brIdx);
        }
      } else {
        var idx = sns.indexOf(item.sn);
        if (idx >= 0) {
          handleTL(idx);
        }
      }
    });

    function handleTL(idx){
      var endTL = startTimer('handleTL');
      try {
        try { cancelRange.getCell(idx + 2, 1).setValue(true); } catch(e) {}
        var data = {
          location: locations[idx],
          name: names[idx],
          seller: sellers[idx],
          sku: skus[idx],
          sn: sns[idx],
          unique: uniques[idx],
          brand: brands[idx]
        };
        appendToInventory(tlSs, data, false);
      } finally {
        endTL();
      }
    }

    function handleBR(idx){
      var endBR = startTimer('handleBR');
      try {
        if (idx < 0) return;
        try { brCancelRange.getCell(idx + 2, 1).setValue(true); } catch(e) {}
        var data = {
          location: brLocations[idx],
          name: brNames[idx],
          seller: brSellers[idx],
          sku: brSkus[idx],
          sn: brSns[idx],
          unique: brUniques[idx],
          brand: brBrands[idx]
        };
        appendToInventory(brSs, data, true);
      } finally {
        endBR();
      }
    }

    function appendToInventory(ss, data, isStore){
      var endAI = startTimer('appendToInventory');
      try {
        var locRange = ss.getRangeByName('InventoryLocation');
        var sheet = locRange.getSheet();
        var row = sheet.getLastRow() + 1;
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
      } finally {
        endAI();
      }
    }
  } finally {
    end();
  }
}

addTiming(['getLastDataRow', 'showCancelDialog']);
