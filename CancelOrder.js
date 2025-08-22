// Removed logging and timing utilities to simplify code and avoid side effects.

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
  if (!items || !items.length) {
    return;
  }
    items = items.map(function(it){
      return { sn: String(it.sn), sku: it.sku != null ? String(it.sku) : '' };
    });
    var tlSs = SpreadsheetApp.openById('1LIR_q1xrpdzcqoBJmNXTO0UJ9dksoBjS7h3Me4PRB1s');
    var tlOrders = tlSs.getRangeByName('Orders');
    var brSs = SpreadsheetApp.openById('12-Khe_IZ9S7z_VN_LZQCHdcKEIgKDquviar8cSR_wG8');
    var brOrders = brSs.getRangeByName('StoreOrders');
    var lastRows = {};
    lastRows[tlOrders.getSheet().getSheetId()] = getLastDataRow(tlOrders.offset(0,3,tlOrders.getNumRows(),1));
    lastRows[brOrders.getSheet().getSheetId()] = getLastDataRow(brOrders.offset(0,3,brOrders.getNumRows(),1));
    var getValues = function(range, idx){
      if (!range) return [];
      var sheet = range.getSheet();
      var sheetId = sheet.getSheetId();
      var lastRow = lastRows[sheetId];
      var startRow = range.getRow() + 1;
      var col = range.getColumn() + idx;
      if (lastRow < startRow) return [];
      return sheet
        .getRange(startRow, col, lastRow - startRow + 1, 1)
        .getValues()
        .map(function(r){return r[0];});
    };
    var sns = getValues(tlOrders, 3).map(function(s){ return s != null ? String(s) : ''; });
    var len = sns.length;
    var skus = getValues(tlOrders, 2).map(function(s){ return s != null ? String(s) : ''; }).slice(0, len);
    var locations = getValues(tlOrders, 7).slice(0, len);
    var names = getValues(tlOrders, 1).slice(0, len);
    var sellers = getValues(tlOrders, 8).slice(0, len);
    var uniques = getValues(tlOrders, 10).slice(0, len);
    var brands = getValues(tlOrders, 9).slice(0, len);
    var cancelRange = tlOrders.offset(0,11,tlOrders.getNumRows(),1);

    var brSns = getValues(brOrders, 3).map(function(s){ return s != null ? String(s) : ''; });
    var brSkus = getValues(brOrders, 2).map(function(s){ return s != null ? String(s) : ''; }).slice(0, brSns.length);
    var brLocations = getValues(brOrders, 7).slice(0, brSns.length);
    var brNames = getValues(brOrders, 1).slice(0, brSns.length);
    var brSellers = getValues(brOrders, 8).slice(0, brSns.length);
    var brUniques = getValues(brOrders, 10).slice(0, brSns.length);
    var brBrands = getValues(brOrders, 9).slice(0, brSns.length);
    var brCancelRange = brOrders.offset(0,11,brOrders.getNumRows(),1);

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
    }

      function handleBR(idx){
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
    }

      function appendToInventory(ss, data, isStore){
        var invRange = ss.getRangeByName('Inventory');
        var sheet = invRange.getSheet();
        var row = sheet.getLastRow() + 1;
        var baseCol = invRange.getColumn();
        var locationValue = data.location === 'مغازه' ? 'STORE' : data.location;
        sheet.getRange(row, baseCol + 7).setValue(locationValue);
        sheet.getRange(row, baseCol + 0).setValue(data.name);
        sheet.getRange(row, baseCol + 5).setValue(data.seller);
        sheet.getRange(row, baseCol + 8).setValue(data.sku);
        sheet.getRange(row, baseCol + 4).setValue(data.sn);
        sheet.getRange(row, baseCol + 3).setValue(data.unique);
        sheet.getRange(row, baseCol + 1).setValue(data.brand);
        var lblRange = ss.getRangeByName('InventoryLablePrinted');
        var cell = sheet.getRange(row, lblRange.getColumn());
        cell.insertCheckboxes();
        cell.setValue(false);
      }
}

