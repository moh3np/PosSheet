// Removed logging and timing utilities to simplify code and avoid side effects.

function showCancelDialog() {
  var ss = SpreadsheetApp.getActive();
  var posRange = ss.getRangeByName('PosOrders');
  if (!posRange) return;
  var sheet = posRange.getSheet();
  var lastRow = getLastDataRow(posRange.offset(0, 3));
  var startRow = posRange.getRow() + 1;
  if (lastRow < startRow) return;
  var numRows = lastRow - startRow + 1;
  var values = sheet
    .getRange(startRow, posRange.getColumn(), numRows, posRange.getNumColumns())
    .getValues();
  var ids = values.map(function(r){ return r[0]; });
  var names = values.map(function(r){ return r[1]; });
  var sns = values.map(function(r){ return r[3]; });
  var dates = values.map(function(r){ return r[4]; });
  var prices = values.map(function(r){ return r[5]; });
  var paidPrices = values.map(function(r){ return r[6]; });
  var locations = values.map(function(r){ return r[7]; });
  var sellers = values.map(function(r){ return r[8]; });
  var brands = values.map(function(r){ return r[9]; });
  var uniqueCodes = values.map(function(r){ return r[10]; });
  var cancellations = values.map(function(r){ return r[11]; });
  var skus = values.map(function(r){ return r[12]; });
  var persianSNS = values.map(function(r){ return r[13]; });
  var persianIds = values.map(function(r){ return r[14]; });
  var persianDates = values.map(function(r){ return r[15]; });
  var orderData = {
    ids: ids,
    persianIds: persianIds,
    names: names,
    dates: dates,
    persianDates: persianDates,
    sns: sns.map(String),
    persianSNS: persianSNS,
    prices: prices,
    paidPrices: paidPrices,
    skus: skus.map(function(s){return s != null ? String(s) : ''; }),
    uniqueCodes: uniqueCodes,
    sellers: sellers,
    locations: locations,
    brands: brands,
    cancellations: cancellations
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
    var tlOrders = tlSs.getRangeByName('ToylandOrders');
    var brSs = SpreadsheetApp.openById('12-Khe_IZ9S7z_VN_LZQCHdcKEIgKDquviar8cSR_wG8');
    var brOrders = brSs.getRangeByName('BuyruzPosOrders');
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
        appendToInventory(tlSs, data);
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
        appendToInventory(brSs, data);
    }

      function appendToInventory(ss, data){
        var invRange = ss.getRangeByName('Inventory');
        var sheet = invRange.getSheet();
        var row = sheet.getLastRow() + 1;
        var baseCol = invRange.getColumn();
        var locationValue = data.location === 'مغازه' ? 'STORE' : data.location;
        var rowValues = [];
        rowValues[0] = data.name;
        rowValues[1] = data.brand;
        rowValues[2] = '';
        rowValues[3] = data.unique;
        rowValues[4] = data.sn;
        rowValues[5] = data.seller;
        rowValues[6] = '';
        rowValues[7] = locationValue;
        rowValues[8] = data.sku;
        sheet.getRange(row, baseCol, 1, 9).setValues([rowValues]);
        var lblRange = ss.getRangeByName('InventoryLablePrinted');
        var cell = sheet.getRange(row, lblRange.getColumn());
        cell.insertCheckboxes();
        cell.setValue(false);
      }
}

