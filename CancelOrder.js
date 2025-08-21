var global = this;

function startTimer(name){
  var start = Date.now();
  return function(){
    console.log(name + ' took ' + (Date.now() - start) + 'ms');
  };
}

function timeStep(name, fn){
  var end = startTimer(name);
  try {
    return fn();
  } finally {
    end();
  }
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
  var values = timeStep('getLastDataRow:getValues', function(){
    return sheet.getRange(startRow, col, numRows, 1).getValues();
  });
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
  var ordersRange = ss.getRangeByName('Orders');
  if (!ordersRange) return;
  var sheet = ordersRange.getSheet();
  var snHeader = sheet.getRange(ordersRange.getRow(), ordersRange.getColumn() + 3);
  var lastRow = getLastDataRow(snHeader);
  var numRows = lastRow - ordersRange.getRow();
  var values = numRows > 0 ? sheet.getRange(ordersRange.getRow() + 1, ordersRange.getColumn(), numRows, ordersRange.getNumColumns()).getValues() : [];
  var ids = [], names = [], dates = [], sns = [], prices = [], paidPrices = [], skus = [], uniqueCodes = [], sellers = [], locations = [], brands = [], cancellations = [];
  values.forEach(function(r){
    ids.push(r[0]);
    names.push(r[1]);
    skus.push(r[2] != null ? String(r[2]) : '');
    sns.push(r[3] != null ? String(r[3]) : '');
    dates.push(r[4]);
    prices.push(r[5]);
    paidPrices.push(r[6]);
    locations.push(r[7]);
    sellers.push(r[8]);
    brands.push(r[9]);
    uniqueCodes.push(r[10]);
    cancellations.push(r[11]);
  });
  var len = ids.length;
  var getPersian = function(name){
    var range = ss.getRangeByName(name);
    if (!range) return [];
    var startRow = range.getRow() + 1;
    if (lastRow < startRow) return [];
    return sheet.getRange(startRow, range.getColumn(), lastRow - startRow + 1, 1)
      .getValues()
      .map(function(r){ return r[0]; })
      .slice(0, len);
  };
  var orderData = {
    ids: ids,
    persianIds: getPersian('OrderPersianID'),
    names: names,
    dates: dates,
    persianDates: getPersian('OrderPersianDate'),
    sns: sns,
    persianSNS: getPersian('OrderPersianSN'),
    prices: prices,
    paidPrices: paidPrices,
    skus: skus,
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
  var end = startTimer('cancelOrders');
  try {
    if (!items || !items.length) {
      return;
    }
    items = items.map(function(it){
      return { sn: String(it.sn), sku: it.sku != null ? String(it.sku) : '' };
    });
    var tlSs = timeStep('open:tlSs', function(){ return SpreadsheetApp.openById('1LIR_q1xrpdzcqoBJmNXTO0UJ9dksoBjS7h3Me4PRB1s'); });
    var brSs = timeStep('open:brSs', function(){ return SpreadsheetApp.openById('12-Khe_IZ9S7z_VN_LZQCHdcKEIgKDquviar8cSR_wG8'); });

    var tlOrders = timeStep('tlSs:getRangeByName:Orders', function(){ return tlSs.getRangeByName('Orders'); });
    var brOrders = timeStep('brSs:getRangeByName:StoreOrders', function(){ return brSs.getRangeByName('StoreOrders'); });

    var tlSheet = tlOrders.getSheet();
    var brSheet = brOrders.getSheet();
    var tlStartRow = tlOrders.getRow();
    var brStartRow = brOrders.getRow();
    var tlLastRow = getLastDataRow(tlSheet.getRange(tlStartRow, tlOrders.getColumn() + 3));
    var brLastRow = getLastDataRow(brSheet.getRange(brStartRow, brOrders.getColumn() + 3));

    var tlValues = tlLastRow > tlStartRow ? tlSheet.getRange(tlStartRow + 1, tlOrders.getColumn(), tlLastRow - tlStartRow, tlOrders.getNumColumns()).getValues() : [];
    var brValues = brLastRow > brStartRow ? brSheet.getRange(brStartRow + 1, brOrders.getColumn(), brLastRow - brStartRow, brOrders.getNumColumns()).getValues() : [];

    var sns = tlValues.map(function(r){ return r[3] != null ? String(r[3]) : ''; });
    var skus = tlValues.map(function(r){ return r[2] != null ? String(r[2]) : ''; });
    var locations = tlValues.map(function(r){ return r[7]; });
    var names = tlValues.map(function(r){ return r[1]; });
    var sellers = tlValues.map(function(r){ return r[8]; });
    var uniques = tlValues.map(function(r){ return r[10]; });
    var brands = tlValues.map(function(r){ return r[9]; });
    var cancelCol = tlOrders.getColumn() + 11;

    var brSns = brValues.map(function(r){ return r[3] != null ? String(r[3]) : ''; });
    var brSkus = brValues.map(function(r){ return r[2] != null ? String(r[2]) : ''; });
    var brLocations = brValues.map(function(r){ return r[7]; });
    var brNames = brValues.map(function(r){ return r[1]; });
    var brSellers = brValues.map(function(r){ return r[8]; });
    var brUniques = brValues.map(function(r){ return r[10]; });
    var brBrands = brValues.map(function(r){ return r[9]; });
    var brCancelCol = brOrders.getColumn() + 11;

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
          timeStep('handleTL:setCancel', function(){
            try { tlSheet.getRange(tlStartRow + idx + 1, cancelCol).setValue(true); } catch(e) {}
          });
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
          timeStep('handleBR:setCancel', function(){
            try { brSheet.getRange(brStartRow + idx + 1, brCancelCol).setValue(true); } catch(e) {}
          });
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
          var invRange = timeStep('appendToInventory:getInventoryRange', function(){
            return ss.getRangeByName('Inventory');
          });
          var sheet = timeStep('appendToInventory:getSheet', function(){
            return invRange.getSheet();
          });
          var row = timeStep('appendToInventory:getLastRow', function(){
            return sheet.getLastRow() + 1;
          });
          var baseCol = invRange.getColumn();
          var locationValue = data.location === 'مغازه' ? 'STORE' : data.location;
          timeStep('appendToInventory:setLocation', function(){
            sheet.getRange(row, baseCol + 7).setValue(locationValue);
          });
          timeStep('appendToInventory:setName', function(){
            sheet.getRange(row, baseCol + 0).setValue(data.name);
          });
          timeStep('appendToInventory:setSeller', function(){
            sheet.getRange(row, baseCol + 5).setValue(data.seller);
          });
          timeStep('appendToInventory:setSKU', function(){
            sheet.getRange(row, baseCol + 8).setValue(data.sku);
          });
          timeStep('appendToInventory:setSN', function(){
            sheet.getRange(row, baseCol + 4).setValue(data.sn);
          });
          timeStep('appendToInventory:setUnique', function(){
            sheet.getRange(row, baseCol + 3).setValue(data.unique);
          });
          timeStep('appendToInventory:setBrand', function(){
            sheet.getRange(row, baseCol + 1).setValue(data.brand);
          });
          timeStep('appendToInventory:insertLabel', function(){
            var lblRange = ss.getRangeByName('InventoryLablePrinted');
            var cell = sheet.getRange(row, lblRange.getColumn());
            cell.insertCheckboxes();
            cell.setValue(false);
          });
        } finally {
          endAI();
        }
      }
  } finally {
    end();
  }
}

addTiming(['getLastDataRow', 'showCancelDialog']);
