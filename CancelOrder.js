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
    var tlSs = timeStep('open:tlSs', function(){ return SpreadsheetApp.openById('1LIR_q1xrpdzcqoBJmNXTO0UJ9dksoBjS7h3Me4PRB1s'); });
    var tlOrders = timeStep('tlSs:getRangeByName:Orders', function(){ return tlSs.getRangeByName('Orders'); });
    var brSs = timeStep('open:brSs', function(){ return SpreadsheetApp.openById('12-Khe_IZ9S7z_VN_LZQCHdcKEIgKDquviar8cSR_wG8'); });
    var brOrders = timeStep('brSs:getRangeByName:StoreOrders', function(){ return brSs.getRangeByName('StoreOrders'); });
    var lastRows = {};
    lastRows[tlOrders.getSheet().getSheetId()] = getLastDataRow(tlOrders.offset(0,3,tlOrders.getNumRows(),1));
    lastRows[brOrders.getSheet().getSheetId()] = getLastDataRow(brOrders.offset(0,3,brOrders.getNumRows(),1));
    var getValues = function(range, idx){
      return timeStep('getValues:idx' + idx, function(){
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
      });
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
        var endTL = startTimer('handleTL');
        try {
          timeStep('handleTL:setCancel', function(){
            try { cancelRange.getCell(idx + 2, 1).setValue(true); } catch(e) {}
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
            try { brCancelRange.getCell(idx + 2, 1).setValue(true); } catch(e) {}
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
