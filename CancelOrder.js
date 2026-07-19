// Removed logging and timing utilities to simplify code and avoid side effects.

function showCancelDialog() {
  var ss = SpreadsheetApp.getActive();
  var posOrdersRangeName = Config.namedRanges.posOrders;
  var posRange = ss.getRangeByName(posOrdersRangeName);
  if (!posRange) return;
  var sheet = posRange.getSheet();
  var lastRow = getLastDataRow(posRange.offset(0, 3));
  var startRow = getDataStartRow(posRange);
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
  var tlItems = [];
  var brItems = [];
  items.forEach(function(it){
    var prefix = it.sku.slice(0,2).toUpperCase();
    if (prefix === 'BR') brItems.push(it); else tlItems.push(it);
  });

  if (tlItems.length) {
    var tlSs = SpreadsheetApp.openById(Config.sheets.tolyandSheet);
    processCancelGroup(tlSs, tlItems, Config.namedRanges.toylandOrders, Config.namedRanges.toylandInventory, true);
  }
  if (brItems.length) {
    var brSs = SpreadsheetApp.openById(Config.sheets.buyruzSheet);
    processCancelGroup(brSs, brItems, Config.namedRanges.buyruzOrders, Config.namedRanges.buyruzInventory, false);
  }
}

function processCancelGroup(ss, items, ordersRangeName, inventoryRangeName, transferPrice) {
  var ordersRange = ss.getRangeByName(ordersRangeName);
  if (!ordersRange) return;
  var sheet = ordersRange.getSheet();
  var lastRow = getLastDataRow(ordersRange.offset(0,3,ordersRange.getNumRows(),1));
  var startRow = getDataStartRow(ordersRange);
  if (lastRow < startRow) return;
  var dataRows = lastRow - startRow + 1;
  var numCols = ordersRange.getNumColumns();
  var dataRange = sheet.getRange(startRow, ordersRange.getColumn(), dataRows, numCols);
  var data = dataRange.getValues();
  var displayData = dataRange.getDisplayValues();
  var snMap = {};
  for (var i = 0; i < data.length; i++) {
    snMap[String(data[i][3]).trim()] = i;
  }
  var cancelCol = ordersRange.getColumn() + 11;
  var cancelValues = data.map(function(row){ return [row[11]]; });
  var invRows = [];
  items.forEach(function(it){
    var idx = snMap[it.sn];
    if (idx != null) {
      cancelValues[idx][0] = true;
      var row = data[idx];
      var displayRow = displayData[idx];
      var locationValue = row[7] === 'مغازه' ? 'STORE' : row[7];
      var priceValue = transferPrice && inventoryRangeName !== Config.namedRanges.toylandInventory ? row[5] : '';
      var uniqueCodeValue = displayRow && displayRow.length > 10 ? displayRow[10] : row[10];
      invRows.push([row[1], row[9], row[2], uniqueCodeValue || '', row[3], row[8], priceValue, locationValue, '']);
    }
  });
  sheet.getRange(startRow, cancelCol, dataRows, 1).setValues(cancelValues);
  if (invRows.length) {
    if (inventoryRangeName === Config.namedRanges.buyruzInventory) {
      appendBuyruzCancelInventoryRows_(ss, invRows);
      return;
    }
    var invRange = ss.getRangeByName(inventoryRangeName);
    var invSheet = invRange.getSheet();
    var baseCol = invRange.getColumn();
    var numInvCols = invRange.getNumColumns();
    var start = getLastDataRow(invRange) + 1;
    var skipToylandColumn = Config.inventory.skipToylandColumn;
    var isToylandInventory = inventoryRangeName === Config.namedRanges.toylandInventory && numInvCols > skipToylandColumn;
    if (isToylandInventory) {
      var leftCols = Math.min(skipToylandColumn, numInvCols);
      if (leftCols > 0) {
        var leftRange = invSheet.getRange(start, baseCol, invRows.length, leftCols);
        var leftValues = invRows.map(function(r){ return r.slice(0, skipToylandColumn); });
        leftRange.setValues(leftValues);
      }
      var rightCols = numInvCols - (skipToylandColumn + 1);
      if (rightCols > 0) {
        var rightRange = invSheet.getRange(start, baseCol + skipToylandColumn + 1, invRows.length, rightCols);
        var rightValues = invRows.map(function(r){ return r.slice(skipToylandColumn + 1); });
        rightRange.setValues(rightValues);
      }
    } else {
      invSheet.getRange(start, baseCol, invRows.length, numInvCols).setValues(invRows);
    }
    var labelRange = invSheet.getRange(start, baseCol + 8, invRows.length, 1);
    labelRange.insertCheckboxes();
    labelRange.setValues(Array(invRows.length).fill([false]));
  }
}

function appendBuyruzCancelInventoryRows_(ss, invRows) {
  var ranges = getBuyruzCancelInventoryRanges_(ss);
  if (!ranges) return;

  var start = getLastDataRow(ranges.sn) + 1;
  getBuyruzCancelColumnRange_(ranges.sku, start, invRows.length)
    .setValues(invRows.map(function(r){ return [r[2]]; }));
  getBuyruzCancelColumnRange_(ranges.uniqueCode, start, invRows.length)
    .setValues(invRows.map(function(r){ return [r[3]]; }));
  getBuyruzCancelColumnRange_(ranges.sn, start, invRows.length)
    .setValues(invRows.map(function(r){ return [r[4]]; }));
  getBuyruzCancelColumnRange_(ranges.location, start, invRows.length)
    .setValues(invRows.map(function(r){ return [r[7]]; }));

  var labelRange = getBuyruzCancelColumnRange_(ranges.labelPrinted, start, invRows.length);
  labelRange.insertCheckboxes();
  labelRange.setValues(Array(invRows.length).fill([false]));
}

function getBuyruzCancelInventoryRanges_(ss) {
  var namedRanges = Config.namedRanges;
  var ranges = {
    productName: ss.getRangeByName(namedRanges.buyruzInventoryProductName),
    productBrand: ss.getRangeByName(namedRanges.buyruzInventoryProductBrand),
    sku: ss.getRangeByName(namedRanges.buyruzInventorySku),
    uniqueCode: ss.getRangeByName(namedRanges.buyruzInventoryUniqueCode),
    sn: ss.getRangeByName(namedRanges.buyruzInventorySN),
    supplier: ss.getRangeByName(namedRanges.buyruzInventorySupplier),
    price: ss.getRangeByName(namedRanges.buyruzInventoryPrice),
    location: ss.getRangeByName(namedRanges.buyruzInventoryLocation),
    labelPrinted: ss.getRangeByName(namedRanges.buyruzInventoryLabelPrinted)
  };
  for (var key in ranges) {
    if (!ranges[key]) {
      return null;
    }
  }
  return ranges;
}

function getBuyruzCancelColumnRange_(range, startRow, numRows) {
  return range.getSheet().getRange(startRow, range.getColumn(), numRows, 1);
}
