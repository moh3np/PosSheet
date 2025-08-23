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
  var tlItems = [];
  var brItems = [];
  items.forEach(function(it){
    var prefix = it.sku.slice(0,2).toUpperCase();
    if (prefix === 'BR') brItems.push(it); else tlItems.push(it);
  });

  if (tlItems.length) {
    var tlSs = SpreadsheetApp.openById('1LIR_q1xrpdzcqoBJmNXTO0UJ9dksoBjS7h3Me4PRB1s');
    processCancelGroup(tlSs, tlItems, 'ToylandOrders', 'ToylandInventory', true);
  }
  if (brItems.length) {
    var brSs = SpreadsheetApp.openById('12-Khe_IZ9S7z_VN_LZQCHdcKEIgKDquviar8cSR_wG8');
    processCancelGroup(brSs, brItems, 'BuyruzPosOrders', 'BuyruzInventory', false);
  }
}

function processCancelGroup(ss, items, ordersRangeName, inventoryRangeName, transferPrice) {
  var ordersRange = ss.getRangeByName(ordersRangeName);
  if (!ordersRange) return;
  var sheet = ordersRange.getSheet();
  var lastRow = getLastDataRow(ordersRange.offset(0,3,ordersRange.getNumRows(),1));
  var startRow = ordersRange.getRow() + 1;
  if (lastRow < startRow) return;
  var dataRows = lastRow - startRow + 1;
  var numCols = ordersRange.getNumColumns();
  var data = sheet.getRange(startRow, ordersRange.getColumn(), dataRows, numCols).getValues();
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
      var locationValue = row[7] === 'مغازه' ? 'STORE' : row[7];
      var priceValue = transferPrice ? row[5] : '';
      invRows.push([row[1], row[9], row[2], row[10], row[3], row[8], priceValue, locationValue, '']);
    }
  });
  sheet.getRange(startRow, cancelCol, dataRows, 1).setValues(cancelValues);
  if (invRows.length) {
    var invRange = ss.getRangeByName(inventoryRangeName);
    var invSheet = invRange.getSheet();
    var baseCol = invRange.getColumn();
    var start = invSheet.getLastRow() + 1;
    invSheet.getRange(start, baseCol, invRows.length, invRange.getNumColumns()).setValues(invRows);
    var labelRange = invSheet.getRange(start, baseCol + 8, invRows.length, 1);
    labelRange.insertCheckboxes();
    labelRange.setValues(Array(invRows.length).fill([false]));
  }
}

