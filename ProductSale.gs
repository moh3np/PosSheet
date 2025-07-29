function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('توسعه')
    .addItem('فروش محصول', 'showSaleDialog')
    .addToUi();
}

var DEBUG = true;

function debugLog() {
  if (!DEBUG) return;
  var msg = Array.prototype.slice.call(arguments).join(' ');
  Logger.log(new Date().toISOString() + ' - ' + msg);
}

function showSaleDialog() {
  debugLog('Opening sale dialog');
  var tpl = HtmlService.createTemplateFromFile('sale');
  // Load the serial number list asynchronously on the client to avoid
  // delaying the dialog from opening.
  tpl.snList = [];
  var html = tpl.evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'فروش محصول');
}

function getInventorySNList() {
  var ss = SpreadsheetApp.getActive();
  var range = ss.getRangeByName('InventorySN');
  if (!range) return [];
  var sheet = range.getSheet();
  var frozen = sheet.getFrozenRows();
  var startIndex = Math.max(0, frozen - (range.getRow() - 1));
  var values = range.getValues();
  debugLog('Inventory SN list loaded', values.length - startIndex);
  return values.slice(startIndex).map(function(r){ return r[0]; });
}

function getInventoryData() {
  var ss = SpreadsheetApp.getActive();
  var snRange = ss.getRangeByName('InventorySN');
  if (!snRange) return [];
  var sheet = snRange.getSheet();
  var frozen = sheet.getFrozenRows();
  var startIndex = Math.max(0, frozen - (snRange.getRow() - 1));

  // Load all related ranges once to avoid per-row calls which are slow on
  // large datasets. We'll later trim the arrays to only the rows that
  // actually contain a serial number to speed up processing.
  var snValues = snRange.getValues();
  var nameValues = ss.getRangeByName('InventoryName')?.getValues() || [];
  var brandValues = ss.getRangeByName('InventoryBrand')?.getValues() || [];
  var priceValues = ss.getRangeByName('InventoryPrice')?.getValues() || [];
  var locationValues = ss.getRangeByName('InventoryLocation')?.getValues() || [];

  // Determine the last row that actually contains a serial number.
  var endIndex = snValues.length;
  while (endIndex > startIndex && !normalizeNumber_(snValues[endIndex - 1][0])) {
    endIndex--;
  }

  var data = [];
  debugLog('Loading inventory data rows:', endIndex - startIndex);
  for (var i = startIndex; i < endIndex; i++) {
    var sn = normalizeNumber_(snValues[i][0]);
    if (!sn) continue; // Skip blank rows inside the defined range
    data.push({
      sn: sn,
      name: nameValues[i] ? nameValues[i][0] : '-',
      brand: brandValues[i] ? brandValues[i][0] : '-',
      price: priceValues[i] ? priceValues[i][0] : '-',
      location: locationValues[i] ? locationValues[i][0] : '-'
    });
  }
  return data;
}

function toEnglishNumber_(str) {
  return String(str)
    .replace(/[\u06F0-\u06F9]/g, function(d){return d.charCodeAt(0)-1728;})
    .replace(/[\u0660-\u0669]/g, function(d){return d.charCodeAt(0)-1584;});
}

function normalizeNumber_(str) {
  return toEnglishNumber_(str).replace(/[^0-9]/g, '');
}

function searchInventory(sn) {
  var ss = SpreadsheetApp.getActive();
  var snRange = ss.getRangeByName('InventorySN');
  if (!snRange) {
    debugLog('InventorySN range not found');
    return null;
  }
  var snNorm = normalizeNumber_(sn);
  debugLog('Server search', sn, 'normalized to', snNorm);
  var snNum = Number(snNorm);
  var values = snRange.getValues();
  debugLog('Checking', values.length, 'inventory rows');
  for (var i = 0; i < values.length; i++) {
    var cellValOriginal = values[i][0];
    var cellVal = normalizeNumber_(cellValOriginal);
    var cellNum = Number(cellVal);
    debugLog('Row', i + 1, 'value', cellValOriginal, 'normalized', cellVal);
    if (snNorm === cellVal || (snNum && cellNum && snNum === cellNum)) {
      var row = snRange.getCell(i + 1, 1).getRow();
      debugLog(
        'Server found at row',
        row,
        'sn',
        cellVal,
        'name',
        getCellValueByName('InventoryName', row),
        'brand',
        getCellValueByName('InventoryBrand', row)
      );
      return {
        sn: cellVal,
        name: getCellValueByName('InventoryName', row),
        brand: getCellValueByName('InventoryBrand', row),
        price: getCellValueByName('InventoryPrice', row),
        location: getCellValueByName('InventoryLocation', row)
      };
    }
  }
  debugLog(
    'Server search failed for',
    sn,
    'normalized',
    snNorm,
    'after checking',
    values.length,
    'rows'
  );
  return null;
}

function getCellValueByName(rangeName, row) {
  var range = SpreadsheetApp.getActive().getRangeByName(rangeName);
  if (!range) return '-';
  var sheet = range.getSheet();
  var col = range.getColumn();
  var offset = row - range.getRow();
  if (offset >= 0 && offset < range.getNumRows()) {
    return sheet.getRange(row, col).getValue();
  }
  return '-';
}
