function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('توسعه')
    .addItem('فروش محصول', 'showSaleDialog')
    .addToUi();
}

var DEBUG = false;

function debugLog() {
  if (!DEBUG) return;
  for (var i = 0; i < arguments.length; i++) {
    Logger.log(arguments[i]);
  }
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

  var nameRange = ss.getRangeByName('InventoryName');
  var brandRange = ss.getRangeByName('InventoryBrand');
  var priceRange = ss.getRangeByName('InventoryPrice');
  var locationRange = ss.getRangeByName('InventoryLocation');

  var snValues = snRange.getValues();
  var data = [];
  debugLog('Loading inventory data rows:', snValues.length - startIndex);
  for (var i = startIndex; i < snValues.length; i++) {
    var row = snRange.getRow() + i;
    var sn = normalizeNumber_(snValues[i][0]);
    data.push({
      sn: sn,
      name: getCellValueByName('InventoryName', row),
      brand: getCellValueByName('InventoryBrand', row),
      price: getCellValueByName('InventoryPrice', row),
      location: getCellValueByName('InventoryLocation', row)
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
  if (!snRange) return null;
  var snNorm = normalizeNumber_(sn);
  debugLog('Server search', sn, 'normalized to', snNorm);
  var snNum = Number(snNorm);
  var values = snRange.getValues();
  for (var i = 0; i < values.length; i++) {
    var cellVal = normalizeNumber_(values[i][0]);
    var cellNum = Number(cellVal);
    if (snNorm === cellVal || (snNum && cellNum && snNum === cellNum)) {
      var row = snRange.getCell(i + 1, 1).getRow();
      debugLog('Server found at row', row, 'sn', cellVal);
      return {
        sn: cellVal,
        name: getCellValueByName('InventoryName', row),
        brand: getCellValueByName('InventoryBrand', row),
        price: getCellValueByName('InventoryPrice', row),
        location: getCellValueByName('InventoryLocation', row)
      };
    }
  }
  debugLog('Server search failed for', sn);
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
