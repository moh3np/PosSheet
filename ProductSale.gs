function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('توسعه')
    .addItem('فروش محصول', 'showSaleDialog')
    .addToUi();
}

function showSaleDialog() {
  var tpl = HtmlService.createTemplateFromFile('sale');
  // Load SN list asynchronously on the client to speed up dialog opening
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
  for (var i = startIndex; i < snValues.length; i++) {
    var row = snRange.getRow() + i;
    var sn = sanitizeNumber_(snValues[i][0]);
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

function sanitizeNumber_(str) {
  return toEnglishNumber_(str).replace(/\D+/g, '');
}

function searchInventory(sn) {
  var ss = SpreadsheetApp.getActive();
  var snRange = ss.getRangeByName('InventorySN');
  if (!snRange) return null;
  var snNorm = sanitizeNumber_(sn);
  var snNum = Number(snNorm);
  var values = snRange.getValues();
  for (var i = 0; i < values.length; i++) {
    var cellVal = sanitizeNumber_(values[i][0]);
    var cellNum = Number(cellVal);
    if ((snNum && cellNum && snNum === cellNum) || cellVal === snNorm) {
      var row = snRange.getCell(i + 1, 1).getRow();
      return {
        name: getCellValueByName('InventoryName', row),
        brand: getCellValueByName('InventoryBrand', row),
        price: getCellValueByName('InventoryPrice', row),
        location: getCellValueByName('InventoryLocation', row)
      };
    }
  }
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
