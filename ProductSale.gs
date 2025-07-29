function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('توسعه')
    .addItem('فروش محصول', 'showSaleDialog')
    .addToUi();
}

function showSaleDialog() {
  var html = HtmlService.createHtmlOutputFromFile('sale')
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'فروش محصول');
}

function searchInventory(sn) {
  var ss = SpreadsheetApp.getActive();
  var snRange = ss.getRangeByName('InventorySN');
  if (!snRange) return null;
  var values = snRange.getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(sn)) {
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
