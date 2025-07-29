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

function toEnglishNumber_(str) {
  return String(str)
    .replace(/[\u06F0-\u06F9]/g, function(d){return d.charCodeAt(0)-1728;})
    .replace(/[\u0660-\u0669]/g, function(d){return d.charCodeAt(0)-1584;});
}

var inventoryCache_ = null;
var inventoryCacheTime_ = 0;
var CACHE_TTL_MS_ = 5 * 60 * 1000; // 5 minutes

function getInventoryMap_() {
  var now = Date.now();
  if (inventoryCache_ && now - inventoryCacheTime_ < CACHE_TTL_MS_) {
    return inventoryCache_;
  }
  var map = {};
  var ss = SpreadsheetApp.getActive();
  var range = ss.getRangeByName('InventorySN');
  if (!range) {
    inventoryCache_ = map;
    inventoryCacheTime_ = now;
    return map;
  }
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    var sn = toEnglishNumber_(values[i][0]).replace(/\s+/g, '');
    var row = range.getCell(i + 1, 1).getRow();
    if (sn) {
      map[sn] = row;
      var num = Number(sn);
      if (num) map[num] = row;
    }
  }
  inventoryCache_ = map;
  inventoryCacheTime_ = now;
  return map;
}

function searchInventory(sn) {
  var snNorm = toEnglishNumber_(sn).replace(/\s+/g, '');
  var snNum = Number(snNorm);
  var map = getInventoryMap_();
  var row = map[snNorm] || (snNum && map[snNum]);
  if (!row) return null;
  return {
    name: getCellValueByName('InventoryName', row),
    brand: getCellValueByName('InventoryBrand', row),
    price: getCellValueByName('InventoryPrice', row),
    location: getCellValueByName('InventoryLocation', row)
  };
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
