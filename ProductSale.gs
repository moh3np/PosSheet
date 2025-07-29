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
    .replace(/[\u06F0-\u06F9]/g, function(d) {
      return String.fromCharCode(d.charCodeAt(0) - 1728);
    })
    .replace(/[\u0660-\u0669]/g, function(d) {
      return String.fromCharCode(d.charCodeAt(0) - 1584);
    });
}

function loadInventoryMap_() {
  var cache = CacheService.getDocumentCache();
  var cached = cache.get('inventory_map');
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }
  var map = {};
  var ss = SpreadsheetApp.getActive();
  var snRange = ss.getRangeByName('InventorySN');
  if (!snRange) return map;
  var values = snRange.getValues();
  for (var i = 0; i < values.length; i++) {
    var key = toEnglishNumber_(values[i][0]).replace(/\s+/g, '');
    if (key) {
      map[key] = snRange.getCell(i + 1, 1).getRow();
    }
  }
  cache.put('inventory_map', JSON.stringify(map), 300);
  return map;
}

function searchInventory(sn) {
  var snNorm = toEnglishNumber_(sn).replace(/\s+/g, '');
  var map = loadInventoryMap_();
  var row = map[snNorm];
  if (row) {
    return {
      name: getCellValueByName('InventoryName', row),
      brand: getCellValueByName('InventoryBrand', row),
      price: getCellValueByName('InventoryPrice', row),
      location: getCellValueByName('InventoryLocation', row)
    };
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
