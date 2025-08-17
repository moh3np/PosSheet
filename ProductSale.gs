function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('توسعه')
    .addSubMenu(ui.createMenu('فروش').addItem('فروش محصول', 'showSaleDialog'))
    .addToUi();
}

function showSaleDialog() {
  var data = getInventoryData();
  var template = HtmlService.createTemplateFromFile('sale');
  template.inventoryData = data;
  var html = template.evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'فروش محصول');
}

function getInventoryData() {
  var ss = SpreadsheetApp.getActive();
  function fetchRange(name) {
    var values = ss.getRangeByName(name).getValues().flat();
    return values.slice(1).filter(String);
  }
  return {
    name: fetchRange('InventoryName'),
    sn: fetchRange('InventorySN'),
    location: fetchRange('InventoryLocation'),
    price: fetchRange('InventoryPrice'),
  };
}
