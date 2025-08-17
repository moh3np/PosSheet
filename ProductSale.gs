function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('توسعه')
    .addSubMenu(ui.createMenu('فروش').addItem('فروش محصول', 'showSaleDialog'))
    .addToUi();
}

function showSaleDialog() {
  var ss = SpreadsheetApp.getActive();
  var snRange = ss.getRangeByName('InventorySN').getValues();
  var nameRange = ss.getRangeByName('InventoryName').getValues();
  var locRange = ss.getRangeByName('InventoryLocation').getValues();
  var priceRange = ss.getRangeByName('InventoryPrice').getValues();

  var inventory = { sns: [], names: [], locations: [], prices: [] };
  for (var i = 1; i < snRange.length; i++) {
    var sn = snRange[i][0];
    if (sn) {
      inventory.sns.push(sn);
      inventory.names.push(nameRange[i][0]);
      inventory.locations.push(locRange[i][0]);
      inventory.prices.push(priceRange[i][0]);
    }
  }

  var template = HtmlService.createTemplateFromFile('sale');
  template.inventory = inventory;
  var html = template.evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'فروش محصول');
}
