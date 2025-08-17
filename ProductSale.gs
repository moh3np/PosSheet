function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('توسعه')
    .addSubMenu(ui.createMenu('فروش').addItem('فروش محصول', 'showSaleDialog'))
    .addToUi();
}

function showSaleDialog() {
  var ss = SpreadsheetApp.getActive();
  var names = ss.getRangeByName('InventoryName').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;});
  var sns = ss.getRangeByName('InventorySN').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;});
  var locations = ss.getRangeByName('InventoryLocation').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;});
  var prices = ss.getRangeByName('InventoryPrice').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;});
  var template = HtmlService.createTemplateFromFile('sale');
  template.inventoryData = {names:names, sns:sns, locations:locations, prices:prices};
  var html = template.evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'فروش محصول');
}
