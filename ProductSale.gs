function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('توسعه')
    .addSubMenu(ui.createMenu('فروش').addItem('فروش محصول', 'showSaleDialog'))
    .addToUi();
}

function showSaleDialog() {
  var ss = SpreadsheetApp.getActive();

  function getValues(name) {
    var range = ss.getRangeByName(name);
    if (!range) return [];
    var values = range.getValues().flat();
    if (values.length > 0) {
      values = values.slice(1); // remove header
    }
    return values.filter(function (v) { return v !== '' && v !== null; });
  }

  var data = {
    names: getValues('InventoryName'),
    sns: getValues('InventorySN'),
    locations: getValues('InventoryLocation'),
    prices: getValues('InventoryPrice')
  };

  var template = HtmlService.createTemplateFromFile('sale');
  template.data = data;

  var html = template.evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'فروش محصول');
}
