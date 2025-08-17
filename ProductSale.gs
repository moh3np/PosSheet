function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('توسعه')
    .addSubMenu(ui.createMenu('فروش').addItem('فروش محصول', 'showSaleDialog'))
    .addToUi();
}

function showSaleDialog() {
  var html = HtmlService.createHtmlOutputFromFile('sale')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'فروش محصول');
}
