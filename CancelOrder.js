function showCancelDialog() {
  var ss = SpreadsheetApp.getActive();
  var getValuesByName = function(name) {
    var range = ss.getRangeByName(name);
    return range ? range.getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;}) : [];
  };
  var orderData = {
    ids: getValuesByName('OrderID'),
    persianIds: getValuesByName('OrderPersianID'),
    names: getValuesByName('OrderName'),
    dates: getValuesByName('OrderDate'),
    persianDates: getValuesByName('OrderPersianDate'),
    sns: getValuesByName('OrderSN'),
    persianSNS: getValuesByName('OrderPersianSN'),
    prices: getValuesByName('OrderPrice'),
    paidPrices: getValuesByName('OrderPaidPrice'),
    skus: getValuesByName('OrderSKU'),
    uniqueCodes: getValuesByName('OrderUniqueCode'),
    sellers: getValuesByName('OrderSeller'),
    locations: getValuesByName('OrderLocation')
  };
  var template = HtmlService.createTemplateFromFile('cancel');
  template.orderData = orderData;
  var html = template.evaluate().setWidth(1200).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'لغو سفارش');
}

function cancelOrders(orderIds) {
  if (!orderIds || !orderIds.length) return;
  var ss = SpreadsheetApp.getActive();
  var idRange = ss.getRangeByName('OrderID');
  var sheet = idRange.getSheet();
  var values = idRange.getValues().map(function(r){ return r[0]; });
  var rowsToDelete = [];
  orderIds.forEach(function(id){
    var idx = values.indexOf(id);
    if (idx > -1) {
      rowsToDelete.push(idRange.getRow() + idx);
    }
  });
  rowsToDelete.sort(function(a,b){ return b - a; });
  rowsToDelete.forEach(function(row){ sheet.deleteRow(row); });
}
