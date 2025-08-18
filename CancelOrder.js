function showCancelDialog() {
  var ss = SpreadsheetApp.getActive();
  var orderData = {
    ids: ss.getRangeByName('OrderID').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;}),
    persianIds: ss.getRangeByName('OrderPersianID').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;}),
    names: ss.getRangeByName('OrderName').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;}),
    dates: ss.getRangeByName('OrderDate').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;}),
    sns: ss.getRangeByName('OrderSN').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;}),
    persianSNS: ss.getRangeByName('OrderPersianSN').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;}),
    prices: ss.getRangeByName('OrderPrice').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;}),
    paidPrices: ss.getRangeByName('OrderPaidPrice').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;}),
    skus: ss.getRangeByName('OrderSKU').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;}),
    uniqueCodes: ss.getRangeByName('OrderUniqueCode').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;}),
    sellers: ss.getRangeByName('OrderSeller').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;}),
    locations: ss.getRangeByName('OrderLocation').getValues().map(function(r){return r[0];}).filter(function(v,i){return v && i>0;})
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
