/**
 * Records an order in the first available row of the sales sheet.
 *
 * @param {Object[]} items - Array of sale items with name, sku, serial, price and paid fields.
 */
function submitOrder(items) {
  if (!items || items.length === 0) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Sales') || ss.getSheets()[0];

  const values = items.map(item => [item.name, item.sku, item.serial, item.price, item.paid]);

  const startRow = sheet.getLastRow() + 1;
  const neededRows = startRow + values.length - 1;
  if (neededRows > sheet.getMaxRows()) {
    sheet.insertRows(sheet.getMaxRows() + 1, neededRows - sheet.getMaxRows());
  }

  sheet.getRange(startRow, 1, values.length, values[0].length).setValues(values);
}
