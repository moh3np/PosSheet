function getLastDataRow(range) {
  var sheet = range.getSheet();
  var startRow = range.getRow() + 1;
  var col = range.getColumn();
  var lastRow = sheet.getLastRow();
  var numRows = lastRow - range.getRow();
  if (numRows < 1) return range.getRow();
  var values = sheet.getRange(startRow, col, numRows, 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    var val = values[i][0];
    if (val !== '' && val !== null) {
      return startRow + i;
    }
  }
  return range.getRow();
}
