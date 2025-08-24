function getDataStartRow(range) {
  var sheet = range.getSheet();
  var frozenRows = sheet.getFrozenRows();
  return range.getRow() <= frozenRows ? frozenRows + 1 : range.getRow();
}

function getLastDataRow(range) {
  var sheet = range.getSheet();
  var startRow = getDataStartRow(range);
  var startCol = range.getColumn();
  var numRows = range.getNumRows() - (startRow - range.getRow());
  var numCols = range.getNumColumns();
  if (numRows < 1) return startRow - 1;
  var values = sheet.getRange(startRow, startCol, numRows, numCols).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    for (var j = 0; j < numCols; j++) {
      var val = values[i][j];
      if (val !== '' && val !== null) {
        return startRow + i;
      }
    }
  }
  return startRow - 1;
}

function getPersianDateTime() {
  var parts = Utilities.formatDate(new Date(), 'Asia/Tehran', 'yyyy-M-d-HH:mm:ss').split('-');
  var gYear = Number(parts[0]);
  var gMonth = Number(parts[1]);
  var gDay = Number(parts[2]);
  var time = parts[3];
  var j = gregorianToJalali(gYear, gMonth, gDay);
  var jy = j[0];
  var jm = ('0' + j[1]).slice(-2);
  var jd = ('0' + j[2]).slice(-2);
  return jy + '/' + jm + '/' + jd + ' ' + time;
}

function gregorianToJalali(gy, gm, gd) {
  var g_d_m = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334];
  var jy;
  if (gy > 1600) {
    jy = 979;
    gy -= 1600;
  } else {
    jy = 0;
    gy -= 621;
  }
  var gy2 = gm > 2 ? gy + 1 : gy;
  var days = (365 * gy) + Math.floor((gy2 + 3) / 4) - Math.floor((gy2 + 99) / 100) + Math.floor((gy2 + 399) / 400) - 80 + gd + g_d_m[gm - 1];
  jy += 33 * Math.floor(days / 12053);
  days %= 12053;
  jy += 4 * Math.floor(days / 1461);
  days %= 1461;
  if (days > 365) {
    jy += Math.floor((days - 1) / 365);
    days = (days - 1) % 365;
  }
  var jm = (days < 186) ? 1 + Math.floor(days / 31) : 7 + Math.floor((days - 186) / 30);
  var jd = 1 + ((days < 186) ? (days % 31) : ((days - 186) % 30));
  return [jy, jm, jd];
}
