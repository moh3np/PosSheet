function showUnlabeledProductsDialog() {
  var data = getUnlabeledInventoryData();
  var template = HtmlService.createTemplateFromFile('unlabeled');
  template.inventoryData = data;
  var html = template.evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'محصولات بی‌لیبل');
}

function getUnlabeledInventoryData() {
  var ss = SpreadsheetApp.getActive();
  var inventoryRange = ss.getRangeByName('PosInventory');
  if (!inventoryRange) {
    return { items: [] };
  }
  var sheet = inventoryRange.getSheet();
  var startRow = getDataStartRow(inventoryRange);
  var lastRow = getLastDataRow(inventoryRange);
  if (lastRow < startRow) {
    return { items: [] };
  }
  var numRows = lastRow - startRow + 1;
  var baseCol = inventoryRange.getColumn();

  var nameValues = sheet.getRange(startRow, baseCol, numRows, 1).getValues();
  var skuValues = sheet.getRange(startRow, baseCol + 2, numRows, 1).getValues();

  var uniqueMap = {};
  for (var i = 0; i < numRows; i++) {
    var rawSku = skuValues[i][0] != null ? String(skuValues[i][0]) : '';
    var displaySku = rawSku.trim();
    var normalizedSku = normalizeDigits(displaySku);
    var nameValue = nameValues[i][0] != null ? String(nameValues[i][0]).trim() : '';
    var key = normalizedSku || nameValue.toLowerCase();
    if (!key) {
      continue;
    }
    if (!uniqueMap[key]) {
      uniqueMap[key] = {
        sku: normalizedSku || displaySku,
        name: nameValue,
        count: 0
      };
    } else if (!uniqueMap[key].name && nameValue) {
      uniqueMap[key].name = nameValue;
    }
    uniqueMap[key].count += 1;
    if (!uniqueMap[key].sku && displaySku) {
      uniqueMap[key].sku = normalizedSku || displaySku;
    }
  }

  var toylandSkuMap = {};
  try {
    var externalSs = SpreadsheetApp.openById('1LIR_q1xrpdzcqoBJmNXTO0UJ9dksoBjS7h3Me4PRB1s');
    var toylandRange = externalSs.getRangeByName('ToylandInventory');
    if (toylandRange) {
      var toySheet = toylandRange.getSheet();
      var toyStart = getDataStartRow(toylandRange);
      var toyLast = getLastDataRow(toylandRange);
      if (toyLast >= toyStart) {
        var toyRows = toyLast - toyStart + 1;
        var toySkuValues = toySheet.getRange(toyStart, 3, toyRows, 1).getValues();
        for (var t = 0; t < toySkuValues.length; t++) {
          var skuVal = toySkuValues[t][0];
          if (skuVal !== null && skuVal !== '') {
            var skuStr = String(skuVal).trim();
            var normSku = normalizeDigits(skuStr);
            if (normSku && !toylandSkuMap[normSku]) {
              toylandSkuMap[normSku] = skuStr;
            }
          }
        }
      }
    }
  } catch (err) {
    // ignore external fetch issues; fallback to local SKU
  }

  var items = Object.keys(uniqueMap).map(function(k) { return uniqueMap[k]; });
  items.sort(function(a, b) {
    var nameA = a.name || '';
    var nameB = b.name || '';
    return nameA.localeCompare(nameB, 'fa');
  });
  items = items.map(function(item) {
    var normalizedSku = normalizeDigits(item.sku || '');
    var displaySku = normalizedSku && toylandSkuMap[normalizedSku] ? toylandSkuMap[normalizedSku] : item.sku;
    return {
      sku: item.sku,
      name: item.name,
      count: item.count,
      displaySku: displaySku
    };
  });
  return { items: items };
}

function normalizeDigits(str) {
  if (typeof str !== 'string') {
    str = String(str || '');
  }
  var out = '';
  for (var i = 0; i < str.length; i++) {
    var code = str.charCodeAt(i);
    if (code >= 1776 && code <= 1785) {
      out += String(code - 1776);
    } else if (code >= 1632 && code <= 1641) {
      out += String(code - 1632);
    } else {
      out += str[i];
    }
  }
  return out.trim();
}

function toPersianDigits(str) {
  var digits = ['۰','۱','۲','۳','۴','۵','۶','۷','۸','۹'];
  return String(str).replace(/[0-9]/g, function(d) { return digits[d]; });
}

function findRowMatches(range, value) {
  if (!value) {
    return [];
  }
  var finder = range.createTextFinder(value);
  if (!finder) {
    return [];
  }
  finder.matchCase(false).matchEntireCell(true);
  var cells = finder.findAll();
  if (!cells || !cells.length) {
    return [];
  }
  return cells.map(function(cell) { return cell.getRow(); });
}

function setColumnBoolean(sheet, columnIndex, rows, value) {
  if (!rows || !rows.length) {
    return;
  }
  var a1List = rows.map(function(row) {
    return sheet.getRange(row, columnIndex, 1, 1).getA1Notation();
  });
  sheet.getRangeList(a1List).setValue(value);
}

function updateUnlabeledInventoryStatus(payload) {
  payload = payload || {};
  var desiredCount = Number(payload.desiredCount);
  if (isNaN(desiredCount) || desiredCount < 0) {
    desiredCount = 0;
  }
  var nameKey = payload.name != null ? String(payload.name).trim() : '';
  var skuKey = payload.sku != null ? normalizeDigits(String(payload.sku)) : '';

  var ss = SpreadsheetApp.getActive();
  var inventoryRange = ss.getRangeByName('PosInventory');
  if (!inventoryRange) {
    throw new Error('PosInventory تعریف نشده است.');
  }
  var sheet = inventoryRange.getSheet();
  var startRow = getDataStartRow(inventoryRange);
  var lastRow = getLastDataRow(inventoryRange);
  if (lastRow < startRow) {
    return {
      totalCount: 0,
      desiredCount: 0,
      missingCount: 0,
      message: 'هیچ موجودی در PosInventory یافت نشد.'
    };
  }
  var numRows = lastRow - startRow + 1;
  var baseCol = inventoryRange.getColumn();
  var nameCol = baseCol;
  var skuCol = baseCol + 2;
  var serialCol = baseCol + 4;
  var tlSkuCol = baseCol + 8;

  var rowOrder = [];
  var seenRows = {};

  function appendRows(rows) {
    rows.forEach(function(row) {
      if (row >= startRow && row <= lastRow && !seenRows[row]) {
        seenRows[row] = true;
        rowOrder.push(row);
      }
    });
  }

  if (skuKey) {
    var skuRangeObj = sheet.getRange(startRow, skuCol, numRows, 1);
    appendRows(findRowMatches(skuRangeObj, skuKey));
    var persianSku = toPersianDigits(skuKey);
    if (persianSku !== skuKey) {
      appendRows(findRowMatches(skuRangeObj, persianSku));
    }
  }

  if (nameKey) {
    var nameRangeObj = sheet.getRange(startRow, nameCol, numRows, 1);
    appendRows(findRowMatches(nameRangeObj, nameKey));
  }

  if (!rowOrder.length) {
    return {
      totalCount: 0,
      desiredCount: 0,
      missingCount: 0,
      message: 'محصولی با این مشخصات یافت نشد.'
    };
  }

  rowOrder.sort(function(a, b) { return a - b; });

  var totalCount = rowOrder.length;
  var appliedDesired = Math.min(desiredCount, totalCount);
  var missingCount = totalCount - appliedDesired;
  var keptRows = rowOrder.slice(missingCount);

  var toylandUpdated = 0;
  if (keptRows.length) {
    var skuColumnValues = sheet.getRange(startRow, tlSkuCol, numRows, 1).getValues();
    var tlSkuCounts = {};
    for (var i = 0; i < keptRows.length; i++) {
      var rowIndex = keptRows[i] - startRow;
      if (rowIndex < 0 || rowIndex >= skuColumnValues.length) {
        continue;
      }
      var rawSku = skuColumnValues[rowIndex][0];
      if (rawSku === null || rawSku === '') {
        continue;
      }
      var candidate = String(rawSku).trim();
      if (!/^TL/i.test(candidate)) {
        continue;
      }
      var variantKeys = getToylandSkuKeys(candidate);
      if (!variantKeys.length) {
        continue;
      }
      var primaryKey = variantKeys[0];
      if (!tlSkuCounts[primaryKey]) {
        tlSkuCounts[primaryKey] = {
          count: 0,
          keys: {}
        };
      }
      tlSkuCounts[primaryKey].count += 1;
      variantKeys.forEach(function(vKey) {
        tlSkuCounts[primaryKey].keys[vKey] = true;
      });
    }
    if (Object.keys(tlSkuCounts).length) {
      toylandUpdated = updateToylandInventoryFlags(tlSkuCounts);
    }
  }

  var serialValues = sheet.getRange(startRow, serialCol, numRows, 1).getValues();
  var keptSerials = [];
  // Reset external label flag for rows we keep in inventory.
  if (keptRows.length) {
    keptRows.forEach(function(row) {
      var serial = serialValues[row - startRow][0];
      if (serial !== null && serial !== '') {
        keptSerials.push(String(serial).trim());
      }
    });
  }

  var remoteUpdated = 0;
  return {
    totalCount: totalCount,
    desiredCount: appliedDesired,
    missingCount: missingCount,
    remoteUpdated: remoteUpdated,
    toylandUpdated: toylandUpdated
  };
}

function updateToylandInventoryFlags(tlSkuCounts) {
  if (!tlSkuCounts) {
    return 0;
  }
  var skuKeys = Object.keys(tlSkuCounts);
  if (!skuKeys.length) {
    return 0;
  }

  var externalSs;
  try {
    externalSs = SpreadsheetApp.openById('1LIR_q1xrpdzcqoBJmNXTO0UJ9dksoBjS7h3Me4PRB1s');
  } catch (err) {
    throw new Error('عدم دسترسی به شیت موجودی بیرونی: ' + err);
  }
  var toylandRange = externalSs.getRangeByName('ToylandInventory');
  if (!toylandRange) {
    throw new Error('ToylandInventory در شیت بیرونی پیدا نشد.');
  }
  var toylandSheet = toylandRange.getSheet();
  var toyStartRow = getDataStartRow(toylandRange);
  var toyLastRow = getLastDataRow(toylandRange);
  if (toyLastRow < toyStartRow) {
    return 0;
  }
  var toyNumRows = toyLastRow - toyStartRow + 1;
  var toyBaseCol = toylandRange.getColumn();
  var toySkuCol = toyBaseCol + 2;
  var toyFlagCol = toyBaseCol + 8;
  var toyReadyCol = toyBaseCol + 9;

  var toySkuValues = toylandSheet.getRange(toyStartRow, toySkuCol, toyNumRows, 1).getValues();
  var rowLookup = {};
  for (var i = 0; i < toySkuValues.length; i++) {
    var rawSku = toySkuValues[i][0];
    if (rawSku === null || rawSku === '') {
      continue;
    }
    var trimmedSku = String(rawSku).trim();
    if (!trimmedSku) {
      continue;
    }
    var variants = getToylandSkuKeys(trimmedSku);
    if (!variants.length) {
      continue;
    }
    var rowNumber = toyStartRow + i;
    variants.forEach(function(key) {
      if (!rowLookup[key]) {
        rowLookup[key] = [];
      }
      rowLookup[key].push(rowNumber);
    });
  }

  var rowsToFlag = [];
  var seenRowsGlobal = {};
  for (var k = 0; k < skuKeys.length; k++) {
    var primaryKey = skuKeys[k];
    var skuInfo = tlSkuCounts[primaryKey];
    if (!skuInfo || !skuInfo.count) {
      continue;
    }
    var matches = rowLookup[primaryKey];
    if ((!matches || !matches.length) && skuInfo.keys) {
      var variantKeys = Object.keys(skuInfo.keys);
      for (var v = 0; v < variantKeys.length; v++) {
        var altMatches = rowLookup[variantKeys[v]];
        if (altMatches && altMatches.length) {
          matches = altMatches;
          break;
        }
      }
    }
    if (!matches || !matches.length) {
      continue;
    }
    var seenRows = {};
    var uniqueMatches = [];
    for (var m = 0; m < matches.length; m++) {
      var rowId = matches[m];
      if (!seenRows[rowId]) {
        seenRows[rowId] = true;
        uniqueMatches.push(rowId);
      }
    }
    var limit = Math.min(skuInfo.count, uniqueMatches.length);
    for (var idx = 0; idx < limit; idx++) {
      var rowNumber = uniqueMatches[idx];
      if (!seenRowsGlobal[rowNumber]) {
        seenRowsGlobal[rowNumber] = true;
        rowsToFlag.push(rowNumber);
      }
    }
  }
  if (!rowsToFlag.length) {
    return 0;
  }
  var flagA1 = [];
  var readyA1 = [];
  rowsToFlag.forEach(function(row) {
    flagA1.push(toylandSheet.getRange(row, toyFlagCol, 1, 1).getA1Notation());
    readyA1.push(toylandSheet.getRange(row, toyReadyCol, 1, 1).getA1Notation());
  });
  toylandSheet.getRangeList(flagA1).setValue(false);
  toylandSheet.getRangeList(readyA1).setValue(true);
  return rowsToFlag.length;
}

function findSerialCells(range, serial) {
  if (!serial) {
    return [];
  }
  var results = [];
  var finder = range.createTextFinder(serial);
  if (finder) {
    finder.matchCase(false).matchEntireCell(true);
    var found = finder.findAll();
    if (found && found.length) {
      results = results.concat(found);
    }
  }
  var persian = toPersianDigits(serial);
  if (persian !== serial) {
    var persianFinder = range.createTextFinder(persian);
    if (persianFinder) {
      persianFinder.matchCase(false).matchEntireCell(true);
      var alt = persianFinder.findAll();
      if (alt && alt.length) {
        results = results.concat(alt);
      }
    }
  }
  return results;
}

function getToylandSkuKeys(value) {
  var keys = [];
  var seen = {};
  if (value === null || value === undefined) {
    return keys;
  }
  var raw = String(value).trim();
  if (!raw) {
    return keys;
  }

  function addKey(val) {
    if (!val) {
      return;
    }
    var key = String(val).trim();
    if (!key) {
      return;
    }
    key = key.toUpperCase();
    if (!seen[key]) {
      seen[key] = true;
      keys.push(key);
    }
  }

  var normalized = normalizeDigits(raw).trim();
  var candidates = [];
  candidates.push(raw);
  if (normalized && normalized !== raw) {
    candidates.push(normalized);
  }

  candidates.forEach(function(candidate) {
    if (!candidate) {
      return;
    }
    var upper = candidate.toUpperCase().trim();
    if (!upper) {
      return;
    }
    var noPrefix = upper.replace(/^TL[\s\-]*/i, '').trim();
    var digitsOnly = upper.replace(/\D+/g, '').trim();
    var digitsNoPrefix = noPrefix ? noPrefix.replace(/\D+/g, '').trim() : '';

    addKey(digitsNoPrefix);
    addKey(digitsOnly);
    addKey(noPrefix);
    addKey(upper);
  });

  return keys;
}
