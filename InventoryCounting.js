function showInventoryCountingDialog() {
  var template = HtmlService.createTemplateFromFile('inventory-count');
  var lookupData = getInventorySerialMetadata();
  template.inventoryLookupJson = JSON.stringify(lookupData).replace(/</g, '\\u003c');
  var html = template.evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'انبارگردانی تویلند');
}

function getInventorySerialMetadata() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(Config.sheets.tolyandSheet);
  } catch (err) {
    return [];
  }
  if (!ss) {
    return [];
  }
  var inventoryRange = ss.getRangeByName(Config.namedRanges.toylandInventory);
  if (!inventoryRange) {
    return [];
  }

  var sheet = inventoryRange.getSheet();
  var startRow = getDataStartRow(inventoryRange);
  var lastRow = getLastDataRow(inventoryRange);
  if (lastRow < startRow) {
    return [];
  }

  var numRows = lastRow - startRow + 1;
  var baseCol = inventoryRange.getColumn();
  var nameValues = sheet.getRange(startRow, baseCol, numRows, 1).getValues();
  var skuValues = sheet.getRange(startRow, baseCol + 2, numRows, 1).getValues();
  var serialValues = sheet.getRange(startRow, baseCol + 4, numRows, 1).getValues();

  var lookup = [];
  var seen = {};
  for (var i = 0; i < numRows; i++) {
    var rawSerial = serialValues[i][0];
    if (rawSerial === null || rawSerial === '') {
      continue;
    }
    var serial = String(rawSerial).trim();
    if (!serial) {
      continue;
    }
    var normalized = normalizeDigits(serial);
    if (!normalized || seen[normalized]) {
      continue;
    }
    seen[normalized] = true;

    var name = nameValues[i][0];
    var sku = skuValues[i][0];

    lookup.push({
      serial: serial,
      serialPersian: typeof toPersianDigits === 'function' ? toPersianDigits(serial) : '',
      normalized: normalized,
      name: name != null ? String(name).trim() : '',
      sku: sku != null ? String(sku).trim() : ''
    });
  }
  return lookup;
}

function submitInventoryCounting(serials) {
  serials = Array.isArray(serials) ? serials : [];
  var uniqueEntries = [];
  var duplicates = [];
  var seen = {};

  serials.forEach(function(item) {
    if (item === null || item === undefined) {
      return;
    }
    var raw = String(item).trim();
    if (!raw) {
      return;
    }
    var normalized = normalizeDigits(raw);
    if (!normalized) {
      return;
    }
    if (seen[normalized]) {
      duplicates.push(raw);
      return;
    }
    seen[normalized] = true;
    uniqueEntries.push({
      original: raw,
      normalized: normalized
    });
  });

  var baseResult = {
    totalInput: serials.length,
    uniqueCount: uniqueEntries.length,
    duplicates: duplicates,
    missingInPos: [],
    unrouted: [],
    br: { requested: 0, updated: 0, missingExternal: [] },
    tl: { requested: uniqueEntries.length, updated: 0, missingExternal: [] }
  };

  if (!uniqueEntries.length) {
    return baseResult;
  }

  var tlResult = updateInventoryFlagForSerials({
    spreadsheetId: Config.sheets.tolyandSheet,
    rangeName: Config.namedRanges.toylandInventory,
    columns: [
      { offset: 8, value: true },
      { offset: 9, value: 'Available' }
    ]
  }, uniqueEntries);

  baseResult.tl.updated = tlResult.updated;
  baseResult.tl.missingExternal = tlResult.missing;
  baseResult.missingInPos = tlResult.missing.slice();
  return baseResult;
}

function updateInventoryFlagForSerials(config, entries) {
  if (!entries || !entries.length) {
    return { updated: 0, missing: [] };
  }
  var ss;
  try {
    ss = SpreadsheetApp.openById(config.spreadsheetId);
  } catch (err) {
    throw new Error('عدم دسترسی به شیت مورد نظر: ' + err);
  }
  var range = ss.getRangeByName(config.rangeName);
  if (!range) {
    throw new Error(config.rangeName + ' در شیت خارجی پیدا نشد.');
  }
  var sheet = range.getSheet();
  var startRow = getDataStartRow(range);
  var lastRow = getLastDataRow(range);
  if (lastRow < startRow) {
    return {
      updated: 0,
      missing: entries.map(function(entry) { return entry.original; })
    };
  }

  var numRows = lastRow - startRow + 1;
  var baseCol = range.getColumn();
  var serialCol = baseCol + 4;
  var defaultOffset = typeof config.columnOffset === 'number' ? config.columnOffset : 9;
  var defaultValue = Object.prototype.hasOwnProperty.call(config, 'value') ? config.value : true;

  var columnConfigs = Array.isArray(config.columns) && config.columns.length ? config.columns : [{ offset: defaultOffset, value: defaultValue }];
  var normalizedColumns = columnConfigs.map(function(col) {
    var offset = typeof col.offset === 'number' ? col.offset : defaultOffset;
    var value = Object.prototype.hasOwnProperty.call(col, 'value') ? col.value : defaultValue;
    return { offset: offset, value: value };
  }).filter(function(col) {
    return typeof col.offset === 'number' && isFinite(col.offset);
  });

  if (!normalizedColumns.length) {
    normalizedColumns.push({ offset: defaultOffset, value: defaultValue });
  }

  var serialValues = sheet.getRange(startRow, serialCol, numRows, 1).getValues();
  var serialMap = {};
  for (var i = 0; i < numRows; i++) {
    var value = serialValues[i][0];
    if (value === null || value === '') {
      continue;
    }
    var norm = normalizeDigits(String(value).trim());
    if (!norm) {
      continue;
    }
    if (!serialMap[norm]) {
      serialMap[norm] = [];
    }
    serialMap[norm].push({
      row: startRow + i,
      used: false
    });
  }

  var chosenRows = [];
  var chosenRowSet = {};
  var missing = [];
  entries.forEach(function(entry) {
    var rows = serialMap[entry.normalized];
    if (!rows || !rows.length) {
      missing.push(entry.original);
      return;
    }

    var chosen = null;
    for (var idx = 0; idx < rows.length; idx++) {
      if (!rows[idx].used) {
        chosen = rows[idx];
        rows[idx].used = true;
        break;
      }
    }

    if (!chosen) {
      missing.push(entry.original);
      return;
    }

    if (!chosenRowSet[chosen.row]) {
      chosenRowSet[chosen.row] = true;
      chosenRows.push(chosen.row);
    }
  });

  normalizedColumns.forEach(function(column) {
    var offset = column.offset;
    var value = Object.prototype.hasOwnProperty.call(column, 'value') ? column.value : defaultValue;
    var targetCol = baseCol + offset;
    if (!targetCol || targetCol < 1) {
      return;
    }
    var addressSet = {};
    var addresses = [];
    for (var i = 0; i < chosenRows.length; i++) {
      var row = chosenRows[i];
      var address = sheet.getRange(row, targetCol, 1, 1).getA1Notation();
      if (!addressSet[address]) {
        addressSet[address] = true;
        addresses.push(address);
      }
    }
    if (addresses.length) {
      sheet.getRangeList(addresses).setValue(value);
    }
  });

  return {
    updated: chosenRows.length,
    missing: missing
  };
}
