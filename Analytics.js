var PosSheetInsights = (function(namespace, global) {
  function getPosOrdersRangeName() {
    if (typeof Config === 'undefined' || !Config || !Config.namedRanges || !Config.namedRanges.posOrders) {
      throw new Error('Config.namedRanges.posOrders تعریف نشده است.');
    }
    return Config.namedRanges.posOrders;
  }

  var CACHE_KEY = 'PSI_ROWS_V1';
  var CACHE_META_KEY = 'PSI_CACHE_TIMESTAMP';
  var CACHE_TTL_SECONDS = 21600; // 6 hours
  var PAGE_SIZE = 50;
  var NULL_LABEL = '\u2014';
  var MIN_LOW_STOCK_THRESHOLD = 3;
  var RESTOCK_SAFETY_DAYS = 7;
  var MIN_RESTOCK_SALES = 1;
  var RESTOCK_TARGET_DAYS = 14;

  namespace = namespace || {};
  if (global && typeof global === 'object') {
    global.PosSheetInsights = namespace;
  }

  namespace.showSidebar = function() {
    var template = HtmlService.createTemplateFromFile('analytics');
    template.viewMode = 'sidebar';
    template.namespace = 'PosSheetInsights';
    var html = template.evaluate();
    html.setTitle('تحلیل فروش');
    SpreadsheetApp.getUi().showSidebar(html);
  };

  namespace.showDialog = function() {
    var template = HtmlService.createTemplateFromFile('analytics');
    template.viewMode = 'dialog';
    template.namespace = 'PosSheetInsights';
    var html = template.evaluate();
    html.setWidth(1200);
    html.setHeight(720);
    SpreadsheetApp.getUi().showModalDialog(html, 'نمایش گسترده');
  };

  namespace.getInitialState = function() {
    var rows = getRows(false);
    var options = buildFilterOptions(rows);
    var defaults = buildDefaultFilters(rows);
    var report = buildReport(rows, defaults, 1);
    return {
      options: options,
      defaults: defaults,
      report: report,
      lastUpdated: getLastUpdated()
    };
  };

  namespace.getFilterOptions = function() {
    var rows = getRows(false);
    return {
      options: buildFilterOptions(rows),
      lastUpdated: getLastUpdated()
    };
  };

  namespace.computeReport = function(filters) {
    var rows = getRows(false);
    var cleanFilters = sanitizeFilters(filters || {});
    var report = buildReport(rows, cleanFilters, cleanFilters.page || 1);
    report.lastUpdated = getLastUpdated();
    return report;
  };

  namespace.getChartData = function(filters) {
    var rows = getRows(false);
    var cleanFilters = sanitizeFilters(filters || {});
    var dataset = prepareDataset(rows, cleanFilters);
    return buildCharts(dataset, cleanFilters);
  };

  namespace.exportReport = function(filters) {
    var rows = getRows(false);
    var cleanFilters = sanitizeFilters(filters || {});
    var exportData = prepareDataset(rows, cleanFilters);
    var sheetName = exportToSheet(exportData, cleanFilters);
    return {
      sheetName: sheetName,
      timestamp: new Date().toISOString()
    };
  };

  namespace.refreshCache = function() {
    var rows = getRows(true);
    return {
      refreshed: true,
      totalRows: rows.length,
      timestamp: new Date().toISOString()
    };
  };

  namespace.setupTrigger = function() {
    var handlerName = 'PosSheetInsights_runScheduledRefresh';
    var triggers = ScriptApp.getProjectTriggers();
    var existing = null;
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction && triggers[i].getHandlerFunction() === handlerName) {
        existing = triggers[i];
        break;
      }
    }
    if (existing) {
      ScriptApp.deleteTrigger(existing);
    }
    ScriptApp.newTrigger(handlerName)
      .timeBased()
      .everyHours(6)
      .create();
    return {
      created: true,
      intervalHours: 6,
      timestamp: new Date().toISOString()
    };
  };

  namespace.runScheduledRefresh = function() {
    try {
      getRows(true);
    } catch (err) {
      console.error('Scheduled refresh failed:', err);
    }
  };

  namespace.ensureMenu = function() {
    if (typeof buildPosMenus === 'function') {
      buildPosMenus();
      return;
    }
    var ui = SpreadsheetApp.getUi();
    var salesMenu = ui.createMenu('فروش');
    salesMenu.addItem('خرده فروشی', 'showSaleDialog');
    salesMenu.addItem('لغو سفارش', 'showCancelDialog');
    salesMenu.addToUi();

    ui.createMenu('محصولات')
      .addItem('محصولات بی‌لیبل', 'showUnlabeledProductsDialog')
      .addToUi();

    var analysisMenu = ui.createMenu('تحلیل فروش');
    analysisMenu.addItem('تحلیل فروش (Sidebar)', 'PosSheetInsights_showSidebar');
    analysisMenu.addItem('نمایش گسترده', 'PosSheetInsights_showDialog');
    analysisMenu.addItem('به‌روزرسانی داده‌ها', 'PosSheetInsights_refreshCache');
    analysisMenu.addItem('تنظیم زمان‌بندی', 'PosSheetInsights_setupTrigger');
    analysisMenu.addToUi();
  };

  function getRows(forceRefresh) {
    var cache = CacheService.getScriptCache();
    if (!forceRefresh) {
      var cached = cache.get(CACHE_KEY);
      if (cached) {
        try {
          return JSON.parse(cached);
        } catch (err) {
          // Ignore parse issues, rebuild cache
        }
      }
    }
    var ss = SpreadsheetApp.getActive();
    var posOrdersRangeName = getPosOrdersRangeName();
    var namedRange = ss.getRangeByName(posOrdersRangeName);
    var sheet;
    if (namedRange) {
      sheet = namedRange.getSheet();
    } else {
      sheet = ss.getSheetByName(posOrdersRangeName);
      if (!sheet) {
        setLastUpdated(null);
        return [];
      }
      namedRange = sheet.getDataRange();
    }
    var startRow = getDataStartRow(namedRange);
    var lastRow = getLastDataRow(namedRange);
    if (lastRow < startRow) {
      setLastUpdated(new Date());
      cacheRows(cache, []);
      return [];
    }
    var numRows = lastRow - startRow + 1;
    var numCols = namedRange.getNumColumns();
    var dataRange = sheet.getRange(startRow, namedRange.getColumn(), numRows, numCols);
    var values = dataRange.getValues();
    var displays = dataRange.getDisplayValues();
    var costLookup = getSkuCostMap(ss);
    var rows = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var displayRow = displays[i];
      var orderIdRaw = normalizeDigits(displayRow[0] != null ? String(displayRow[0]) : String(row[0]));
      var orderId = orderIdRaw ? orderIdRaw.trim() : '';
      var productName = (displayRow[1] || row[1] || '').trim();
      var sku = normalizeDigits(displayRow[2] || row[2] || '').trim();
      var sn = normalizeDigits(displayRow[3] || row[3] || '').trim();
      var orderDateJalali = (displayRow[4] || row[4] || '').trim();
      var orderDate = parseJalaliLatin(orderDateJalali);
      var priceList = parseMoney(displayRow[5] || row[5]);
      var pricePaid = parseMoney(displayRow[6] || row[6]);
      var location = (displayRow[7] || row[7] || '').trim();
      var seller = (displayRow[8] || row[8] || '').trim();
      var brand = (displayRow[9] || row[9] || '').trim();
      var uniqueCode = (displayRow[10] || row[10] || '').trim();
      var canceledVal = row[11];
      var canceled = Boolean(canceledVal === true || String(canceledVal).toLowerCase() === 'true');
      var skuPlus = normalizeDigits(displayRow[12] || row[12] || '').trim();
      var discountAmount = priceList > pricePaid ? priceList - pricePaid : Math.max(0, priceList - pricePaid);
      var discountRate = priceList ? discountAmount / priceList : null;
      var brandOrDash = brand || NULL_LABEL;
      var locationOrDash = location || NULL_LABEL;
      var sellerOrDash = seller || NULL_LABEL;
      var amountKey = pricePaid;
      var jalaliShort = orderDateJalali ? orderDateJalali.split(' ')[0] : '';
      var searchIndex = [orderId, sku, skuPlus, sn, productName, sellerOrDash, brandOrDash, uniqueCode].join(' ').toLowerCase();
      var normalizedSearch = normalizeDigits(searchIndex).toLowerCase();
      var cost = null;
      if (costLookup && sku) {
        cost = costLookup[sku] != null ? costLookup[sku] : null;
      }
      rows.push({
        orderId: orderId,
        productName: productName,
        sku: sku,
        skuPlus: skuPlus,
        sn: sn,
        orderDate: orderDate,
        orderDateJalali: orderDateJalali,
        orderDateJalaliShort: jalaliShort,
        priceList: priceList,
        pricePaid: pricePaid,
        discountAmount: discountAmount,
        discountRate: discountRate,
        location: locationOrDash,
        seller: sellerOrDash,
        brand: brandOrDash,
        uniqueCode: uniqueCode,
        canceled: canceled,
        cost: cost,
        searchIndex: normalizedSearch,
        rawSearch: searchIndex
      });
    }
    cacheRows(cache, rows);
    setLastUpdated(new Date());
    return rows;
  }

  function cacheRows(cache, rows) {
    var payload = JSON.stringify(rows);
    if (payload.length < 90000) {
      cache.put(CACHE_KEY, payload, CACHE_TTL_SECONDS);
    } else {
      cache.remove(CACHE_KEY);
    }
  }

  function setLastUpdated(date) {
    var props = PropertiesService.getScriptProperties();
    if (date) {
      props.setProperty(CACHE_META_KEY, date.toISOString());
    } else {
      props.deleteProperty(CACHE_META_KEY);
    }
  }

  function getLastUpdated() {
    var props = PropertiesService.getScriptProperties();
    var ts = props.getProperty(CACHE_META_KEY);
    return ts || null;
  }

  function sanitizeFilters(filters) {
    var clean = {};
    clean.dateFrom = typeof filters.dateFrom === 'string' ? filters.dateFrom.trim() : '';
    clean.dateTo = typeof filters.dateTo === 'string' ? filters.dateTo.trim() : '';
    clean.sellers = Array.isArray(filters.sellers) ? filters.sellers.filter(Boolean) : [];
    clean.brands = Array.isArray(filters.brands) ? filters.brands.filter(Boolean) : [];
    clean.locations = Array.isArray(filters.locations) ? filters.locations.filter(Boolean) : [];
    clean.search = typeof filters.search === 'string' ? filters.search.trim() : '';
    clean.includeCanceled = Boolean(filters.includeCanceled);
    clean.amountType = filters.amountType === 'priceList' ? 'priceList' : 'pricePaid';
    clean.groupBy = filters.groupBy || 'none';
    clean.page = Math.max(1, parseInt(filters.page, 10) || 1);
    return clean;
  }

  function buildFilterOptions(rows) {
    var sellers = {};
    var brands = {};
    var locations = {};
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].seller) sellers[rows[i].seller] = true;
      if (rows[i].brand) brands[rows[i].brand] = true;
      if (rows[i].location) locations[rows[i].location] = true;
    }
    var sellerList = Object.keys(sellers).sort(localeSort);
    if (sellerList.indexOf(NULL_LABEL) === -1) {
      sellerList.unshift(NULL_LABEL);
    }
    var brandList = Object.keys(brands).sort(localeSort);
    if (brandList.indexOf(NULL_LABEL) === -1) {
      brandList.unshift(NULL_LABEL);
    }
    var locationList = Object.keys(locations).sort(localeSort);
    if (locationList.indexOf(NULL_LABEL) === -1) {
      locationList.unshift(NULL_LABEL);
    }
    return {
      sellers: sellerList,
      brands: brandList,
      locations: locationList
    };
  }

  function buildDefaultFilters(rows) {
    var now = new Date();
    var toDate = now;
    var fromDate = new Date(toDate.getTime() - 29 * 86400000);
    if (rows.length) {
      var maxDate = null;
      for (var i = 0; i < rows.length; i++) {
        if (rows[i].orderDate instanceof Date && !isNaN(rows[i].orderDate.getTime())) {
          if (!maxDate || rows[i].orderDate > maxDate) {
            maxDate = rows[i].orderDate;
          }
        }
      }
      if (maxDate) {
        toDate = maxDate;
        fromDate = new Date(maxDate.getTime() - 29 * 86400000);
      }
    }
    return {
      dateFrom: toJalaliString(fromDate),
      dateTo: toJalaliString(toDate),
      sellers: [],
      brands: [],
      locations: [],
      search: '',
      includeCanceled: false,
      amountType: 'pricePaid',
      groupBy: 'none',
      page: 1
    };
  }

  function prepareDataset(rows, filters) {
    var fromDate = filters.dateFrom ? parseJalaliLatin(filters.dateFrom + ' 00:00:00') : null;
    var toDate = filters.dateTo ? parseJalaliLatin(filters.dateTo + ' 23:59:59') : null;
    var sellerSet = filters.sellers && filters.sellers.length ? makeLookup(filters.sellers) : null;
    var brandSet = filters.brands && filters.brands.length ? makeLookup(filters.brands) : null;
    var locationSet = filters.locations && filters.locations.length ? makeLookup(filters.locations) : null;
    var search = normalizeDigits(filters.search || '').toLowerCase();
    var filtered = [];
    var included = [];
    var canceled = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (fromDate && (!r.orderDate || r.orderDate < fromDate)) {
        continue;
      }
      if (toDate && (!r.orderDate || r.orderDate > toDate)) {
        continue;
      }
      if (sellerSet && !sellerSet[r.seller || NULL_LABEL]) {
        continue;
      }
      if (brandSet && !brandSet[r.brand || NULL_LABEL]) {
        continue;
      }
      if (locationSet && !locationSet[r.location || NULL_LABEL]) {
        continue;
      }
      if (search) {
        var rawSearch = normalizeDigits(r.rawSearch).toLowerCase();
        if (r.searchIndex.indexOf(search) === -1 && rawSearch.indexOf(search) === -1) {
          continue;
        }
      }
      filtered.push(r);
      if (r.canceled) {
        canceled.push(r);
      }
      if (filters.includeCanceled || !r.canceled) {
        included.push(r);
      }
    }
    return {
      filtered: filtered,
      included: included,
      canceled: canceled,
      range: {
        from: fromDate,
        to: toDate
      }
    };
  }

  function buildReport(rows, filters, pageNumber) {
    var dataset = prepareDataset(rows, filters);
    var included = dataset.included;
    var canceled = dataset.canceled;
    var amountKey = filters.amountType === 'priceList' ? 'priceList' : 'pricePaid';
    var gross = 0;
    var net = 0;
    var discount = 0;
    var orderSet = {};
    for (var i = 0; i < included.length; i++) {
      gross += included[i].priceList || 0;
      net += included[i][amountKey] || 0;
      discount += included[i].discountAmount || 0;
      if (included[i].orderId) {
        orderSet[included[i].orderId] = true;
      }
    }
    var orderCount = Object.keys(orderSet).length;
    var returnAmount = 0;
    var returnOrderSet = {};
    for (var c = 0; c < canceled.length; c++) {
      returnAmount += canceled[c][amountKey] || 0;
      if (canceled[c].orderId) {
        returnOrderSet[canceled[c].orderId] = true;
      }
    }
    var returnOrderCount = Object.keys(returnOrderSet).length;
    var kpis = {
      orderCount: orderCount,
      itemCount: included.length,
      grossAmount: gross,
      netAmount: net,
      averageOrderValue: orderCount ? net / orderCount : 0,
      discountTotal: discount,
      discountRate: gross ? discount / gross : null,
      returnAmount: returnAmount,
      returnCount: canceled.length,
      returnOrderCount: returnOrderCount
    };

    var table = buildTable(dataset, filters, amountKey);
    var pagination = paginateTable(table.rows, pageNumber || 1);
    var inventorySnapshot = getInventorySnapshot();
    var insights = buildProductInsights(dataset, filters, amountKey, inventorySnapshot);

    return {
      filters: filters,
      kpis: kpis,
      table: {
        headers: table.headers,
        rows: pagination.rows,
        isGrouped: table.isGrouped,
        valueTypes: table.valueTypes
      },
      pagination: {
        page: pagination.page,
        totalPages: pagination.totalPages,
        pageSize: PAGE_SIZE,
        totalRows: table.rows.length
      },
      datasetInfo: {
        filteredCount: dataset.filtered.length,
        includedCount: included.length,
        canceledCount: canceled.length
      },
      insights: insights
    };
  }

  function buildTable(dataset, filters, amountKey) {
    if (filters.groupBy && filters.groupBy !== 'none') {
      var grouped = groupRows(dataset.included, filters.groupBy, amountKey);
      return {
        headers: ['کلید گروه', 'تعداد آیتم', 'تعداد سفارش یکتا', 'مبلغ ناخالص', 'مبلغ خالص', 'مجموع تخفیف', 'AOV'],
        valueTypes: ['text', 'number', 'number', 'currency', 'currency', 'currency', 'currency'],
        rows: grouped,
        isGrouped: true
      };
    }
    var rows = [];
    for (var i = 0; i < dataset.included.length; i++) {
      var r = dataset.included[i];
      rows.push([
        r.orderDateJalaliShort || formatDateForDisplay(r.orderDate),
        r.orderId || '',
        r.sku || '',
        r.productName || '',
        r.seller || NULL_LABEL,
        r.brand || NULL_LABEL,
        r.location || NULL_LABEL,
        r.priceList || 0,
        r[amountKey] || 0,
        r.discountAmount || 0,
        r.canceled ? 'بله' : 'خیر'
      ]);
    }
    return {
      headers: ['تاریخ شمسی', 'شماره سفارش', 'SKU', 'نام محصول', 'فروشنده', 'برند', 'موقعیت', 'قیمت محصول', 'قیمت با تخفیف', 'تخفیف', 'لغو شده'],
      valueTypes: ['text', 'text', 'text', 'text', 'text', 'text', 'text', 'currency', 'currency', 'currency', 'text'],
      rows: rows,
      isGrouped: false
    };
  }

  function groupRows(rows, groupBy, amountKey) {
    var groups = {};
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var key = getGroupKey(r, groupBy);
      if (!groups[key]) {
        groups[key] = {
          key: key,
          itemCount: 0,
          orderSet: {},
          gross: 0,
          net: 0,
          discount: 0
        };
      }
      var g = groups[key];
      g.itemCount += 1;
      g.gross += r.priceList || 0;
      g.net += r[amountKey] || 0;
      g.discount += r.discountAmount || 0;
      if (r.orderId) {
        g.orderSet[r.orderId] = true;
      }
    }
    var result = [];
    for (var key in groups) {
      if (groups.hasOwnProperty(key)) {
        var grp = groups[key];
        var orderCount = Object.keys(grp.orderSet).length;
        result.push([
          key,
          grp.itemCount,
          orderCount,
          grp.gross,
          grp.net,
          grp.discount,
          orderCount ? grp.net / orderCount : 0
        ]);
      }
    }
    result.sort(function(a, b) {
      if (typeof a[0] === 'number' && typeof b[0] === 'number') {
        return a[0] - b[0];
      }
      var aKey = String(a[0]);
      var bKey = String(b[0]);
      return aKey.localeCompare(bKey, 'fa');
    });
    return result;
  }

  function paginateTable(rows, pageNumber) {
    var total = rows.length;
    var totalPages = total ? Math.ceil(total / PAGE_SIZE) : 1;
    var page = Math.min(Math.max(1, pageNumber), totalPages);
    var start = (page - 1) * PAGE_SIZE;
    var end = Math.min(start + PAGE_SIZE, total);
    return {
      rows: rows.slice(start, end),
      page: page,
      totalPages: totalPages
    };
  }

  function buildCharts(dataset, filters) {
    var amountKey = filters.amountType === 'priceList' ? 'priceList' : 'pricePaid';
    var byDay = {};
    var byProduct = {};
    var bySeller = {};
    var byBrand = {};
    for (var i = 0; i < dataset.included.length; i++) {
      var r = dataset.included[i];
      var keyDay = r.orderDateJalaliShort || formatDateForDisplay(r.orderDate);
      accumulate(byDay, keyDay, r[amountKey] || 0);
      var productKey = r.productName || r.sku || NULL_LABEL;
      accumulate(byProduct, productKey, r[amountKey] || 0);
      accumulate(bySeller, r.seller || NULL_LABEL, r[amountKey] || 0);
      accumulate(byBrand, r.brand || NULL_LABEL, r[amountKey] || 0);
    }
    var daySeries = toChartSeries(byDay, true);
    var productSeries = toChartSeries(byProduct, false).slice(0, 10);
    var sellerSeries = toChartSeries(bySeller, false);
    var brandSeries = toChartSeries(byBrand, false);
    return {
      day: daySeries,
      product: productSeries,
      seller: sellerSeries,
      brand: brandSeries
    };
  }

  function accumulate(map, key, value) {
    if (!map[key]) {
      map[key] = 0;
    }
    map[key] += value;
  }

  function toChartSeries(map, sortByDate) {
    var entries = [];
    for (var key in map) {
      if (map.hasOwnProperty(key)) {
        entries.push([key, map[key]]);
      }
    }
    if (sortByDate) {
      entries.sort(function(a, b) {
        return a[0].localeCompare(b[0], 'fa');
      });
    } else {
      entries.sort(function(a, b) {
        return b[1] - a[1];
      });
    }
    return entries;
  }

  function exportToSheet(dataset, filters) {
    var ss = SpreadsheetApp.getActive();
    var baseName = 'تحلیل فروش';
    var timestamp = Utilities.formatDate(new Date(), 'Asia/Tehran', 'yyyyMMdd-HHmmss');
    var sheetName = baseName + ' ' + timestamp;
    var safeName = sheetName;
    var suffix = 1;
    while (ss.getSheetByName(safeName)) {
      safeName = sheetName + ' (' + suffix + ')';
      suffix++;
    }
    var sheet = ss.insertSheet(safeName);
    var amountKey = filters.amountType === 'priceList' ? 'priceList' : 'pricePaid';
    var table = buildTable(dataset, filters, amountKey);
    var headers = table.headers;
    if (table.rows.length === 0) {
      sheet.getRange(1, 1).setValue('هیچ داده‌ای مطابق فیلترها یافت نشد.');
      sheet.autoResizeColumns(1, 1);
      return safeName;
    }
    var range = sheet.getRange(1, 1, table.rows.length + 1, headers.length);
    var values = [headers].concat(table.rows);
    range.setValues(values);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
    applyNumberFormats(sheet, headers.length, table.valueTypes, table.rows.length + 1);
    return safeName;
  }

  function applyNumberFormats(sheet, columnCount, valueTypes, totalRows) {
    if (totalRows <= 1) {
      return;
    }
    for (var col = 0; col < columnCount; col++) {
      var type = valueTypes[col];
      if (type === 'currency' || type === 'number') {
        var range = sheet.getRange(2, col + 1, totalRows - 1, 1);
        if (type === 'currency') {
          range.setNumberFormat('#,##0" تومان"');
        } else {
          range.setNumberFormat('#,##0');
        }
      }
    }
  }

  function getGroupKey(row, groupBy) {
    switch (groupBy) {
      case 'day':
        return row.orderDateJalaliShort || formatDateForDisplay(row.orderDate);
      case 'week':
        return getWeekKey(row.orderDate);
      case 'month':
        return getMonthKey(row);
      case 'seller':
        return row.seller || NULL_LABEL;
      case 'brand':
        return row.brand || NULL_LABEL;
      case 'location':
        return row.location || NULL_LABEL;
      case 'sku':
        return row.sku || NULL_LABEL;
      case 'productName':
      case 'product':
        return row.productName || row.sku || NULL_LABEL;
      default:
        return row.orderDateJalaliShort || formatDateForDisplay(row.orderDate);
    }
  }

  function getWeekKey(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) {
      return 'هفته نامشخص';
    }
    var tmp = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    var dayNum = tmp.getUTCDay() || 7;
    tmp.setUTCDate(tmp.getUTCDate() + 4 - dayNum);
    var yearStart = new Date(Date.UTC(tmp.getUTCFullYear(), 0, 1));
    var weekNo = Math.ceil((((tmp - yearStart) / 86400000) + 1) / 7);
    return tmp.getUTCFullYear() + ' - هفته ' + weekNo;
  }

  function getMonthKey(row) {
    if (row.orderDateJalaliShort && row.orderDateJalaliShort.indexOf('/') > -1) {
      var parts = row.orderDateJalaliShort.split('/');
      if (parts.length >= 2) {
        return parts[0] + '/' + parts[1];
      }
    }
    if (row.orderDate instanceof Date && !isNaN(row.orderDate.getTime())) {
      var month = row.orderDate.getMonth() + 1;
      return row.orderDate.getFullYear() + '/' + (month < 10 ? '0' + month : month);
    }
    return 'ماه نامشخص';
  }

  function formatDateForDisplay(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) {
      return '';
    }
    return Utilities.formatDate(date, 'Asia/Tehran', 'yyyy/MM/dd');
  }

  function makeLookup(list) {
    var lookup = {};
    for (var i = 0; i < list.length; i++) {
      lookup[list[i]] = true;
    }
    return lookup;
  }

  function parseMoney(value) {
    if (value === null || value === undefined || value === '') {
      return 0;
    }
    var str = normalizeDigits(String(value));
    if (!str) return 0;
    str = str.replace(/[\u066C\u066B,]/g, '');
    str = str.replace(/[^0-9.-]/g, '');
    var num = parseFloat(str);
    return isNaN(num) ? 0 : num;
  }

  function parseJalaliLatin(input) {
    if (!input) {
      return null;
    }
    var normalized = normalizeDigits(String(input));
    var parts = normalized.split(' ');
    var dateParts = parts[0] ? parts[0].split('/') : [];
    if (dateParts.length !== 3) {
      return null;
    }
    var jy = parseInt(dateParts[0], 10);
    var jm = parseInt(dateParts[1], 10);
    var jd = parseInt(dateParts[2], 10);
    if (isNaN(jy) || isNaN(jm) || isNaN(jd)) {
      return null;
    }
    var timeParts = parts[1] ? parts[1].split(':') : ['00', '00', '00'];
    var hh = parseInt(timeParts[0], 10) || 0;
    var mm = parseInt(timeParts[1], 10) || 0;
    var ss = parseInt(timeParts[2], 10) || 0;
    var g = jalaliToGregorian(jy, jm, jd);
    return new Date(g[0], g[1] - 1, g[2], hh, mm, ss);
  }

  function jalaliToGregorian(jy, jm, jd) {
    jy = jy - 979;
    jm = jm - 1;
    jd = jd - 1;
    var jDayNo = 365 * jy + Math.floor(jy / 33) * 8 + Math.floor((jy % 33 + 3) / 4);
    for (var i = 0; i < jm; ++i) {
      jDayNo += (i < 6) ? 31 : 30;
    }
    jDayNo += jd;
    var gDayNo = jDayNo + 79;
    var gy = 1600 + 400 * Math.floor(gDayNo / 146097);
    gDayNo = gDayNo % 146097;
    var leap = true;
    if (gDayNo >= 36525) {
      gDayNo--;
      gy += 100 * Math.floor(gDayNo / 36524);
      gDayNo = gDayNo % 36524;
      if (gDayNo >= 365) {
        gDayNo++;
      } else {
        leap = false;
      }
    }
    gy += 4 * Math.floor(gDayNo / 1461);
    gDayNo %= 1461;
    if (gDayNo >= 366) {
      leap = false;
      gDayNo--;
      gy += Math.floor(gDayNo / 365);
      gDayNo = gDayNo % 365;
    }
    var gd;
    var gm;
    var days = [0, 31, (leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    for (gm = 1; gm <= 12; gm++) {
      var v = days[gm];
      if (gDayNo < v) {
        break;
      }
      gDayNo -= v;
    }
    gd = gDayNo + 1;
    return [gy, gm, gd];
  }

  function toJalaliString(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) {
      return '';
    }
    var gy = date.getFullYear();
    var gm = date.getMonth() + 1;
    var gd = date.getDate();
    var j = gregorianToJalali(gy, gm, gd);
    var jm = j[1] < 10 ? '0' + j[1] : j[1];
    var jd = j[2] < 10 ? '0' + j[2] : j[2];
    return j[0] + '/' + jm + '/' + jd;
  }

  function localeSort(a, b) {
    return String(a).localeCompare(String(b), 'fa');
  }

  function makeProductKey(sku, name) {
    var skuStr = sku != null ? String(sku).trim() : '';
    var normalizedSku = normalizeDigits(skuStr).trim();
    if (normalizedSku) {
      return 'SKU:' + normalizedSku;
    }
    var nameStr = name != null ? String(name).trim().toLowerCase() : '';
    if (nameStr) {
      return 'NAME:' + nameStr;
    }
    return null;
  }

  function getInventorySnapshot() {
    var ss = SpreadsheetApp.getActive();
    if (!ss) {
      return { byKey: {}, totalUnits: 0, distinctSkuCount: 0 };
    }
    var posInventoryRangeName = Config.namedRanges.posInventory;
    var invRange = ss.getRangeByName(posInventoryRangeName);
    if (!invRange) {
      return { byKey: {}, totalUnits: 0, distinctSkuCount: 0 };
    }
    var sheet = invRange.getSheet();
    var startRow = getDataStartRow(invRange);
    var lastRow = getLastDataRow(invRange);
    if (lastRow < startRow) {
      return { byKey: {}, totalUnits: 0, distinctSkuCount: 0 };
    }
    var numRows = lastRow - startRow + 1;
    var numCols = invRange.getNumColumns();
    var values = sheet.getRange(startRow, invRange.getColumn(), numRows, numCols).getValues();
    var byKey = {};
    var totalUnits = 0;
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var status = numCols > 10 ? row[10] : null;
      var isMissing = status === true || String(status).toLowerCase() === 'true';
      if (isMissing) {
        continue;
      }
      var rawSku = row[8] != null ? String(row[8]).trim() : '';
      var name = row[0] != null ? String(row[0]).trim() : '';
      var key = makeProductKey(rawSku, name);
      if (!key) {
        continue;
      }
      if (!byKey[key]) {
        byKey[key] = {
          key: key,
          sku: normalizeDigits(rawSku).trim(),
          displaySku: rawSku || normalizeDigits(rawSku).trim(),
          name: name,
          brand: row[1] != null ? String(row[1]).trim() : '',
          uniqueCode: row[3] != null ? String(row[3]).trim() : '',
          seller: row[5] != null ? String(row[5]).trim() : '',
          location: row[7] != null ? String(row[7]).trim() : '',
          stockCount: 0
        };
      }
      byKey[key].stockCount += 1;
      totalUnits += 1;
    }
    return {
      byKey: byKey,
      totalUnits: totalUnits,
      distinctSkuCount: Object.keys(byKey).length
    };
  }

  function buildProductInsights(dataset, filters, amountKey, inventorySnapshot) {
    inventorySnapshot = inventorySnapshot || { byKey: {}, totalUnits: 0, distinctSkuCount: 0 };
    var included = (dataset && dataset.included) || [];
    var aggregated = {};
    var earliest = null;
    var latest = null;
    var totalSold = 0;
    for (var i = 0; i < included.length; i++) {
      var row = included[i];
      var key = makeProductKey(row.sku, row.productName);
      if (!key) {
        key = makeProductKey(row.uniqueCode, row.productName);
      }
      if (!key) {
        key = 'ROW:' + i;
      }
      if (!aggregated[key]) {
        aggregated[key] = {
          key: key,
          sku: row.sku || '',
          displaySku: row.sku || '',
          name: row.productName || '',
          brand: row.brand || '',
          seller: row.seller || '',
          location: row.location || '',
          salesCount: 0,
          grossAmount: 0,
          netAmount: 0,
          lastSoldAt: null
        };
      }
      var entry = aggregated[key];
      entry.salesCount += 1;
      entry.grossAmount += row.priceList || 0;
      entry.netAmount += row[amountKey] || 0;
      if (!entry.displaySku && row.sku) {
        entry.displaySku = row.sku;
      }
      if (!entry.name && row.productName) {
        entry.name = row.productName;
      }
      if (!entry.brand && row.brand) {
        entry.brand = row.brand;
      }
      if (!entry.seller && row.seller) {
        entry.seller = row.seller;
      }
      if (!entry.location && row.location) {
        entry.location = row.location;
      }
      if (row.orderDate instanceof Date && !isNaN(row.orderDate.getTime())) {
        if (!entry.lastSoldAt || row.orderDate > entry.lastSoldAt) {
          entry.lastSoldAt = row.orderDate;
        }
        if (!earliest || row.orderDate < earliest) {
          earliest = row.orderDate;
        }
        if (!latest || row.orderDate > latest) {
          latest = row.orderDate;
        }
      }
      totalSold += 1;
    }
    var windowDays = determineAnalysisWindowDays(dataset, filters, earliest, latest);
    var records = [];
    var invMap = inventorySnapshot.byKey || {};
    for (var key in aggregated) {
      if (!aggregated.hasOwnProperty(key)) {
        continue;
      }
      var info = aggregated[key];
      var inv = invMap[key] || null;
      var stockCount = inv && inv.stockCount ? inv.stockCount : 0;
      if (inv) {
        if (!info.sku && inv.sku) {
          info.sku = inv.sku;
        }
        if (!info.displaySku && inv.displaySku) {
          info.displaySku = inv.displaySku;
        }
        if (!info.name && inv.name) {
          info.name = inv.name;
        }
        if (!info.brand && inv.brand) {
          info.brand = inv.brand;
        }
        if (!info.seller && inv.seller) {
          info.seller = inv.seller;
        }
        if (!info.location && inv.location) {
          info.location = inv.location;
        }
      }
      var velocity = windowDays > 0 ? info.salesCount / windowDays : 0;
      var velocityRounded = velocity > 0 ? Number(velocity.toFixed(2)) : 0;
      var restockThreshold = velocity > 0 ? Math.max(MIN_LOW_STOCK_THRESHOLD, Math.ceil(velocity * RESTOCK_SAFETY_DAYS)) : MIN_LOW_STOCK_THRESHOLD;
      var lowThreshold = velocity > 0 ? Math.max(MIN_LOW_STOCK_THRESHOLD, Math.ceil(restockThreshold / 2)) : MIN_LOW_STOCK_THRESHOLD;
      var isOutOfStock = stockCount <= 0;
      var isLowStock = !isOutOfStock && stockCount <= lowThreshold;
      var needsRestock = stockCount < restockThreshold && info.salesCount >= MIN_RESTOCK_SALES;
      var coverageRaw = velocity > 0 ? stockCount / velocity : null;
      var coverageDays = coverageRaw != null && isFinite(coverageRaw) ? Number(coverageRaw.toFixed(1)) : null;
      var targetLevel = needsRestock ? Math.max(restockThreshold, Math.ceil(velocity * RESTOCK_TARGET_DAYS)) : restockThreshold;
      var suggested = needsRestock ? Math.max(0, targetLevel - stockCount) : 0;
      records.push({
        key: key,
        name: info.name || '—',
        sku: info.sku || '',
        displaySku: info.displaySku || info.sku || '',
        brand: info.brand || '',
        seller: info.seller || '',
        location: info.location || '',
        salesCount: info.salesCount,
        grossAmount: info.grossAmount,
        netAmount: info.netAmount,
        stockCount: stockCount,
        salesVelocityDaily: velocityRounded,
        restockThreshold: restockThreshold,
        lowStockThreshold: lowThreshold,
        needsRestock: needsRestock,
        isLowStock: isLowStock,
        isOutOfStock: isOutOfStock,
        coverageDays: coverageDays,
        suggestedReorder: suggested,
        restockUrgency: needsRestock ? (coverageRaw != null ? coverageRaw : null) : null,
        lastSoldAt: info.lastSoldAt instanceof Date ? info.lastSoldAt.toISOString() : null
      });
    }
    records.sort(function(a, b) {
      return b.salesCount - a.salesCount || b.netAmount - a.netAmount;
    });
    var topByCount = records.slice(0, 10);
    var outOfStock = records.filter(function(r) { return r.isOutOfStock && r.salesCount > 0; }).slice(0, 10);
    var lowStock = records.filter(function(r) { return r.isLowStock && !r.needsRestock; }).sort(function(a, b) {
      return a.stockCount - b.stockCount || b.salesCount - a.salesCount;
    }).slice(0, 10);
    var restock = records.filter(function(r) { return r.needsRestock; }).sort(function(a, b) {
      var aUrgency = a.restockUrgency != null ? a.restockUrgency : Infinity;
      var bUrgency = b.restockUrgency != null ? b.restockUrgency : Infinity;
      if (aUrgency === bUrgency) {
        return b.salesCount - a.salesCount;
      }
      return aUrgency - bUrgency;
    }).slice(0, 10);
    return {
      metrics: {
        windowDays: windowDays,
        totalSold: totalSold,
        distinctProducts: records.length,
        outOfStockCount: outOfStock.length,
        lowStockCount: lowStock.length,
        restockCount: restock.length,
        inventoryUnits: inventorySnapshot.totalUnits || 0,
        inventorySkuCount: inventorySnapshot.distinctSkuCount || 0
      },
      topByCount: topByCount,
      outOfStock: outOfStock,
      lowStock: lowStock,
      restock: restock
    };
  }

  function determineAnalysisWindowDays(dataset, filters, earliest, latest) {
    var from = dataset && dataset.range ? dataset.range.from : null;
    var to = dataset && dataset.range ? dataset.range.to : null;
    if (from instanceof Date && to instanceof Date) {
      var diff = Math.abs(to.getTime() - from.getTime());
      return Math.max(1, Math.round(diff / 86400000) + 1);
    }
    if (filters && filters.dateFrom && filters.dateTo) {
      var parsedFrom = parseJalaliLatin(filters.dateFrom + ' 00:00:00');
      var parsedTo = parseJalaliLatin(filters.dateTo + ' 23:59:59');
      if (parsedFrom instanceof Date && parsedTo instanceof Date) {
        var filterDiff = Math.abs(parsedTo.getTime() - parsedFrom.getTime());
        return Math.max(1, Math.round(filterDiff / 86400000) + 1);
      }
    }
    if (earliest instanceof Date && latest instanceof Date) {
      var dataDiff = Math.abs(latest.getTime() - earliest.getTime());
      return Math.max(1, Math.round(dataDiff / 86400000) + 1);
    }
    return 30;
  }

  function getSkuCostMap(ss) {
    var sheet = ss.getSheetByName('SKU_Costs');
    if (!sheet) {
      return null;
    }
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return null;
    }
    var range = sheet.getRange(2, 1, lastRow - 1, Math.min(sheet.getLastColumn(), 2));
    var values = range.getValues();
    var map = {};
    for (var i = 0; i < values.length; i++) {
      var sku = normalizeDigits(values[i][0] || '').trim();
      if (!sku) continue;
      var cost = parseMoney(values[i][1]);
      map[sku] = cost;
    }
    return map;
  }

  return namespace;
})(this.PosSheetInsights, this);

function PosSheetInsights_showSidebar() {
  PosSheetInsights.showSidebar();
}

function PosSheetInsights_showDialog() {
  PosSheetInsights.showDialog();
}

function PosSheetInsights_refreshCache() {
  var info = PosSheetInsights.refreshCache();
  SpreadsheetApp.getActive().toast('داده‌ها به‌روزرسانی شدند.');
  return info;
}

function PosSheetInsights_setupTrigger() {
  var info = PosSheetInsights.setupTrigger();
  SpreadsheetApp.getActive().toast('زمان‌بندی هر ۶ ساعت تنظیم شد.');
  return info;
}

function PosSheetInsights_runScheduledRefresh() {
  PosSheetInsights.runScheduledRefresh();
}

function PosSheetInsights_getInitialState() {
  return PosSheetInsights.getInitialState();
}

function PosSheetInsights_getFilterOptions() {
  return PosSheetInsights.getFilterOptions();
}

function PosSheetInsights_computeReport(filters) {
  return PosSheetInsights.computeReport(filters);
}

function PosSheetInsights_getChartData(filters) {
  return PosSheetInsights.getChartData(filters);
}

function PosSheetInsights_exportReport(filters) {
  return PosSheetInsights.exportReport(filters);
}
