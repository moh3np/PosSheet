var APP_VERSION = '1.0.0';

var Config = (function(global) {
  if (global.Config && global.Config.__version === 2) {
    return global.Config;
  }

  var config = {
    __version: 2,
    sheets: {
      tolyandSheet: '1LIR_q1xrpdzcqoBJmNXTO0UJ9dksoBjS7h3Me4PRB1s',
      buyruzSheet: '12-Khe_IZ9S7z_VN_LZQCHdcKEIgKDquviar8cSR_wG8'
    },
    namedRanges: {
      posInventory: 'PosInventory',
      posOrders: 'PosOrders',
      toylandOrders: 'ToylandOrders',
      toylandInventory: 'ToylandInventory',
      buyruzOrders: 'BuyruzPosOrders',
      buyruzInventory: 'BuyruzInventory',
      toylandProducts: 'Products',
      buyruzProductName: 'ProductName',
      buyruzProductSku: 'ProductId',
      buyruzProductsPrice: 'ProductsPrice',
      buyruzProductPriceUpdateTime: 'ProductPriceUpdateTime',
      buyruzProductsStock: 'ProductsStock',
      buyruzFactorSku: 'FactorSKU',
      buyruzFactorPurchaseDate: 'FactorPurchaseDate',
      buyruzFactorNetUnitPrice: 'FactorNetUnitPrice',
      buyruzInventoryProductName: 'InventoryProductName',
      buyruzInventoryProductBrand: 'InventoryProductBrand',
      buyruzInventorySku: 'InventorySKU',
      buyruzInventoryUniqueCode: 'InventoryUniqueCode',
      buyruzInventorySN: 'InventorySN',
      buyruzInventorySupplier: 'InventorySupplier',
      buyruzInventoryPrice: 'InventoryPrice',
      buyruzInventoryLocation: 'InventoryLocation',
      buyruzInventoryLabelPrinted: 'InventoryLablePrinted'
    },
    inventory: {
      skipToylandColumn: 6
    }
  };

  config = Object.freeze(config);
  global.Config = config;
  return config;
})(typeof globalThis === 'object' ? globalThis : this);
