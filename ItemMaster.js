/**
 * ============================================================================
 * MODULE: ItemMaster.gs
 * MỤC ĐÍCH: Xử lý quét dữ liệu raw_name, sinh mã SKU/nhãn unmapped và đồng bộ.
 * ============================================================================
 */
/**
 * Đồng bộ tự động ITEM_MASTER từ TRANSACTION và STOCKTAKE.
 * Chỉ thêm mới mã SKU, đồng thời điền 2 trường: Mã SKU và Tên chuẩn từ MAPPING.
 */
// function runSyncItemMasterFromInventoryAndStocktake() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const ui = SpreadsheetApp.getUi();

//   // 1. Gọi API lấy cấu trúc SCHEMA chuẩn hóa
//   const schemaMap = schemaGetMap();
//   if (!schemaMap) {
//     ui.alert("⚠️ Lỗi Schema", "Không thể đọc hoặc cấu trúc sheet SCHEMA không hợp lệ!", ui.ButtonSet.OK);
//     return;
//   }

//   const mappingSheet = ss.getSheetByName("MAPPING");
//   const itemMasterSheet = ss.getSheetByName("ITEM_MASTER");
//   const routeSheet = ss.getSheetByName("ROUTE_MAP");
//   const transSheet = ss.getSheetByName("TRANSACTION");
//   const stocktakeSheet = ss.getSheetByName("STOCKTAKE") || ss.getSheetByName("STOCK_TAKE");

//   if (!mappingSheet || !itemMasterSheet || !routeSheet || !transSheet) {
//     ui.alert("⚠️ Lỗi", "Thiếu một trong các sheet hệ thống cốt lõi (MAPPING, ITEM_MASTER, ROUTE_MAP, TRANSACTION)!", ui.ButtonSet.OK);
//     return;
//   }

//   // Hàm chuyển đổi wildcard (*) thành RegEx động
//   function patternToRegex(pattern) {
//     if (!pattern || pattern === "*") return /^.*$/;
//     const escaped = pattern.replace(/([.+?^=!:${}()|[\]/\\])/g, "\\$1");
//     return new RegExp("^" + escaped.replace(/\*/g, ".*") + "$", "i");
//   }

//   const validItemCodes = new Set();

//   // --- NGUỒN 1: QUÉT TRANSACTION (THEO ROUTE_MAP) ---
//   const rFromCode = schemaGetColIndex(schemaMap, "ROUTE_MAP", "from_code");
//   const rToCode   = schemaGetColIndex(schemaMap, "ROUTE_MAP", "to_code");
//   const rIsInv    = schemaGetColIndex(schemaMap, "ROUTE_MAP", "is_inventory");

//   if (rFromCode !== -1 && rToCode !== -1 && rIsInv !== -1) {
//     const validRouteRegexes = []; 
//     const routeData = routeSheet.getDataRange().getValues();
//     for (let i = 1; i < routeData.length; i++) {
//       const row = routeData[i];
//       const fromPat = row[rFromCode] ? row[rFromCode].toString().trim() : "";
//       const toPat   = row[rToCode] ? row[rToCode].toString().trim() : "";
//       const isInv   = row[rIsInv];

//       const isTrueInv = (isInv === true || isInv.toString().toUpperCase() === "TRUE" || isInv === 1);
//       if (isTrueInv) {
//         validRouteRegexes.push({
//           fromRegex: patternToRegex(fromPat),
//           toRegex: patternToRegex(toPat)
//         });
//       }
//     }

//     const colFrom = schemaGetColIndex(schemaMap, "TRANSACTION", "from_code");
//     const colTo   = schemaGetColIndex(schemaMap, "TRANSACTION", "to_code");
//     const colItem = schemaGetColIndex(schemaMap, "TRANSACTION", "item_code");

//     if (colFrom !== -1 && colTo !== -1 && colItem !== -1) {
//       const transData = transSheet.getDataRange().getValues();
//       for (let i = 1; i < transData.length; i++) {
//         const row = transData[i];
//         const fromCode = row[colFrom] ? row[colFrom].toString().trim() : "";
//         const toCode = row[colTo] ? row[colTo].toString().trim() : "";
//         const itemVal = row[colItem] ? row[colItem].toString().trim() : "";

//         if (!itemVal || itemVal.toLowerCase().includes("unmapped")) continue;

//         let isMatch = false;
//         for (let route of validRouteRegexes) {
//           if (route.fromRegex.test(fromCode) && route.toRegex.test(toCode)) {
//             isMatch = true;
//             break;
//           }
//         }

//         if (isMatch) {
//           validItemCodes.add(itemVal);
//         }
//       }
//     }
//   }

//   // --- NGUỒN 2: QUÉT STOCKTAKE ---
//   if (stocktakeSheet) {
//     const stockSchemaName = schemaMap["STOCKTAKE"] ? "STOCKTAKE" : (schemaMap["STOCK_TAKE"] ? "STOCK_TAKE" : "");
//     if (stockSchemaName) {
//       const colStockItem = schemaGetColIndex(schemaMap, stockSchemaName, "item_code");
//       if (colStockItem !== -1) {
//         const stockData = stocktakeSheet.getDataRange().getValues();
//         for (let i = 1; i < stockData.length; i++) {
//           const itemVal = stockData[i][colStockItem] ? stockData[i][colStockItem].toString().trim() : "";
//           if (itemVal && !itemVal.toLowerCase().includes("unmapped")) {
//             validItemCodes.add(itemVal);
//           }
//         }
//       }
//     }
//   }

//   if (validItemCodes.size === 0) {
//     ui.alert("ℹ️ Thông báo", "Không tìm thấy mã SKU nào từ TRANSACTION hoặc STOCKTAKE.", ui.ButtonSet.OK);
//     return;
//   }

//   // 2. Đọc MAPPING để lấy Tên chuẩn (item_name) tương ứng với từng SKU
//   const colMapItemName = schemaGetColIndex(schemaMap, "MAPPING", "item_name");
//   const colMapItemCode = schemaGetColIndex(schemaMap, "MAPPING", "item_code");

//   if (colMapItemCode === -1) {
//     ui.alert("⚠️ Lỗi Schema MAPPING", "Thiếu cột item_code trong SCHEMA của MAPPING!", ui.ButtonSet.OK);
//     return;
//   }

//   const mappingData = mappingSheet.getDataRange().getValues();
//   const mappingNameByCode = new Map();
//   for (let i = 1; i < mappingData.length; i++) {
//     const row = mappingData[i];
//     const code = row[colMapItemCode] ? row[colMapItemCode].toString().trim() : "";
//     const name = colMapItemName !== -1 && row[colMapItemName] ? row[colMapItemName].toString().trim() : "";

//     if (code) {
//       mappingNameByCode.set(code, name || code);
//     }
//   }

//   // 3. Kiểm tra trùng với ITEM_MASTER hiện tại (cột SKU nằm ở index 0)
//   const masterData = itemMasterSheet.getDataRange().getValues();
//   const existingMasterSkus = new Set();
//   for (let i = 1; i < masterData.length; i++) {
//     const sku = masterData[i][0] ? masterData[i][0].toString().trim() : "";
//     if (sku) existingMasterSkus.add(sku);
//   }

//   // 4. Lọc và chuẩn bị danh sách dòng mới (Chỉ chứa SKU và Tên chuẩn theo đúng yêu cầu)
//   const rowsToAdd = [];
//   const processedSkus = new Set();

//   validItemCodes.forEach(skuCode => {
//     if (!existingMasterSkus.has(skuCode) && !processedSkus.has(skuCode)) {
//       processedSkus.add(skuCode);
      
//       const finalItemName = mappingNameByCode.get(skuCode) || skuCode;

//       // Chỉ đẩy 2 trường quan trọng: [Mã SKU, Tên hàng chuẩn hóa]
//       rowsToAdd.push([
//         skuCode,        // Cột 1: Mã SKU
//         finalItemName   // Cột 2: Tên chuẩn từ MAPPING
//       ]);
//     }
//   });

//   if (rowsToAdd.length === 0) {
//     ui.alert("ℹ️ Thông báo", "Tất cả các mã SKU quét được đều đã có sẵn trong ITEM_MASTER.", ui.ButtonSet.OK);
//     return;
//   }

//   // 5. Ghi dữ liệu xuống ITEM_MASTER (chỉ ghi vào cột 1 và cột 2)
//   const lastRow = itemMasterSheet.getLastRow();
//   itemMasterSheet.getRange(lastRow + 1, 1, rowsToAdd.length, 2).setValues(rowsToAdd);

//   ui.alert("✅ Thành công", `Đã đồng bộ thành công ${rowsToAdd.length} mã SKU mới (với Mã SKU & Tên chuẩn) vào ITEM_MASTER!`, ui.ButtonSet.OK);
// }
function runSyncItemMasterFromInventoryAndStocktake() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. Gọi API lấy cấu trúc SCHEMA chuẩn hóa
  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi Schema", "Không thể đọc hoặc cấu trúc sheet SCHEMA không hợp lệ!", ui.ButtonSet.OK);
    return;
  }

  // Lấy sheet_name vật lý động từ SCHEMA
  const mappingSheetName    = schemaGetSheetName(schemaMap, "MAPPING");
  const itemMasterSheetName = schemaGetSheetName(schemaMap, "ITEM_MASTER");
  const routeSheetName      = schemaGetSheetName(schemaMap, "ROUTE_MAP");
  const transSheetName      = schemaGetSheetName(schemaMap, "TRANSACTION");
  const stocktakeSheetName  = schemaGetSheetName(schemaMap, "STOCKTAKE") || schemaGetSheetName(schemaMap, "STOCK_TAKE");

  const mappingSheet    = mappingSheetName ? ss.getSheetByName(mappingSheetName) : null;
  const itemMasterSheet = itemMasterSheetName ? ss.getSheetByName(itemMasterSheetName) : null;
  const routeSheet      = routeSheetName ? ss.getSheetByName(routeSheetName) : null;
  const transSheet      = transSheetName ? ss.getSheetByName(transSheetName) : null;
  const stocktakeSheet  = stocktakeSheetName ? ss.getSheetByName(stocktakeSheetName) : null;

  if (!mappingSheet || !itemMasterSheet || !routeSheet || !transSheet) {
    ui.alert("⚠️ Lỗi", "Thiếu một trong các sheet hệ thống cốt lõi (MAPPING, ITEM_MASTER, ROUTE_MAP, TRANSACTION)!", ui.ButtonSet.OK);
    return;
  }

  // Hàm chuyển đổi wildcard (*) thành RegEx động
  function patternToRegex(pattern) {
    if (!pattern || pattern === "*") return /^.*$/;
    const escaped = pattern.replace(/([.+?^=!:${}()|[\]/\\])/g, "\\$1");
    return new RegExp("^" + escaped.replace(/\*/g, ".*") + "$", "i");
  }

  const validItemCodes = new Set();

  // --- NGUỒN 1: QUÉT TRANSACTION (THEO ROUTE_MAP) ---
  const rFromCode = schemaGetColIndex(schemaMap, "ROUTE_MAP", "from_code");
  const rToCode   = schemaGetColIndex(schemaMap, "ROUTE_MAP", "to_code");
  const rIsInv    = schemaGetColIndex(schemaMap, "ROUTE_MAP", "is_inventory");

  if (rFromCode !== -1 && rToCode !== -1 && rIsInv !== -1) {
    const validRouteRegexes = []; 
    const routeData = routeSheet.getDataRange().getValues();
    for (let i = 1; i < routeData.length; i++) {
      const row = routeData[i];
      const fromPat = row[rFromCode] ? row[rFromCode].toString().trim() : "";
      const toPat   = row[rToCode] ? row[rToCode].toString().trim() : "";
      const isInv   = row[rIsInv];

      const isTrueInv = (isInv === true || isInv.toString().toUpperCase() === "TRUE" || isInv === 1);
      if (isTrueInv) {
        validRouteRegexes.push({
          fromRegex: patternToRegex(fromPat),
          toRegex: patternToRegex(toPat)
        });
      }
    }

    const colFrom = schemaGetColIndex(schemaMap, "TRANSACTION", "from_code");
    const colTo   = schemaGetColIndex(schemaMap, "TRANSACTION", "to_code");
    const colItem = schemaGetColIndex(schemaMap, "TRANSACTION", "item_code");

    if (colFrom !== -1 && colTo !== -1 && colItem !== -1) {
      const transData = transSheet.getDataRange().getValues();
      for (let i = 1; i < transData.length; i++) {
        const row = transData[i];
        const fromCode = row[colFrom] ? row[colFrom].toString().trim() : "";
        const toCode = row[colTo] ? row[colTo].toString().trim() : "";
        const itemVal = row[colItem] ? row[colItem].toString().trim() : "";

        if (!itemVal || itemVal.toLowerCase().includes("unmapped")) continue;

        let isMatch = false;
        for (let route of validRouteRegexes) {
          if (route.fromRegex.test(fromCode) && route.toRegex.test(toCode)) {
            isMatch = true;
            break;
          }
        }

        if (isMatch) {
          validItemCodes.add(itemVal);
        }
      }
    }
  }

  // --- NGUỒN 2: QUÉT STOCKTAKE ---
  if (stocktakeSheet) {
    const stockSchemaName = schemaMap["STOCKTAKE"] ? "STOCKTAKE" : (schemaMap["STOCK_TAKE"] ? "STOCK_TAKE" : "");
    if (stockSchemaName) {
      const colStockItem = schemaGetColIndex(schemaMap, stockSchemaName, "item_code");
      if (colStockItem !== -1) {
        const stockData = stocktakeSheet.getDataRange().getValues();
        for (let i = 1; i < stockData.length; i++) {
          const itemVal = stockData[i][colStockItem] ? stockData[i][colStockItem].toString().trim() : "";
          if (itemVal && !itemVal.toLowerCase().includes("unmapped")) {
            validItemCodes.add(itemVal);
          }
        }
      }
    }
  }

  if (validItemCodes.size === 0) {
    ui.alert("ℹ️ Thông báo", "Không tìm thấy mã SKU nào từ TRANSACTION hoặc STOCKTAKE.", ui.ButtonSet.OK);
    return;
  }

  // 2. Đọc MAPPING để lấy Tên chuẩn (item_name) tương ứng với từng SKU
  const colMapItemName = schemaGetColIndex(schemaMap, "MAPPING", "item_name");
  const colMapItemCode = schemaGetColIndex(schemaMap, "MAPPING", "item_code");

  if (colMapItemCode === -1) {
    ui.alert("⚠️ Lỗi Schema MAPPING", "Thiếu cột item_code trong SCHEMA của MAPPING!", ui.ButtonSet.OK);
    return;
  }

  const mappingData = mappingSheet.getDataRange().getValues();
  const mappingNameByCode = new Map();
  for (let i = 1; i < mappingData.length; i++) {
    const row = mappingData[i];
    const code = row[colMapItemCode] ? row[colMapItemCode].toString().trim() : "";
    const name = colMapItemName !== -1 && row[colMapItemName] ? row[colMapItemName].toString().trim() : "";

    if (code) {
      mappingNameByCode.set(code, name || code);
    }
  }

  // 3. Tra cứu vị trí cột động trong ITEM_MASTER
  const idxMasterCode = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_code");
  const idxMasterName = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_name");

  if (idxMasterCode === -1 || idxMasterName === -1) {
    ui.alert("⚠️ Lỗi Schema ITEM_MASTER", "Mã cột 'item_code' hoặc 'item_name' không tồn tại trong SCHEMA của ITEM_MASTER!", ui.ButtonSet.OK);
    return;
  }

  const masterData = itemMasterSheet.getDataRange().getValues();
  const maxMasterColCount = schemaMap["ITEM_MASTER"].columns.length;

  const existingMasterSkus = new Set();
  for (let i = 1; i < masterData.length; i++) {
    const sku = masterData[i][idxMasterCode] ? masterData[i][idxMasterCode].toString().trim() : "";
    if (sku) existingMasterSkus.add(sku);
  }

  // 4. Lọc và chuẩn bị danh sách dòng mới
  const rowsToAdd = [];
  const processedSkus = new Set();

  validItemCodes.forEach(skuCode => {
    if (!existingMasterSkus.has(skuCode) && !processedSkus.has(skuCode)) {
      processedSkus.add(skuCode);
      
      const finalItemName = mappingNameByCode.get(skuCode) || skuCode;

      // Tạo dòng mới có số cột tương ứng cấu trúc schema ITEM_MASTER
      const newRow = new Array(maxMasterColCount).fill("");
      newRow[idxMasterCode] = skuCode;
      newRow[idxMasterName] = finalItemName;

      rowsToAdd.push(newRow);
    }
  });

  if (rowsToAdd.length === 0) {
    ui.alert("ℹ️ Thông báo", "Tất cả các mã SKU quét được đều đã có sẵn trong ITEM_MASTER.", ui.ButtonSet.OK);
    return;
  }

  // 5. Ghi dữ liệu xuống ITEM_MASTER đúng vị trí các cột
  const lastRow = itemMasterSheet.getLastRow();
  itemMasterSheet.getRange(lastRow + 1, 1, rowsToAdd.length, maxMasterColCount).setValues(rowsToAdd);

  ui.alert("✅ Thành công", `Đã đồng bộ thành công ${rowsToAdd.length} mã SKU mới (với Mã SKU & Tên chuẩn) vào ITEM_MASTER!`, ui.ButtonSet.OK);
}

// /**
//  * Hàm Global Core nghiệp vụ đồng bộ từ MENU sang ITEM_MASTER.
//  * Chỉ lọc bỏ các mã có item_type = "Dịch vụ".
//  * Quy tắc: [moduleName][Action][Entity] -> itemMasterSyncFromMenu
//  */
// function itemMasterSyncFromMenu() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const ui = SpreadsheetApp.getUi();

//   // 1. Lấy cấu trúc SCHEMA chuẩn hóa qua API SchemaCore
//   const schemaMap = schemaGetMap();
//   if (!schemaMap) {
//     ui.alert("⚠️ Lỗi Schema", "Không thể đọc hoặc cấu trúc sheet SCHEMA không hợp lệ!", ui.ButtonSet.OK);
//     return;
//   }

//   // Lấy sheet_name vật lý động từ SCHEMA
//   const menuSheetName       = schemaGetSheetName(schemaMap, "MENU");
//   const itemMasterSheetName = schemaGetSheetName(schemaMap, "ITEM_MASTER");
//   const mappingSheetName    = schemaGetSheetName(schemaMap, "MAPPING");

//   const menuSheet       = menuSheetName ? ss.getSheetByName(menuSheetName) : null;
//   const itemMasterSheet = itemMasterSheetName ? ss.getSheetByName(itemMasterSheetName) : null;
//   const mappingSheet    = mappingSheetName ? ss.getSheetByName(mappingSheetName) : null;

//   if (!menuSheet || !itemMasterSheet || !mappingSheet) {
//     ui.alert("⚠️ Lỗi", "Thiếu một trong các sheet hệ thống cốt lõi (MENU, ITEM_MASTER, MAPPING)!", ui.ButtonSet.OK);
//     return;
//   }

//   // 2. Tra cứu vị trí các cột trong sheet MENU thông qua SCHEMA
//   const colMenuItemCode = schemaGetColIndex(schemaMap, "MENU", "item_code");
//   const colMenuItemType = schemaGetColIndex(schemaMap, "MENU", "item_type");
//   const colMenuStatus   = schemaGetColIndex(schemaMap, "MENU", "status");

//   if (colMenuItemCode === -1 || colMenuItemType === -1) {
//     ui.alert("⚠️ Lỗi Schema MENU", "Thiếu cấu hình cột item_code hoặc item_type trong SCHEMA của MENU!", ui.ButtonSet.OK);
//     return;
//   }

//   // 3. Đọc dữ liệu từ MENU và lọc loại bỏ item_type = "Dịch vụ"
//   const menuData = menuSheet.getDataRange().getValues();
//   const validSkusFromMenu = new Set();

//   for (let i = 1; i < menuData.length; i++) {
//     const row = menuData[i];
//     const skuCode = cleanCodeValue_(row[colMenuItemCode]);
//     const itemType = cleanCodeValue_(row[colMenuItemType]);
//     const status = colMenuStatus !== -1 ? cleanCodeValue_(row[colMenuStatus]) : "ACTIVE";

//     // Bỏ qua mã rỗng, chứa "unmapped" hoặc trạng thái INACTIVE (nếu có)
//     if (!skuCode || skuCode.toLowerCase().includes("unmapped")) continue;

//     // TIÊU CHÍ LỌC LOẠI TRỪ: Bỏ qua các mã có item_type = "Dịch vụ"
//     if (itemType.toLowerCase() === "dịch vụ" || itemType.toLowerCase() === "dich vu") continue;

//     validSkusFromMenu.add(skuCode);
//   }

//   if (validSkusFromMenu.size === 0) {
//     ui.alert("ℹ️ Thông báo", "Không tìm thấy mã SKU hợp lệ nào từ MENU (đã lọc bỏ các mã Dịch vụ).", ui.ButtonSet.OK);
//     return;
//   }

//   // 4. Đọc MAPPING để lấy Tên chuẩn (item_name) tương ứng
//   const colMapItemCode = schemaGetColIndex(schemaMap, "MAPPING", "item_code");
//   const colMapItemName = schemaGetColIndex(schemaMap, "MAPPING", "item_name");

//   if (colMapItemCode === -1) {
//     ui.alert("⚠️ Lỗi Schema MAPPING", "Thiếu cột item_code trong SCHEMA của MAPPING!", ui.ButtonSet.OK);
//     return;
//   }

//   const mappingData = mappingSheet.getDataRange().getValues();
//   const mappingNameByCode = new Map();
//   for (let i = 1; i < mappingData.length; i++) {
//     const row = mappingData[i];
//     const code = cleanCodeValue_(row[colMapItemCode]);
//     const name = colMapItemName !== -1 ? cleanCodeValue_(row[colMapItemName]) : "";
//     if (code) {
//       mappingNameByCode.set(code, name || code);
//     }
//   }

//   // 5. Đọc vị trí cột trong ITEM_MASTER qua SCHEMA và lấy danh sách SKU hiện tại
//   const idxMasterSku  = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_code");
//   const idxMasterName = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_name");

//   if (idxMasterSku === -1 || idxMasterName === -1) {
//     ui.alert("⚠️ Lỗi Schema ITEM_MASTER", "Thiếu cấu hình cột item_code hoặc item_name trong SCHEMA của ITEM_MASTER!", ui.ButtonSet.OK);
//     return;
//   }

//   const masterData = itemMasterSheet.getDataRange().getValues();
//   const maxMasterColCount = schemaMap["ITEM_MASTER"].columns.length;

//   const existingMasterSkus = new Set();
//   for (let i = 1; i < masterData.length; i++) {
//     const sku = cleanCodeValue_(masterData[i][idxMasterSku]);
//     if (sku) existingMasterSkus.add(sku);
//   }

//   // 6. Lọc và chuẩn bị danh sách dòng mới thêm vào ITEM_MASTER
//   const rowsToAdd = [];
//   const processedSkus = new Set();

//   validSkusFromMenu.forEach(skuCode => {
//     if (!existingMasterSkus.has(skuCode) && !processedSkus.has(skuCode)) {
//       processedSkus.add(skuCode);
      
//       const finalItemName = mappingNameByCode.get(skuCode) || skuCode;

//       // Tạo dòng mới chuẩn hóa theo tổng số cột của ITEM_MASTER trong SCHEMA
//       const newRow = new Array(maxMasterColCount).fill("");
//       newRow[idxMasterSku]  = skuCode;
//       newRow[idxMasterName] = finalItemName;

//       rowsToAdd.push(newRow);
//     }
//   });

//   if (rowsToAdd.length === 0) {
//     ui.alert("ℹ️ Thông báo", "Tất cả các mã SKU hợp lệ từ MENU đều đã tồn tại trong ITEM_MASTER.", ui.ButtonSet.OK);
//     return;
//   }

//   // 7. Ghi dữ liệu xuống ITEM_MASTER đúng số lượng cột theo SCHEMA
//   const lastRow = itemMasterSheet.getLastRow();
//   itemMasterSheet.getRange(lastRow + 1, 1, rowsToAdd.length, maxMasterColCount).setValues(rowsToAdd);

//   ui.alert("✅ Thành công", `Đã đồng bộ thành công ${rowsToAdd.length} mã SKU mới từ MENU sang ITEM_MASTER (đã loại bỏ mã Dịch vụ)!`, ui.ButtonSet.OK);
// }

/**
 * Entry Point gọi từ Menu UI để đồng bộ danh mục món từ MENU sang ITEM_MASTER.
 * Quy tắc: [actionEntityDescription] -> runSyncItemMasterFromMenu
 */
function runSyncItemMasterFromMenu() {
  itemMasterSyncFromMenu();
}

/**
 * Hàm Global Core nghiệp vụ đồng bộ MENU -> ITEM_MASTER.
 * - Khóa PK đồng bộ chính là menu_code.
 * - Lọc bỏ item_type = "Dịch vụ".
 * - Tham chiếu 100% qua SCHEMA.
 * Quy tắc: [moduleName][Action][Entity] -> itemMasterSyncFromMenu
 */
function itemMasterSyncFromMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. Lấy cấu trúc SCHEMA chuẩn hóa từ SchemaCore
  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi Schema", "Không thể đọc hoặc cấu trúc sheet SCHEMA không hợp lệ!", ui.ButtonSet.OK);
    return;
  }

  // Lấy sheet_name vật lý động từ SCHEMA
  const menuSheetName       = schemaGetSheetName(schemaMap, "MENU");
  const itemMasterSheetName = schemaGetSheetName(schemaMap, "ITEM_MASTER");

  const menuSheet       = menuSheetName ? ss.getSheetByName(menuSheetName) : null;
  const itemMasterSheet = itemMasterSheetName ? ss.getSheetByName(itemMasterSheetName) : null;

  if (!menuSheet || !itemMasterSheet) {
    ui.alert("⚠️ Lỗi", "Thiếu sheet MENU hoặc ITEM_MASTER!", ui.ButtonSet.OK);
    return;
  }

  // 2. Tra cứu cột trong MENU via SCHEMA
  const colMenuCode     = schemaGetColIndex(schemaMap, "MENU", "menu_code");
  const colMenuName     = schemaGetColIndex(schemaMap, "MENU", "menu_name");
  const colMenuItemType = schemaGetColIndex(schemaMap, "MENU", "item_type");

  if (colMenuCode === -1 || colMenuName === -1 || colMenuItemType === -1) {
    ui.alert("⚠️ Lỗi Schema MENU", "Thiếu cột menu_code, menu_name hoặc item_type trong SCHEMA của MENU!", ui.ButtonSet.OK);
    return;
  }

  // 3. Đọc dữ liệu MENU và lọc mã hợp lệ (Dùng menu_code làm Mã SKU chính trong ITEM_MASTER)
  const menuData = menuSheet.getDataRange().getValues();
  const validMenuItems = new Map(); // Key: menu_code, Value: menu_name

  for (let i = 1; i < menuData.length; i++) {
    const row = menuData[i];
    const menuCode = cleanCodeValue_(row[colMenuCode]);
    const menuName = cleanCodeValue_(row[colMenuName]);
    const itemType = cleanCodeValue_(row[colMenuItemType]);

    // Lọc bỏ mã rỗng, unmapped
    if (!menuCode || menuCode.toLowerCase().includes("unmapped")) continue;

    // LOẠI TRỪ: Bỏ qua item_type = "Dịch vụ"
    if (itemType.toLowerCase() === "dịch vụ" || itemType.toLowerCase() === "dich vu") continue;

    validMenuItems.set(menuCode, menuName || menuCode);
  }

  if (validMenuItems.size === 0) {
    ui.alert("ℹ️ Thông báo", "Không tìm thấy Mã món hợp lệ nào từ MENU để đồng bộ.", ui.ButtonSet.OK);
    return;
  }

  // 4. Tra cứu vị trí cột động trong ITEM_MASTER via SCHEMA
  const idxMasterSku  = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_code");
  const idxMasterName = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_name");

  if (idxMasterSku === -1 || idxMasterName === -1) {
    ui.alert("⚠️ Lỗi Schema ITEM_MASTER", "Thiếu cột item_code hoặc item_name trong SCHEMA của ITEM_MASTER!", ui.ButtonSet.OK);
    return;
  }

  const masterData = itemMasterSheet.getDataRange().getValues();
  const maxMasterColCount = schemaMap["ITEM_MASTER"].columns.length;

  const existingMasterSkus = new Set();
  for (let i = 1; i < masterData.length; i++) {
    const sku = cleanCodeValue_(masterData[i][idxMasterSku]);
    if (sku) existingMasterSkus.add(sku);
  }

  // 5. Lọc và chuẩn bị mảng dữ liệu mới theo đúng kích thước schema ITEM_MASTER
  const rowsToAdd = [];
  validMenuItems.forEach((menuName, menuCode) => {
    if (!existingMasterSkus.has(menuCode)) {
      const newRow = new Array(maxMasterColCount).fill("");
      newRow[idxMasterSku]  = menuCode; // Gán Mã SKU
      newRow[idxMasterName] = menuName; // Gán Tên mặt hàng
      
      rowsToAdd.push(newRow);
    }
  });

  if (rowsToAdd.length === 0) {
    ui.alert("ℹ️ Thông báo", "Tất cả các Mã món từ MENU đều đã có sẵn trong ITEM_MASTER.", ui.ButtonSet.OK);
    return;
  }

  // 6. Ghi dữ liệu xuống ITEM_MASTER đúng vị trí các cột theo SCHEMA
  const lastRow = itemMasterSheet.getLastRow();
  itemMasterSheet.getRange(lastRow + 1, 1, rowsToAdd.length, maxMasterColCount).setValues(rowsToAdd);

  ui.alert("✅ Thành công", `Đã đồng bộ thành công ${rowsToAdd.length} Mã món từ MENU sang ITEM_MASTER!`, ui.ButtonSet.OK);
}

/**
 * Private Helper làm sạch dữ liệu
 */
function cleanCodeValue_(val) {
  if (val === null || val === undefined) return "";
  return val.toString().trim();
}
