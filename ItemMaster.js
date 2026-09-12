/**
 * ============================================================================
 * MODULE: ItemMaster.gs
 * MỤC ĐÍCH: Xử lý logic nghiệp vụ danh mục vật tư/hàng hóa ITEM_MASTER.
 * Quy tắc đặt tên: itemMasterActionEntity (CamelCase)
 * ============================================================================
 */

/**
 * Core: Đồng bộ tự động ITEM_MASTER từ các sheet Raw Data thuộc nhóm INVENTORY.
 * - Đọc bảng RAW_SOURCE_GROUP để xác định các Schema thuộc nhóm INVENTORY.
 * - Đọc tbl_route_map để LỌC BỎ các giao dịch thuộc tuyến Không tính tồn kho (Stock = FALSE).
 * - Quét tất cả SKU hợp lệ từ các tuyến Stock = TRUE và tự động gán source_group = 'INVENTORY'.
 */
function itemMasterSyncFromInventoryAndStocktake() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi", "Không thể đọc SCHEMA!", ui.ButtonSet.OK);
    return;
  }

  const mappingSheetName    = schemaGetSheetName(schemaMap, "MAPPING");
  const itemMasterSheetName = schemaGetSheetName(schemaMap, "ITEM_MASTER");

  const mappingSheet    = mappingSheetName ? ss.getSheetByName(mappingSheetName) : null;
  const itemMasterSheet = itemMasterSheetName ? ss.getSheetByName(itemMasterSheetName) : null;

  if (!mappingSheet || !itemMasterSheet) {
    ui.alert("⚠️ Lỗi", "Thiếu sheet MAPPING hoặc ITEM_MASTER!", ui.ButtonSet.OK);
    return;
  }

  // -------------------------------------------------------------
  // 1. PRE-LOAD: LẠM DỤNG BỘ NHỚ ĐỆM CHO ROUTE_MAP & LOCATION_MAP
  // -------------------------------------------------------------
  const validStockRoutes = new Set();
  const routeInventoryPairMap = new Map(); // Luồng Pair (fromType -> toType) -> boolean

  const routeSheetName = schemaGetSheetName(schemaMap, "ROUTE_MAP");
  const routeSheet = routeSheetName ? ss.getSheetByName(routeSheetName) : null;

  if (routeSheet) {
    const routeData = routeSheet.getDataRange().getValues();
    const colRouteCode   = schemaGetColIndex(schemaMap, "ROUTE_MAP", "route_code");
    const colFromType    = schemaGetColIndex(schemaMap, "ROUTE_MAP", "from_code");
    const colToType      = schemaGetColIndex(schemaMap, "ROUTE_MAP", "to_code");
    const colIsInventory = schemaGetColIndex(schemaMap, "ROUTE_MAP", "is_inventory");

    if (colIsInventory !== -1 && routeData.length > 1) {
      for (let i = 1; i < routeData.length; i++) {
        const isStock = routeData[i][colIsInventory] === true || String(routeData[i][colIsInventory]).toUpperCase() === "TRUE";
        
        // Cache theo route_code
        if (colRouteCode !== -1) {
          const route = cleanCodeValue_(routeData[i][colRouteCode]);
          if (route && isStock) validStockRoutes.add(route);
        }

        // Cache theo cặp (from_code -> to_code)
        if (colFromType !== -1 && colToType !== -1) {
          const fT = cleanCodeValue_(routeData[i][colFromType]).toUpperCase();
          const tT = cleanCodeValue_(routeData[i][colToType]).toUpperCase();
          if (fT && tT) {
            routeInventoryPairMap.set(`${fT}->${tT}`, isStock);
          }
        }
      }
    }
  }

  // Cache LOCATION_MAP (Code -> Type)
  const locationTypeMap = new Map();
  const locSheetName = schemaGetSheetName(schemaMap, "LOCATION_MAP");
  const locSheet = locSheetName ? ss.getSheetByName(locSheetName) : null;

  if (locSheet) {
    const colLocCode = schemaGetColIndex(schemaMap, "LOCATION_MAP", "location_code");
    const colLocType = schemaGetColIndex(schemaMap, "LOCATION_MAP", "location_type");

    if (colLocCode !== -1 && colLocType !== -1) {
      const locData = locSheet.getDataRange().getValues();
      for (let i = 1; i < locData.length; i++) {
        const code = cleanCodeValue_(locData[i][colLocCode]).toUpperCase();
        const type = cleanCodeValue_(locData[i][colLocType]).toUpperCase();
        if (code) locationTypeMap.set(code, type);
      }
    }
  }

  // Hàm tra cứu siêu tốc O(1) ngay trên bộ nhớ RAM
  const fastCheckIsStockRoute = (fromCode, toCode) => {
    if (!fromCode || !toCode) return false;
    const fUpper = cleanCodeValue_(fromCode).toUpperCase();
    const tUpper = cleanCodeValue_(toCode).toUpperCase();

    const fromType = locationTypeMap.get(fUpper) || fUpper;
    const toType   = locationTypeMap.get(tUpper) || tUpper;

    if (fromType === "EXPENSE" || toType === "EXPENSE") return false;
    if (fromType === "VIRTUAL" && toType === "VIRTUAL") return false;

    const pairKey = `${fromType}->${toType}`;
    return routeInventoryPairMap.has(pairKey) ? routeInventoryPairMap.get(pairKey) : false;
  };

  // -------------------------------------------------------------
  // 2. LẤY SCHEMA THUỘC NHÓM INVENTORY
  // -------------------------------------------------------------
  const sourceGroupMap = getSourceGroupMap_(schemaMap, ss);
  const inventorySchemas = new Set();

  Object.keys(schemaMap).forEach(schemaName => {
    const sGroup = sourceGroupMap.get(schemaName.toLowerCase());
    if (sGroup === "inventory") inventorySchemas.add(schemaName);
  });

  if (inventorySchemas.size === 0) {
    inventorySchemas.add("TRANSACTION");
    inventorySchemas.add("STOCKTAKE");
  }

  // -------------------------------------------------------------
  // 3. QUÉT SKU TRONG BỘ NHỚ
  // -------------------------------------------------------------
  const validItemCodes = new Set();

  inventorySchemas.forEach(schemaName => {
    const colItem  = schemaGetColIndex(schemaMap, schemaName, "item_code");
    const colRoute = schemaGetColIndex(schemaMap, schemaName, "route_code");
    const colFrom  = schemaGetColIndex(schemaMap, schemaName, "from_code");
    const colTo    = schemaGetColIndex(schemaMap, schemaName, "to_code");

    const targetSheetName = schemaGetSheetName(schemaMap, schemaName);
    const targetSheet = targetSheetName ? ss.getSheetByName(targetSheetName) : null;

    if (colItem !== -1 && targetSheet) {
      const data = targetSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const itemVal = cleanCodeValue_(data[i][colItem]);
        
        if (itemVal && !itemVal.toLowerCase().includes("unmapped")) {
          let isStockCalculated = false;

          const routeCode = colRoute !== -1 ? cleanCodeValue_(data[i][colRoute]) : "";
          const fCode     = colFrom !== -1 ? cleanCodeValue_(data[i][colFrom]) : "";
          const tCode     = colTo !== -1 ? cleanCodeValue_(data[i][colTo]) : "";

          if (routeCode) {
            isStockCalculated = validStockRoutes.has(routeCode);
          } else if (fCode && tCode) {
            isStockCalculated = fastCheckIsStockRoute(fCode, tCode);
          } else {
            isStockCalculated = true; // Stocktake
          }

          if (isStockCalculated) {
            validItemCodes.add(itemVal);
          }
        }
      }
    }
  });

  // -------------------------------------------------------------
  // 4. LẤY MAPPING VÀ CẬP NHẬT ITEM_MASTER
  // -------------------------------------------------------------
  const colMapItemName = schemaGetColIndex(schemaMap, "MAPPING", "item_name");
  const colMapItemCode = schemaGetColIndex(schemaMap, "MAPPING", "item_code");

  const mappingNameByCode = new Map();
  if (colMapItemCode !== -1) {
    const mappingData = mappingSheet.getDataRange().getValues();
    for (let i = 1; i < mappingData.length; i++) {
      const code = cleanCodeValue_(mappingData[i][colMapItemCode]);
      const name = colMapItemName !== -1 ? cleanCodeValue_(mappingData[i][colMapItemName]) : "";
      if (code) mappingNameByCode.set(code, name || code);
    }
  }

  const idxMasterCode   = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_code");
  const idxMasterName   = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_name");
  const idxMasterSource = schemaGetColIndex(schemaMap, "ITEM_MASTER", "source_group");

  if (idxMasterCode === -1 || idxMasterName === -1) {
    ui.alert("⚠️ Lỗi Schema ITEM_MASTER", "Thiếu cột 'item_code' hoặc 'item_name'!", ui.ButtonSet.OK);
    return;
  }

  const masterData = itemMasterSheet.getDataRange().getValues();
  const maxMasterColCount = schemaMap["ITEM_MASTER"].columns.length;

  const existingMasterSkus = new Set();
  let updatedExistingCount = 0;

  for (let i = 1; i < masterData.length; i++) {
    const sku = cleanCodeValue_(masterData[i][idxMasterCode]);
    if (sku) {
      existingMasterSkus.add(sku);
      if (idxMasterSource !== -1 && !masterData[i][idxMasterSource]) {
        if (validItemCodes.has(sku)) {
          masterData[i][idxMasterSource] = "INVENTORY";
          updatedExistingCount++;
        }
      }
    }
  }

  const rowsToAdd = [];
  const processedSkus = new Set();

  validItemCodes.forEach(skuCode => {
    if (!existingMasterSkus.has(skuCode) && !processedSkus.has(skuCode)) {
      processedSkus.add(skuCode);
      
      const finalItemName = mappingNameByCode.get(skuCode) || skuCode;
      const newRow = new Array(maxMasterColCount).fill("");
      newRow[idxMasterCode] = skuCode;
      newRow[idxMasterName] = finalItemName;
      
      if (idxMasterSource !== -1) newRow[idxMasterSource] = "INVENTORY";
      rowsToAdd.push(newRow);
    }
  });

  if (updatedExistingCount > 0) {
    itemMasterSheet.getRange(1, 1, masterData.length, maxMasterColCount).setValues(masterData);
  }

  if (rowsToAdd.length > 0) {
    const lastRow = itemMasterSheet.getLastRow();
    itemMasterSheet.getRange(lastRow + 1, 1, rowsToAdd.length, maxMasterColCount).setValues(rowsToAdd);
  }

  let msg = [];
  if (rowsToAdd.length > 0) msg.push(`Đã thêm mới ${rowsToAdd.length} SKU vào ITEM_MASTER`);
  if (updatedExistingCount > 0) msg.push(`Đã bổ sung Nhóm nguồn = 'INVENTORY' cho ${updatedExistingCount} SKU cũ`);

  if (msg.length === 0) {
    ui.alert("ℹ️ Thông báo", "Tất cả các mã SKU trong ITEM_MASTER đã đầy đủ.", ui.ButtonSet.OK);
  } else {
    ui.alert("✅ Thành công", msg.join("\n"), ui.ButtonSet.OK);
  }
}

/**
 * Core: Đồng bộ MENU -> ITEM_MASTER.
 * - Thêm mới SKU với source_group = "SALES".
 * - Bổ sung source_group cho các dòng cũ từ Menu nếu đang để trống.
 */
function itemMasterSyncFromMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi Schema", "Không thể đọc hoặc cấu trúc sheet SCHEMA không hợp lệ!", ui.ButtonSet.OK);
    return;
  }

  const menuSheetName       = schemaGetSheetName(schemaMap, "MENU");
  const itemMasterSheetName = schemaGetSheetName(schemaMap, "ITEM_MASTER");

  const menuSheet       = menuSheetName ? ss.getSheetByName(menuSheetName) : null;
  const itemMasterSheet = itemMasterSheetName ? ss.getSheetByName(itemMasterSheetName) : null;

  if (!menuSheet || !itemMasterSheet) {
    ui.alert("⚠️ Lỗi", "Thiếu sheet MENU hoặc ITEM_MASTER!", ui.ButtonSet.OK);
    return;
  }

  const colMenuCode     = schemaGetColIndex(schemaMap, "MENU", "menu_code");
  const colMenuName     = schemaGetColIndex(schemaMap, "MENU", "menu_name");
  const colMenuItemType = schemaGetColIndex(schemaMap, "MENU", "item_type");

  if (colMenuCode === -1 || colMenuName === -1 || colMenuItemType === -1) {
    ui.alert("⚠️ Lỗi Schema MENU", "Thiếu cột menu_code, menu_name hoặc item_type trong SCHEMA của MENU!", ui.ButtonSet.OK);
    return;
  }

  const menuData = menuSheet.getDataRange().getValues();
  const validMenuItems = new Map();

  for (let i = 1; i < menuData.length; i++) {
    const row = menuData[i];
    const menuCode = cleanCodeValue_(row[colMenuCode]);
    const menuName = cleanCodeValue_(row[colMenuName]);
    const itemType = cleanCodeValue_(row[colMenuItemType]);

    if (!menuCode || menuCode.toLowerCase().includes("unmapped")) continue;
    if (itemType.toLowerCase() === "dịch vụ" || itemType.toLowerCase() === "dich vu") continue;

    validMenuItems.set(menuCode, menuName || menuCode);
  }

  if (validMenuItems.size === 0) {
    ui.alert("ℹ️ Thông báo", "Không tìm thấy Mã món hợp lệ nào từ MENU để đồng bộ.", ui.ButtonSet.OK);
    return;
  }

  const idxMasterSku    = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_code");
  const idxMasterName   = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_name");
  const idxMasterSource = schemaGetColIndex(schemaMap, "ITEM_MASTER", "source_group");

  if (idxMasterSku === -1 || idxMasterName === -1) {
    ui.alert("⚠️ Lỗi Schema ITEM_MASTER", "Thiếu cột item_code hoặc item_name trong SCHEMA của ITEM_MASTER!", ui.ButtonSet.OK);
    return;
  }

  const masterData = itemMasterSheet.getDataRange().getValues();
  const maxMasterColCount = schemaMap["ITEM_MASTER"].columns.length;

  const existingMasterSkus = new Set();
  let updatedExistingCount = 0;

  // Quét dòng cũ: Cập nhật source_group = "SALES" nếu chưa có
  for (let i = 1; i < masterData.length; i++) {
    const sku = cleanCodeValue_(masterData[i][idxMasterSku]);
    if (sku) {
      existingMasterSkus.add(sku);

      if (validMenuItems.has(sku) && idxMasterSource !== -1) {
        if (!masterData[i][idxMasterSource]) {
          masterData[i][idxMasterSource] = "SALES";
          updatedExistingCount++;
        }
      }
    }
  }

  const rowsToAdd = [];
  validMenuItems.forEach((menuName, menuCode) => {
    if (!existingMasterSkus.has(menuCode)) {
      const newRow = new Array(maxMasterColCount).fill("");
      newRow[idxMasterSku]  = menuCode;
      newRow[idxMasterName] = menuName;
      
      if (idxMasterSource !== -1) {
        newRow[idxMasterSource] = "SALES";
      }

      rowsToAdd.push(newRow);
    }
  });

  if (updatedExistingCount > 0) {
    itemMasterSheet.getRange(1, 1, masterData.length, maxMasterColCount).setValues(masterData);
  }

  if (rowsToAdd.length > 0) {
    const lastRow = itemMasterSheet.getLastRow();
    itemMasterSheet.getRange(lastRow + 1, 1, rowsToAdd.length, maxMasterColCount).setValues(rowsToAdd);
  }

  let msg = [];
  if (rowsToAdd.length > 0) msg.push(`Đã thêm mới ${rowsToAdd.length} Mã món (nguồn SALES)`);
  if (updatedExistingCount > 0) msg.push(`Đã bổ sung source_group = 'SALES' cho ${updatedExistingCount} SKU cũ`);

  if (msg.length === 0) {
    ui.alert("ℹ️ Thông báo", "Tất cả các Mã món từ MENU đều đã đầy đủ thông tin.", ui.ButtonSet.OK);
  } else {
    ui.alert("✅ Thành công", msg.join("\n"), ui.ButtonSet.OK);
  }
}

/**
 * Core: Tự động gợi ý/điền danh mục (item_type, category, base_unit, default_storage) dựa trên ITEM_CLASSIFICATION
 */
function itemMasterAutoSuggestClassification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi Schema", "Không thể đọc cấu trúc sheet SCHEMA!", ui.ButtonSet.OK);
    return;
  }

  const masterSheetName = schemaGetSheetName(schemaMap, "ITEM_MASTER");
  const ruleSchemaName  = schemaMap["ITEM_CLASSIFICATION"] ? "ITEM_CLASSIFICATION" : "ITEM_CLASSIFICATION_RULE";
  const ruleSheetName   = schemaGetSheetName(schemaMap, ruleSchemaName);

  const masterSheet = masterSheetName ? ss.getSheetByName(masterSheetName) : null;
  const ruleSheet   = ruleSheetName ? ss.getSheetByName(ruleSheetName) : null;

  if (!masterSheet || !ruleSheet) {
    ui.alert("⚠️ Lỗi", "Không tìm thấy Sheet ITEM_MASTER hoặc ITEM_CLASSIFICATION!", ui.ButtonSet.OK);
    return;
  }

  // 1. ĐỌC QUY TẮC PHÂN LOẠI TỪ SHEET ITEM_CLASSIFICATION
  const idxRulePattern = schemaGetColIndex(schemaMap, ruleSchemaName, "item_name_pattern");
  const idxRuleGroup   = schemaGetColIndex(schemaMap, ruleSchemaName, "target_source_group");
  const idxRuleType    = schemaGetColIndex(schemaMap, ruleSchemaName, "item_type");
  const idxRuleCat     = schemaGetColIndex(schemaMap, ruleSchemaName, "category");
  const idxRuleUnit    = schemaGetColIndex(schemaMap, ruleSchemaName, "base_unit");
  const idxRuleStorage = schemaGetColIndex(schemaMap, ruleSchemaName, "default_storage");
  const idxRulePrio    = schemaGetColIndex(schemaMap, ruleSchemaName, "priority");

  const ruleData = ruleSheet.getDataRange().getValues();
  const rules = [];

  for (let i = 1; i < ruleData.length; i++) {
    const row = ruleData[i];
    const patternStr  = idxRulePattern !== -1 && row[idxRulePattern] ? row[idxRulePattern].toString().trim() : "";
    const sourceGroup = idxRuleGroup !== -1 && row[idxRuleGroup] ? row[idxRuleGroup].toString().trim().toUpperCase() : "ALL";
    const itemType    = idxRuleType !== -1 && row[idxRuleType] ? row[idxRuleType].toString().trim() : "";
    const category    = idxRuleCat !== -1 && row[idxRuleCat] ? row[idxRuleCat].toString().trim() : "";
    const baseUnit    = idxRuleUnit !== -1 && row[idxRuleUnit] ? row[idxRuleUnit].toString().trim() : "";
    const storage     = idxRuleStorage !== -1 && row[idxRuleStorage] ? row[idxRuleStorage].toString().trim() : "";
    const priority    = idxRulePrio !== -1 && !isNaN(row[idxRulePrio]) ? Number(row[idxRulePrio]) : 999;

    if (patternStr) {
      const patterns = patternStr.split(",").map(p => p.trim()).filter(p => p);
      rules.push({
        regexes: patterns.map(p => patternToRegex_(p)),
        sourceGroup: sourceGroup || "ALL",
        itemType: itemType,
        category: category,
        baseUnit: baseUnit,
        defaultStorage: storage,
        priority: priority
      });
    }
  }

  // Sắp xếp ưu tiên (Priority số nhỏ áp dụng trước)
  rules.sort((a, b) => a.priority - b.priority);

  // 2. TRA CỨU CÁC CỘT CẦN ĐIỀN TRONG ITEM_MASTER
  const idxMasterName    = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_name");
  const idxMasterSource  = schemaGetColIndex(schemaMap, "ITEM_MASTER", "source_group");
  const idxMasterType    = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_type");
  const idxMasterCat     = schemaGetColIndex(schemaMap, "ITEM_MASTER", "category");
  const idxMasterUnit    = schemaGetColIndex(schemaMap, "ITEM_MASTER", "base_unit");
  const idxMasterStorage = schemaGetColIndex(schemaMap, "ITEM_MASTER", "default_storage");

  if (idxMasterName === -1) {
    ui.alert("⚠️ Lỗi Schema", "Chưa định nghĩa col_key 'item_name' trong SCHEMA cho ITEM_MASTER!", ui.ButtonSet.OK);
    return;
  }

  const masterData = masterSheet.getDataRange().getValues();
  let updatedCount = 0;

  // 3. ĐỐI SOÁT QUY TẮC VÀ BỔ SUNG THUỘC TÍNH
  for (let i = 1; i < masterData.length; i++) {
    const itemName = masterData[i][idxMasterName] ? masterData[i][idxMasterName].toString().trim() : "";
    const itemSource = (idxMasterSource !== -1 && masterData[i][idxMasterSource]) 
                       ? masterData[i][idxMasterSource].toString().trim().toUpperCase() 
                       : "";

    if (!itemName) continue;

    for (const rule of rules) {
      // Lọc điều kiện Nhóm nguồn (Nếu quy tắc chỉ áp dụng riêng cho SALES hoặc INVENTORY)
      if (rule.sourceGroup !== "ALL" && itemSource && rule.sourceGroup !== itemSource) {
        continue;
      }

      const isMatched = rule.regexes.some(rx => rx.test(itemName));
      if (isMatched) {
        let isRowChanged = false;

        if (idxMasterType !== -1 && !masterData[i][idxMasterType] && rule.itemType) {
          masterData[i][idxMasterType] = rule.itemType;
          isRowChanged = true;
        }

        if (idxMasterCat !== -1 && !masterData[i][idxMasterCat] && rule.category) {
          masterData[i][idxMasterCat] = rule.category;
          isRowChanged = true;
        }

        if (idxMasterUnit !== -1 && !masterData[i][idxMasterUnit] && rule.baseUnit) {
          masterData[i][idxMasterUnit] = rule.baseUnit;
          isRowChanged = true;
        }

        if (idxMasterStorage !== -1 && !masterData[i][idxMasterStorage] && rule.defaultStorage) {
          masterData[i][idxMasterStorage] = rule.defaultStorage;
          isRowChanged = true;
        }

        if (isRowChanged) updatedCount++;
        break; // Áp dụng quy tắc ưu tiên đầu tiên khớp
      }
    }
  }

  // 4. GHI DỮ LIỆU ĐÃ PHÂN LOẠI XUỐNG SHEET
  if (updatedCount > 0) {
    const maxCols = schemaMap["ITEM_MASTER"].columns.length;
    masterSheet.getRange(1, 1, masterData.length, maxCols).setValues(masterData);
    ui.alert("✅ Hoàn thành", `Đã tự động phân loại thành công cho ${updatedCount} SKU trong ITEM_MASTER!`, ui.ButtonSet.OK);
  } else {
    ui.alert("ℹ️ Thông báo", "Không có SKU nào cần bổ sung phân loại mới.", ui.ButtonSet.OK);
  }
}
