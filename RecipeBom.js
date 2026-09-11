/**
 * Entry Point gọi từ Menu UI để tự động điền/cập nhật Tên hiển thị trong sheet RECIPE_BOM.
 * Quy tắc: [actionEntityDescription] -> runPopulateRecipeBomNames
 */
function runPopulateRecipeBomNames() {
  recipeBomPopulateNames();
}

/**
 * Hàm Global Core điền tên tự động cho RECIPE_BOM qua SCHEMA.
 * Quy tắc: [moduleName][Action][Entity] -> recipeBomPopulateNames
 */
function recipeBomPopulateNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. Lấy cấu trúc SCHEMA chuẩn hóa
  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi Schema", "Không thể đọc hoặc cấu trúc sheet SCHEMA không hợp lệ!", ui.ButtonSet.OK);
    return;
  }

  const bomSheet = ss.getSheetByName("RECIPE_BOM");
  const menuSheet = ss.getSheetByName("MENU");
  const masterSheet = ss.getSheetByName("ITEM_MASTER");

  if (!bomSheet || !menuSheet || !masterSheet) {
    ui.alert("⚠️ Lỗi", "Thiếu một trong các sheet: RECIPE_BOM, MENU, ITEM_MASTER!", ui.ButtonSet.OK);
    return;
  }

  // 2. Tra cứu cột trong RECIPE_BOM qua SCHEMA
  const colParentCode = schemaGetColIndex(schemaMap, "RECIPE_BOM", "parent_code");
  const colParentName = schemaGetColIndex(schemaMap, "RECIPE_BOM", "parent_item_name");
  const colChildCode  = schemaGetColIndex(schemaMap, "RECIPE_BOM", "child_item_code");
  const colChildName  = schemaGetColIndex(schemaMap, "RECIPE_BOM", "child_item_name");

  if (colParentCode === -1 || colChildCode === -1) {
    ui.alert("⚠️ Lỗi Schema", "Thiếu cấu hình parent_code hoặc child_item_code trong SCHEMA của RECIPE_BOM!", ui.ButtonSet.OK);
    return;
  }

  // 3. Nạp danh mục Tên từ MENU và ITEM_MASTER
  const parentNameMap = new Map();
  const childNameMap = new Map();

  // Đọc MENU (parent_code có thể là menu_code)
  const colMenuCode = schemaGetColIndex(schemaMap, "MENU", "menu_code");
  const colMenuName = schemaGetColIndex(schemaMap, "MENU", "menu_name");
  const menuData = menuSheet.getDataRange().getValues();
  for (let i = 1; i < menuData.length; i++) {
    const code = cleanCodeValue_(menuData[i][colMenuCode]);
    const name = cleanCodeValue_(menuData[i][colMenuName]);
    if (code) parentNameMap.set(code, name || code);
  }

  // Đọc ITEM_MASTER (parent_code có thể là BTP, child_item_code luôn là ITEM_MASTER)
  const colMasterCode = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_code");
  const colMasterName = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_name");
  const masterData = masterSheet.getDataRange().getValues();
  for (let i = 1; i < masterData.length; i++) {
    const code = cleanCodeValue_(masterData[i][colMasterCode]);
    const name = cleanCodeValue_(masterData[i][colMasterName]);
    if (code) {
      if (!parentNameMap.has(code)) parentNameMap.set(code, name || code);
      childNameMap.set(code, name || code);
    }
  }

  // 4. Quét RECIPE_BOM và tự động điền Tên
  const bomData = bomSheet.getDataRange().getValues();
  let updatedCount = 0;

  for (let i = 1; i < bomData.length; i++) {
    const pCode = cleanCodeValue_(bomData[i][colParentCode]);
    const cCode = cleanCodeValue_(bomData[i][colChildCode]);

    if (pCode && colParentName !== -1) {
      const pNameTarget = parentNameMap.get(pCode) || "";
      if (bomData[i][colParentName] !== pNameTarget) {
        bomSheet.getRange(i + 1, colParentName + 1).setValue(pNameTarget);
        updatedCount++;
      }
    }

    if (cCode && colChildName !== -1) {
      const cNameTarget = childNameMap.get(cCode) || "";
      if (bomData[i][colChildName] !== cNameTarget) {
        bomSheet.getRange(i + 1, colChildName + 1).setValue(cNameTarget);
        updatedCount++;
      }
    }
  }

  ui.alert("✅ Thành công", `Đã cập nhật tự động tên cho ${updatedCount} ô trong RECIPE_BOM!`, ui.ButtonSet.OK);
}

