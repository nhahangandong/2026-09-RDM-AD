/**
 * [Global Core] Khởi tạo danh sách các món/BTP có has_bom = TRUE từ ITEM_MASTER sang RECIPE_BOM.
 * Nguyên tắc: Chỉ thêm mới parent_code nếu nó hoàn toàn chưa xuất hiện trong RECIPE_BOM.
 */
function recipeBomInitFromItemMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi Schema", "Không thể đọc cấu trúc SCHEMA!", ui.ButtonSet.OK);
    return;
  }

  const masterSheetName = schemaGetSheetName(schemaMap, "ITEM_MASTER");
  const bomSheetName = schemaGetSheetName(schemaMap, "RECIPE_BOM");

  const masterSheet = masterSheetName ? ss.getSheetByName(masterSheetName) : null;
  const bomSheet = bomSheetName ? ss.getSheetByName(bomSheetName) : null;

  if (!masterSheet || !bomSheet) {
    ui.alert("⚠️ Lỗi", "Không tìm thấy sheet ITEM_MASTER hoặc RECIPE_BOM theo SCHEMA!", ui.ButtonSet.OK);
    return;
  }

  const idxMasterCode   = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_code");
  const idxMasterName   = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_name");
  const idxMasterHasBom = schemaGetColIndex(schemaMap, "ITEM_MASTER", "has_bom");

  if (idxMasterCode === -1 || idxMasterName === -1 || idxMasterHasBom === -1) {
    ui.alert("⚠️ Lỗi Schema", "ITEM_MASTER thiếu cấu hình cột 'item_code', 'item_name' hoặc 'has_bom'!", ui.ButtonSet.OK);
    return;
  }

  const idxBomParentCode = schemaGetColIndex(schemaMap, "RECIPE_BOM", "parent_code");
  const idxBomParentName = schemaGetColIndex(schemaMap, "RECIPE_BOM", "parent_item_name");

  if (idxBomParentCode === -1) {
    ui.alert("⚠️ Lỗi Schema", "RECIPE_BOM thiếu cấu hình cột 'parent_code'!", ui.ButtonSet.OK);
    return;
  }

  const maxColCount = schemaMap["RECIPE_BOM"].columns.length;

  const masterData = masterSheet.getDataRange().getValues();
  if (masterData.length <= 1) {
    ui.alert("ℹ️ Thông báo", "Sheet ITEM_MASTER không có dữ liệu!", ui.ButtonSet.OK);
    return;
  }

  const itemsToInit = [];
  for (let i = 1; i < masterData.length; i++) {
    const code = masterData[i][idxMasterCode] ? masterData[i][idxMasterCode].toString().trim() : "";
    const name = masterData[i][idxMasterName] ? masterData[i][idxMasterName].toString().trim() : "";
    const hasBomVal = masterData[i][idxMasterHasBom];

    const isHasBom = (hasBomVal === true || (typeof hasBomVal === 'string' && hasBomVal.trim().toUpperCase() === 'TRUE') || hasBomVal === 1);

    if (code && isHasBom) {
      itemsToInit.push({ code: code, name: name });
    }
  }

  if (itemsToInit.length === 0) {
    ui.alert("ℹ️ Thông báo", "Không tìm thấy item nào trong ITEM_MASTER có has_bom = TRUE!", ui.ButtonSet.OK);
    return;
  }

  const bomData = bomSheet.getDataRange().getValues();
  const existingParentCodes = new Set();
  
  for (let i = 1; i < bomData.length; i++) {
    const pCode = bomData[i][idxBomParentCode] ? bomData[i][idxBomParentCode].toString().trim().toLowerCase() : "";
    if (pCode) existingParentCodes.add(pCode);
  }

  const newRowsToAppend = [];
  itemsToInit.forEach(item => {
    if (!existingParentCodes.has(item.code.toLowerCase())) {
      const newRow = new Array(maxColCount).fill("");
      newRow[idxBomParentCode] = item.code;
      if (idxBomParentName !== -1) {
        newRow[idxBomParentName] = item.name;
      }
      newRowsToAppend.push(newRow);
      existingParentCodes.add(item.code.toLowerCase());
    }
  });

  if (newRowsToAppend.length > 0) {
    const lastRow = bomSheet.getLastRow();
    bomSheet.getRange(lastRow + 1, 1, newRowsToAppend.length, maxColCount).setValues(newRowsToAppend);
    
    ui.alert(
      "✅ Hoàn thành", 
      `Đã khởi tạo ${newRowsToAppend.length} mã món/BTP mới có has_bom = TRUE vào RECIPE_BOM.`, 
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      "ℹ️ Thông báo", 
      "Tất cả các món có has_bom = TRUE đều đã có mặt ở cột parent_code của RECIPE_BOM.", 
      ui.ButtonSet.OK
    );
  }
}

/**
 * [Global Core] Tự động điền/cập nhật Tên hiển thị (parent_item_name, child_item_name) 
 * và Đơn vị chuẩn (unit) cho thành phần từ ITEM_MASTER vào sheet RECIPE_BOM.
 */
function recipeBomPopulateNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi Schema", "Không thể đọc cấu trúc SCHEMA!", ui.ButtonSet.OK);
    return;
  }

  const bomSheetName = schemaGetSheetName(schemaMap, "RECIPE_BOM");
  const menuSheetName = schemaGetSheetName(schemaMap, "MENU");
  const masterSheetName = schemaGetSheetName(schemaMap, "ITEM_MASTER");

  const bomSheet = bomSheetName ? ss.getSheetByName(bomSheetName) : null;
  const menuSheet = menuSheetName ? ss.getSheetByName(menuSheetName) : null;
  const masterSheet = masterSheetName ? ss.getSheetByName(masterSheetName) : null;

  if (!bomSheet || !masterSheet) {
    ui.alert("⚠️ Lỗi", "Không tìm thấy sheet RECIPE_BOM hoặc ITEM_MASTER theo SCHEMA!", ui.ButtonSet.OK);
    return;
  }

  const colParentCode = schemaGetColIndex(schemaMap, "RECIPE_BOM", "parent_code");
  const colParentName = schemaGetColIndex(schemaMap, "RECIPE_BOM", "parent_item_name");
  const colChildCode  = schemaGetColIndex(schemaMap, "RECIPE_BOM", "child_item_code");
  const colChildName  = schemaGetColIndex(schemaMap, "RECIPE_BOM", "child_item_name");
  const colBomUnit    = schemaGetColIndex(schemaMap, "RECIPE_BOM", "base_unit"); // Cột đơn vị trong BOM

  if (colParentCode === -1 || colChildCode === -1) {
    ui.alert("⚠️ Lỗi Schema", "Thiếu cấu hình parent_code hoặc child_item_code trong SCHEMA của RECIPE_BOM!", ui.ButtonSet.OK);
    return;
  }

  const parentNameMap = new Map();
  const childNameMap = new Map();
  const childUnitMap = new Map(); // Map lưu đơn vị chuẩn của nguyên liệu từ ITEM_MASTER

  if (menuSheet) {
    const colMenuCode = schemaGetColIndex(schemaMap, "MENU", "menu_code");
    const colMenuName = schemaGetColIndex(schemaMap, "MENU", "menu_name");
    if (colMenuCode !== -1 && colMenuName !== -1) {
      const menuData = menuSheet.getDataRange().getValues();
      for (let i = 1; i < menuData.length; i++) {
        const code = menuData[i][colMenuCode] ? menuData[i][colMenuCode].toString().trim() : "";
        const name = menuData[i][colMenuName] ? menuData[i][colMenuName].toString().trim() : "";
        if (code) parentNameMap.set(code, name || code);
      }
    }
  }

  const colMasterCode = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_code");
  const colMasterName = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_name");
  const colMasterUnit = schemaGetColIndex(schemaMap, "ITEM_MASTER", "base_unit"); // Cột đơn vị trong ITEM_MASTER

  if (colMasterCode !== -1 && colMasterName !== -1) {
    const masterData = masterSheet.getDataRange().getValues();
    for (let i = 1; i < masterData.length; i++) {
      const code = masterData[i][colMasterCode] ? masterData[i][colMasterCode].toString().trim() : "";
      const name = masterData[i][colMasterName] ? masterData[i][colMasterName].toString().trim() : "";
      const unit = colMasterUnit !== -1 && masterData[i][colMasterUnit] ? masterData[i][colMasterUnit].toString().trim() : "";
      
      if (code) {
        if (!parentNameMap.has(code)) parentNameMap.set(code, name || code);
        childNameMap.set(code, name || code);
        if (unit) childUnitMap.set(code, unit);
      }
    }
  }

  const bomData = bomSheet.getDataRange().getValues();
  if (bomData.length <= 1) {
    ui.alert("ℹ️ Thông báo", "Sheet RECIPE_BOM chưa có dữ liệu định lượng nào!", ui.ButtonSet.OK);
    return;
  }

  let updatedCount = 0;

  for (let i = 1; i < bomData.length; i++) {
    const pCode = bomData[i][colParentCode] ? bomData[i][colParentCode].toString().trim() : "";
    const cCode = bomData[i][colChildCode] ? bomData[i][colChildCode].toString().trim() : "";

    // 1. Cập nhật tên parent
    if (pCode && colParentName !== -1) {
      const pNameTarget = parentNameMap.get(pCode) || "";
      if (bomData[i][colParentName] !== pNameTarget) {
        bomData[i][colParentName] = pNameTarget;
        updatedCount++;
      }
    }

    // 2. Cập nhật tên child
    if (cCode && colChildName !== -1) {
      const cNameTarget = childNameMap.get(cCode) || "";
      if (bomData[i][colChildName] !== cNameTarget) {
        bomData[i][colChildName] = cNameTarget;
        updatedCount++;
      }
    }

    // 3. [MỚI] Cập nhật đơn vị chuẩn cho thành phần (child) nếu cột unit trống hoặc lệch
    if (cCode && colBomUnit !== -1) {
      const targetUnit = childUnitMap.get(cCode) || "";
      if (targetUnit && bomData[i][colBomUnit] !== targetUnit) {
        bomData[i][colBomUnit] = targetUnit;
        updatedCount++;
      }
    }
  }

  if (updatedCount > 0) {
    const maxColCount = schemaMap["RECIPE_BOM"].columns.length;
    bomSheet.getRange(1, 1, bomData.length, maxColCount).setValues(bomData);
    ui.alert("✅ Thành công", `Đã cập nhật tự động tên và đơn vị cho ${updatedCount} vị trí trong RECIPE_BOM!`, ui.ButtonSet.OK);
  } else {
    ui.alert("ℹ️ Thông báo", "Tên và đơn vị hiển thị trong RECIPE_BOM đã khớp hoàn toàn, không cần cập nhật.", ui.ButtonSet.OK);
  }
}

/**
 * [Global Core] Dọn dẹp RECIPE_BOM:
 * 1. Loại bỏ các dòng mà parent_code KHÔNG CÒN có has_bom = TRUE trong ITEM_MASTER.
 * 2. Loại bỏ các dòng trùng lặp dựa trên cặp khóa chính (Composite PK: parent_code + child_item_code).
 */
function recipeBomCleanDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi Schema", "Không thể đọc cấu trúc SCHEMA!", ui.ButtonSet.OK);
    return;
  }

  const bomSheetName = schemaGetSheetName(schemaMap, "RECIPE_BOM");
  const masterSheetName = schemaGetSheetName(schemaMap, "ITEM_MASTER");

  const bomSheet = bomSheetName ? ss.getSheetByName(bomSheetName) : null;
  const masterSheet = masterSheetName ? ss.getSheetByName(masterSheetName) : null;

  if (!bomSheet || !masterSheet) {
    ui.alert("⚠️ Lỗi", "Không tìm thấy sheet RECIPE_BOM hoặc ITEM_MASTER theo SCHEMA!", ui.ButtonSet.OK);
    return;
  }

  const idxParentCode = schemaGetColIndex(schemaMap, "RECIPE_BOM", "parent_code");
  const idxChildCode  = schemaGetColIndex(schemaMap, "RECIPE_BOM", "child_item_code");

  const idxMasterCode   = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_code");
  const idxMasterHasBom = schemaGetColIndex(schemaMap, "ITEM_MASTER", "has_bom");

  if (idxParentCode === -1 || idxChildCode === -1 || idxMasterCode === -1 || idxMasterHasBom === -1) {
    ui.alert("⚠️ Lỗi Schema", "Thiếu cấu hình cột khóa chính hoặc has_bom trong SCHEMA!", ui.ButtonSet.OK);
    return;
  }

  // 1. Nạp danh sách các item hợp lệ có has_bom = TRUE từ ITEM_MASTER
  const validHasBomSet = new Set();
  const masterData = masterSheet.getDataRange().getValues();
  for (let i = 1; i < masterData.length; i++) {
    const code = masterData[i][idxMasterCode] ? masterData[i][idxMasterCode].toString().trim().toLowerCase() : "";
    const hasBomVal = masterData[i][idxMasterHasBom];
    const isHasBom = (hasBomVal === true || (typeof hasBomVal === 'string' && hasBomVal.trim().toUpperCase() === 'TRUE') || hasBomVal === 1);
    
    if (code && isHasBom) {
      validHasBomSet.add(code);
    }
  }

  const maxColCount = schemaMap["RECIPE_BOM"].columns.length;
  const bomData = bomSheet.getDataRange().getValues();
  
  if (bomData.length <= 1) {
    ui.alert("ℹ️ Thông báo", "Sheet RECIPE_BOM không có dữ liệu để làm sạch.", ui.ButtonSet.OK);
    return;
  }

  const headerRow = bomData[0];
  const uniqueRowsMap = new Map();
  let removedNoBomCount = 0;
  let duplicateCount = 0;

  // 2. Duyệt qua từng dòng BOM để lọc
  for (let i = 1; i < bomData.length; i++) {
    const row = bomData[i];
    const pCode = row[idxParentCode] ? row[idxParentCode].toString().trim().toLowerCase() : "";
    const cCode = row[idxChildCode] ? row[idxChildCode].toString().trim().toLowerCase() : "";

    // Bỏ qua dòng trống hoàn toàn
    if (!pCode && !cCode) continue;

    // [MỚI] Kiểm tra: Nếu parent_code không tồn tại hoặc không còn has_bom = TRUE trong ITEM_MASTER -> Xóa loại bỏ
    if (pCode && !validHasBomSet.has(pCode)) {
      removedNoBomCount++;
      continue; 
    }

    // Kiểm tra trùng lặp theo cặp khóa chính (Composite PK: parent_code + child_item_code)
    const compositePK = `${pCode}|${cCode}`;

    if (!uniqueRowsMap.has(compositePK)) {
      uniqueRowsMap.set(compositePK, row);
    } else {
      duplicateCount++; // Đã tồn tại cặp PK này -> đếm là trùng lặp cần loại bỏ
    }
  }

  const totalRemoved = removedNoBomCount + duplicateCount;

  // 3. Ghi đè lại dữ liệu sạch vào sheet
  if (totalRemoved > 0) {
    const cleanedRows = [headerRow, ...Array.from(uniqueRowsMap.values())];
    
    bomSheet.getDataRange().clearContent();
    bomSheet.getRange(1, 1, cleanedRows.length, maxColCount).setValues(cleanedRows);

    ui.alert(
      "✅ Hoàn thành dọn dẹp", 
      `Đã làm sạch RECIPE_BOM thành công:\n• Xóa ${removedNoBomCount} dòng có món không bật has_bom.\n• Xóa ${duplicateCount} dòng trùng lặp cặp khóa (parent + child).`, 
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      "ℹ️ Thông báo", 
      "Dữ liệu trong RECIPE_BOM hoàn toàn sạch: Các món đều có has_bom hợp lệ và không có cặp khóa nào bị trùng.", 
      ui.ButtonSet.OK
    );
  }
}
