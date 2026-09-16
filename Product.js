/**
 * Đồng bộ danh mục từ SALES sang MENU dựa trên SCHEMA động.
 * PK: menu_code (Lấy từ SALES.item_code)
 * menu_name: Lấy Tên chuẩn từ MAPPING (qua item_code)
 */
/**
 * Đồng bộ danh mục từ SALES sang MENU theo SCHEMA chuẩn hóa.
 * Tự động gán đúng vị trí cột dựa trên col_key trong tbl_schema.
 */
function productSyncSalesToMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. Gọi API lấy cấu trúc SCHEMA chuẩn hóa từ SchemaCore
  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi Schema", "Không thể đọc hoặc cấu trúc sheet SCHEMA không hợp lệ!", ui.ButtonSet.OK);
    return;
  }

  // Lấy sheet_name vật lý động từ SCHEMA
  const mappingSheetName = schemaGetSheetName(schemaMap, "MAPPING");
  const menuSheetName    = schemaGetSheetName(schemaMap, "MENU");
  const salesSheetName   = schemaGetSheetName(schemaMap, "SALES");

  const mappingSheet = mappingSheetName ? ss.getSheetByName(mappingSheetName) : null;
  const menuSheet    = menuSheetName ? ss.getSheetByName(menuSheetName) : null;
  const salesSheet   = salesSheetName ? ss.getSheetByName(salesSheetName) : null;

  if (!mappingSheet || !menuSheet || !salesSheet) {
    ui.alert("⚠️ Lỗi", "Thiếu một trong các sheet hệ thống (MAPPING, MENU, SALES)!", ui.ButtonSet.OK);
    return;
  }

  // 2. Quét SALES lấy danh sách mã duy nhất
  const colSalesItem = schemaGetColIndex(schemaMap, "SALES", "item_code");
  if (colSalesItem === -1) {
    ui.alert("⚠️ Lỗi Schema SALES", "Thiếu cột item_code trong SCHEMA của SALES!", ui.ButtonSet.OK);
    return;
  }

  const validSalesCodes = new Set();
  const salesData = salesSheet.getDataRange().getValues();
  for (let i = 1; i < salesData.length; i++) {
    const rawVal = salesData[i][colSalesItem];
    const itemVal = rawVal !== null && rawVal !== undefined ? String(rawVal).trim() : "";
    
    // Kiểm tra an toàn chuỗi trước khi so sánh
    if (itemVal && !itemVal.toLowerCase().includes("unmapped")) {
      validSalesCodes.add(itemVal);
    }
  }

  if (validSalesCodes.size === 0) {
    ui.alert("ℹ️ Thông báo", "Không tìm thấy mã hợp lệ nào từ bảng SALES.", ui.ButtonSet.OK);
    return;
  }

  // 3. Đọc MAPPING lấy Tên chuẩn hóa (item_name)
  const colMapItemCode = schemaGetColIndex(schemaMap, "MAPPING", "item_code");
  const colMapItemName = schemaGetColIndex(schemaMap, "MAPPING", "item_name");

  if (colMapItemCode === -1) {
    ui.alert("⚠️ Lỗi Schema MAPPING", "Thiếu cột item_code trong SCHEMA của MAPPING!", ui.ButtonSet.OK);
    return;
  }

  const mappingData = mappingSheet.getDataRange().getValues();
  const mappingNameByCode = new Map();
  for (let i = 1; i < mappingData.length; i++) {
    const row = mappingData[i];
    const code = row[colMapItemCode] !== null && row[colMapItemCode] !== undefined ? String(row[colMapItemCode]).trim() : "";
    const name = colMapItemName !== -1 && row[colMapItemName] !== null && row[colMapItemName] !== undefined ? String(row[colMapItemName]).trim() : "";
    if (code) {
      mappingNameByCode.set(code, name || code);
    }
  }

  // 4. Tra cứu vị trí TẤT CẢ các cột trong SCHEMA MENU dựa vào col_key chính xác
  const idxMenuCode       = schemaGetColIndex(schemaMap, "MENU", "menu_code");
  const idxMenuName       = schemaGetColIndex(schemaMap, "MENU", "menu_name");
  const idxItemType       = schemaGetColIndex(schemaMap, "MENU", "item_type");
  const idxCategory       = schemaGetColIndex(schemaMap, "MENU", "category");
  const idxSellingUnit    = schemaGetColIndex(schemaMap, "MENU", "selling_unit");
  const idxSellingPrice   = schemaGetColIndex(schemaMap, "MENU", "selling_price");
  const idxTaxRate        = schemaGetColIndex(schemaMap, "MENU", "tax_rate");
  const idxAllocationRate = schemaGetColIndex(schemaMap, "MENU", "allocation_rate"); // Cột %NVL/Gia vị
  const idxItemCode       = schemaGetColIndex(schemaMap, "MENU", "item_code");
  const idxDepartment     = schemaGetColIndex(schemaMap, "MENU", "department");
  const idxStatus         = schemaGetColIndex(schemaMap, "MENU", "status");

  if (idxMenuCode === -1) {
    ui.alert("⚠️ Lỗi Schema MENU", "Thiếu cột menu_code trong SCHEMA của MENU!", ui.ButtonSet.OK);
    return;
  }

  const menuData = menuSheet.getDataRange().getValues();
  const maxMenuColCount = schemaMap["MENU"].columns.length;

  const existingMenuCodes = new Set();
  for (let i = 1; i < menuData.length; i++) {
    const code = menuData[i][idxMenuCode] !== null && menuData[i][idxMenuCode] !== undefined ? String(menuData[i][idxMenuCode]).trim() : "";
    if (code) existingMenuCodes.add(code);
  }

  // 5. Chuẩn bị dòng dữ liệu chuẩn xác khớp với chỉ số cột thực tế
  const rowsToAdd = [];
  const processedCodes = new Set();

  validSalesCodes.forEach(code => {
    if (!existingMenuCodes.has(code) && !processedCodes.has(code)) {
      processedCodes.add(code);
      const finalName = mappingNameByCode.get(code) || code;

      // Tạo dòng trống chuẩn độ dài theo số lượng cột SCHEMA MENU
      const newRow = new Array(maxMenuColCount).fill("");

      // Điền đúng từng vị trí cột theo schema mapping
      if (idxMenuCode !== -1)       newRow[idxMenuCode]       = code;       // Mã thực đơn
      if (idxMenuName !== -1)       newRow[idxMenuName]       = finalName;  // Tên món bán
      if (idxItemType !== -1)       newRow[idxItemType]       = "";         // Loại thực đơn
      if (idxCategory !== -1)       newRow[idxCategory]       = "";         // Nhóm thực đơn
      if (idxSellingUnit !== -1)    newRow[idxSellingUnit]    = "";         // Đơn vị tính
      if (idxSellingPrice !== -1)   newRow[idxSellingPrice]   = 0;          // Giá bán
      if (idxTaxRate !== -1)        newRow[idxTaxRate]        = 0;          // Thuế suất
      if (idxAllocationRate !== -1) newRow[idxAllocationRate] = 0;          // %NVL/Gia vị
      if (idxItemCode !== -1)       newRow[idxItemCode]       = code;       // Mã SKU liên kết
      if (idxDepartment !== -1)     newRow[idxDepartment]     = "";         // Bộ phận
      if (idxStatus !== -1)         newRow[idxStatus]         = "ACTIVE";   // Trạng thái

      rowsToAdd.push(newRow);
    }
  });

  if (rowsToAdd.length === 0) {
    ui.alert("ℹ️ Thông báo", "Tất cả các mã từ SALES đều đã có sẵn trong bảng MENU.", ui.ButtonSet.OK);
    return;
  }

  // 6. Ghi dữ liệu xuống MENU chính xác theo độ dài cột SCHEMA
  const startRow = Math.max(menuSheet.getLastRow() + 1, 2);
  menuSheet.getRange(startRow, 1, rowsToAdd.length, maxMenuColCount).setValues(rowsToAdd);

  ui.alert("✅ Thành công", `Đã khởi tạo thành công ${rowsToAdd.length} mã món từ SALES sang MENU theo đúng SCHEMA!`, ui.ButtonSet.OK);
}


/**
 * [Global Core] Tính toán chi phí (Cost) thông minh cho các món trên MENU dựa trên:
 * - Tra cứu đơn giá nguyên liệu từ MONTHLY_AVG_PRICE (dạng bảng dọc) theo nguyên tắc ngược thời gian.
 * - Tính tổng NVL chính từ RECIPE_BOM (nhân đúng định lượng quantity).
 * - Khoán chi phí NVL phụ và Gia vị theo tỷ lệ % cấu hình sẵn trên MENU.
 * - Ghi đè kết quả trực tiếp ra sheet MENU.
 * @param {string} targetPeriod Kỳ cần tra cứu và tính toán (VD: '2026-03')
 */
function menuCalculateCost(targetPeriod) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi Schema", "Không thể đọc cấu trúc SCHEMA!", ui.ButtonSet.OK);
    return;
  }

  const menuSheetName = schemaGetSheetName(schemaMap, "MENU");
  const bomSheetName = schemaGetSheetName(schemaMap, "RECIPE_BOM");
  const masterSheetName = schemaGetSheetName(schemaMap, "ITEM_MASTER");
  const avgPriceSheetName = schemaGetSheetName(schemaMap, "MONTHLY_AVG_PRICE") || "MONTHLY_AVG_PRICE";

  const menuSheet = menuSheetName ? ss.getSheetByName(menuSheetName) : null;
  const bomSheet = bomSheetName ? ss.getSheetByName(bomSheetName) : null;
  const masterSheet = masterSheetName ? ss.getSheetByName(masterSheetName) : null;
  const avgPriceSheet = ss.getSheetByName(avgPriceSheetName);

  if (!menuSheet || !bomSheet || !masterSheet) {
    ui.alert("⚠️ Lỗi", "Không tìm thấy đủ các sheet MENU, RECIPE_BOM hoặc ITEM_MASTER theo SCHEMA!", ui.ButtonSet.OK);
    return;
  }

  // 1. Lấy đơn giá cố định dự phòng (cost_price) từ ITEM_MASTER
  const idxMasterCode = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_code");
  const idxMasterCost = schemaGetColIndex(schemaMap, "ITEM_MASTER", "cost_price");

  if (idxMasterCode === -1 || idxMasterCost === -1) {
    ui.alert("⚠️ Lỗi Schema", "ITEM_MASTER thiếu cấu hình item_code hoặc cost_price!", ui.ButtonSet.OK);
    return;
  }

  const fallbackCostMap = new Map();
  const masterData = masterSheet.getDataRange().getValues();
  for (let i = 1; i < masterData.length; i++) {
    const code = masterData[i][idxMasterCode] ? masterData[i][idxMasterCode].toString().trim().toLowerCase() : "";
    const cost = Number(masterData[i][idxMasterCost]) || 0;
    if (code) fallbackCostMap.set(code, cost);
  }

  // 2. Đọc bảng giá theo kỳ (monthly_avg_price dạng bảng dọc)
  const avgPriceMap = new Map();
  let availablePeriods = new Set();

  if (avgPriceSheet) {
    const priceData = avgPriceSheet.getDataRange().getValues();
    const idxPricePeriod = schemaGetColIndex(schemaMap, "MONTHLY_AVG_PRICE", "period");
    const idxPriceCode   = schemaGetColIndex(schemaMap, "MONTHLY_AVG_PRICE", "item_code");
    const idxPriceVal    = schemaGetColIndex(schemaMap, "MONTHLY_AVG_PRICE", "unit_price");

    if (idxPricePeriod !== -1 && idxPriceCode !== -1 && idxPriceVal !== -1) {
      for (let i = 1; i < priceData.length; i++) {
        const row = priceData[i];
        const pStr = row[idxPricePeriod] ? row[idxPricePeriod].toString().trim() : "";
        const cStr = row[idxPriceCode] ? row[idxPriceCode].toString().trim().toLowerCase() : "";
        const uPrice = Number(row[idxPriceVal]) || 0;

        if (pStr && cStr) {
          avgPriceMap.set(`${cStr}|${pStr}`, uPrice);
          availablePeriods.add(pStr);
        }
      }
    }
  }

  let sortedPeriods = Array.from(availablePeriods).sort();

  // Hàm tra cứu ngược thời gian thông minh dựa trên bảng dọc
  function getSmartUnitCostForPeriod(itemCode, period) {
    const codeKey = itemCode.toLowerCase();
    const validPeriods = sortedPeriods.filter(p => p <= period);
    
    for (let i = validPeriods.length - 1; i >= 0; i--) {
      const pKey = validPeriods[i];
      const lookupKey = `${codeKey}|${pKey}`;
      if (avgPriceMap.has(lookupKey)) {
        const price = avgPriceMap.get(lookupKey);
        if (price > 0) return price;
      }
    }
    return fallbackCostMap.get(codeKey) || 0;
  }

  // 3. Tính tổng NVL chính từ RECIPE_BOM theo kỳ (Nhân đúng định lượng quantity)
  const idxBomParentCode = schemaGetColIndex(schemaMap, "RECIPE_BOM", "parent_code");
  const idxBomChildCode  = schemaGetColIndex(schemaMap, "RECIPE_BOM", "child_item_code");
  const idxBomQty        = schemaGetColIndex(schemaMap, "RECIPE_BOM", "quantity");

  if (idxBomParentCode === -1 || idxBomChildCode === -1 || idxBomQty === -1) {
    ui.alert("⚠️ Lỗi Schema", "RECIPE_BOM thiếu cấu hình parent_code, child_item_code hoặc quantity!", ui.ButtonSet.OK);
    return;
  }

  const bomMainCostMap = new Map();
  const bomData = bomSheet.getDataRange().getValues();
  for (let i = 1; i < bomData.length; i++) {
    const pCode = bomData[i][idxBomParentCode] ? bomData[i][idxBomParentCode].toString().trim().toLowerCase() : "";
    const cCode = bomData[i][idxBomChildCode] ? bomData[i][idxBomChildCode].toString().trim().toLowerCase() : "";
    const qty = Number(bomData[i][idxBomQty]) || 0;

    if (pCode && cCode) {
      const unitCost = getSmartUnitCostForPeriod(cCode, targetPeriod);
      const lineCost = qty * unitCost; // Định lượng * Đơn giá
      bomMainCostMap.set(pCode, (bomMainCostMap.get(pCode) || 0) + lineCost);
    }
  }

  // 4. Đọc dữ liệu MENU, tính toán và ghi đè kết quả trực tiếp lên sheet MENU
  const idxMenuCode      = schemaGetColIndex(schemaMap, "MENU", "menu_code");
  const idxMenuMainCost  = schemaGetColIndex(schemaMap, "MENU", "main_material_cost");
  const idxMenuSubRate   = schemaGetColIndex(schemaMap, "MENU", "sub_material_rate");
  const idxMenuSpiceRate = schemaGetColIndex(schemaMap, "MENU", "spice_rate");
  const idxMenuTotalCost = schemaGetColIndex(schemaMap, "MENU", "total_cost");

  if (idxMenuCode === -1 || idxMenuMainCost === -1 || idxMenuTotalCost === -1) {
    ui.alert("⚠️ Lỗi Schema", "MENU thiếu cấu hình các cột tính cost (menu_code, main_material_cost, total_cost)!", ui.ButtonSet.OK);
    return;
  }

  const menuData = menuSheet.getDataRange().getValues();
  if (menuData.length <= 1) {
    ui.alert("ℹ️ Thông báo", "Sheet MENU không có dữ liệu để tính cost.", ui.ButtonSet.OK);
    return;
  }

  let updatedCount = 0;

  for (let i = 1; i < menuData.length; i++) {
    const row = menuData[i];
    const mCode = row[idxMenuCode] ? row[idxMenuCode].toString().trim().toLowerCase() : "";
    if (!mCode) continue;

    const mainCost = bomMainCostMap.get(mCode) || 0;
    row[idxMenuMainCost] = Math.round(mainCost * 100) / 100;

    const subRate = (idxMenuSubRate !== -1) ? (Number(row[idxMenuSubRate]) || 0) : 0;
    const spiceRate = (idxMenuSpiceRate !== -1) ? (Number(row[idxMenuSpiceRate]) || 0) : 0;

    const subCost = mainCost * subRate;
    const spiceCost = mainCost * spiceRate;
    const totalCost = mainCost + subCost + spiceCost;

    row[idxMenuTotalCost] = Math.round(totalCost * 100) / 100;
    updatedCount++;
  }

  const maxColCount = schemaMap["MENU"].columns.length;
  menuSheet.getRange(1, 1, menuData.length, maxColCount).setValues(menuData);

  ui.alert(
    "✅ Thành công", 
    `Đã tính toán chi phí (Cost) cho kỳ [${targetPeriod}] trên MENU thành công cho ${updatedCount} món!`, 
    ui.ButtonSet.OK
  );
}




/**
 * [Global Core - Debug & Optimized] Tính toán và lưu snapshot chi phí (Cost) cho các món theo kỳ vào sheet MENU_SNAPSHOT.
 * Có bổ sung log kiểm tra chi tiết quá trình tra cứu đơn giá nguyên liệu.
 * @param {string} targetPeriod Kỳ báo cáo cần chốt (VD: '2026-08')
 */
function menuSaveCostSnapshot(targetPeriod) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi Schema", "Không thể đọc cấu trúc SCHEMA!", ui.ButtonSet.OK);
    return;
  }

  const menuSheetName = schemaGetSheetName(schemaMap, "MENU");
  const bomSheetName = schemaGetSheetName(schemaMap, "RECIPE_BOM");
  const masterSheetName = schemaGetSheetName(schemaMap, "ITEM_MASTER");
  const snapshotSheetName = schemaGetSheetName(schemaMap, "MENU_SNAPSHOT") || "MENU_SNAPSHOT";
  const avgPriceSheetName = schemaGetSheetName(schemaMap, "MONTHLY_AVG_PRICE") || "MONTHLY_AVG_PRICE";

  const menuSheet = menuSheetName ? ss.getSheetByName(menuSheetName) : null;
  const bomSheet = bomSheetName ? ss.getSheetByName(bomSheetName) : null;
  const masterSheet = masterSheetName ? ss.getSheetByName(masterSheetName) : null;
  const avgPriceSheet = ss.getSheetByName(avgPriceSheetName);
  
  let snapshotSheet = ss.getSheetByName(snapshotSheetName);
  if (!snapshotSheet) {
    snapshotSheet = ss.insertSheet(snapshotSheetName);
    const cols = schemaMap["MENU_SNAPSHOT"].columns.map(c => c.col_header);
    snapshotSheet.appendRow(cols);
  }

  if (!menuSheet || !bomSheet || !masterSheet) {
    ui.alert("⚠️ Lỗi", "Không tìm thấy đủ các sheet MENU, RECIPE_BOM hoặc ITEM_MASTER theo SCHEMA!", ui.ButtonSet.OK);
    return;
  }

  // 1. Lấy đơn giá cố định dự phòng (cost_price) từ ITEM_MASTER
  const idxMasterCode = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_code");
  const idxMasterCost = schemaGetColIndex(schemaMap, "ITEM_MASTER", "cost_price");

  const fallbackCostMap = new Map();
  const masterData = masterSheet.getDataRange().getValues();
  for (let i = 1; i < masterData.length; i++) {
    const code = masterData[i][idxMasterCode] ? masterData[i][idxMasterCode].toString().trim().toLowerCase() : "";
    const cost = Number(masterData[i][idxMasterCost]) || 0;
    if (code) fallbackCostMap.set(code, cost);
  }

  // 2. Đọc bảng giá theo kỳ (monthly_avg_price dạng bảng dọc)
  const avgPriceMap = new Map();
  let availablePeriods = new Set();

  if (avgPriceSheet) {
    const priceData = avgPriceSheet.getDataRange().getValues();
    const idxPricePeriod = schemaGetColIndex(schemaMap, "MONTHLY_AVG_PRICE", "period");
    const idxPriceCode   = schemaGetColIndex(schemaMap, "MONTHLY_AVG_PRICE", "item_code");
    const idxPriceVal    = schemaGetColIndex(schemaMap, "MONTHLY_AVG_PRICE", "unit_price");

    if (idxPricePeriod !== -1 && idxPriceCode !== -1 && idxPriceVal !== -1) {
      for (let i = 1; i < priceData.length; i++) {
        const row = priceData[i];
        const pStr = row[idxPricePeriod] ? row[idxPricePeriod].toString().trim() : "";
        const cStr = row[idxPriceCode] ? row[idxPriceCode].toString().trim().toLowerCase() : "";
        const uPrice = Number(row[idxPriceVal]) || 0;

        if (pStr && cStr) {
          avgPriceMap.set(`${cStr}|${pStr}`, uPrice);
          availablePeriods.add(pStr);
        }
      }
    }
  }

  let sortedPeriods = Array.from(availablePeriods).sort();
  Logger.log("Các kỳ có trong MONTHLY_AVG_PRICE: " + JSON.stringify(sortedPeriods));

  // Hàm tra cứu ngược thời gian thông minh với cơ chế log debug
  function getSmartUnitCostForPeriod(itemCode, period) {
    const codeKey = itemCode.toLowerCase();
    const validPeriods = sortedPeriods.filter(p => p <= period);
    
    // Duyệt ngược từ kỳ gần nhất trở về trước
    for (let i = validPeriods.length - 1; i >= 0; i--) {
      const pKey = validPeriods[i];
      const lookupKey = `${codeKey}|${pKey}`;
      if (avgPriceMap.has(lookupKey)) {
        const price = avgPriceMap.get(lookupKey);
        if (price > 0) {
          Logger.log(`[FOUND] Mã NVL: ${itemCode} lấy giá từ kỳ ${pKey} = ${price}`);
          return price;
        }
      }
    }
    
    // Nếu không có trong bảng giá theo kỳ, lấy fallback từ ITEM_MASTER
    const fallbackPrice = fallbackCostMap.get(codeKey) || 0;
    Logger.log(`[FALLBACK] Mã NVL: ${itemCode} lấy giá cố định từ ITEM_MASTER = ${fallbackPrice}`);
    return fallbackPrice;
  }

  // 3. Tính tổng NVL chính từ RECIPE_BOM theo kỳ
  const idxBomParentCode = schemaGetColIndex(schemaMap, "RECIPE_BOM", "parent_code");
  const idxBomChildCode  = schemaGetColIndex(schemaMap, "RECIPE_BOM", "child_item_code");
  const idxBomQty        = schemaGetColIndex(schemaMap, "RECIPE_BOM", "quantity");

  const bomMainCostMap = new Map();
  const bomData = bomSheet.getDataRange().getValues();
  for (let i = 1; i < bomData.length; i++) {
    const pCode = bomData[i][idxBomParentCode] ? bomData[i][idxBomParentCode].toString().trim().toLowerCase() : "";
    const cCode = bomData[i][idxBomChildCode] ? bomData[i][idxBomChildCode].toString().trim().toLowerCase() : "";
    const qty = Number(bomData[i][idxBomQty]) || 0;

    if (pCode && cCode) {
      const unitCost = getSmartUnitCostForPeriod(cCode, targetPeriod);
      const lineCost = qty * unitCost;
      bomMainCostMap.set(pCode, (bomMainCostMap.get(pCode) || 0) + lineCost);
    }
  }

  // 4. Đọc Master MENU lấy cấu hình món và tỷ lệ khoán
  const idxMenuCode      = schemaGetColIndex(schemaMap, "MENU", "menu_code");
  const idxMenuName      = schemaGetColIndex(schemaMap, "MENU", "menu_name");
  const idxMenuUom       = schemaGetColIndex(schemaMap, "MENU", "selling_unit");
  const idxMenuSubRate   = schemaGetColIndex(schemaMap, "MENU", "sub_material_rate");
  const idxMenuSpiceRate = schemaGetColIndex(schemaMap, "MENU", "spice_rate");

  if (idxMenuCode === -1) {
    ui.alert("⚠️ Lỗi Schema", "MENU thiếu cấu hình menu_code!", ui.ButtonSet.OK);
    return;
  }

  const menuData = menuSheet.getDataRange().getValues();
  if (menuData.length <= 1) {
    ui.alert("ℹ️ Thông báo", "Sheet MENU không có dữ liệu.", ui.ButtonSet.OK);
    return;
  }

  // 5. Đọc snapshot hiện tại để xử lý Upsert
  const snapshotData = snapshotSheet.getDataRange().getValues();
  const snapshotMap = new Map();
  
  const idxSnapPeriod = schemaGetColIndex(schemaMap, "MENU_SNAPSHOT", "period");
  const idxSnapCode   = schemaGetColIndex(schemaMap, "MENU_SNAPSHOT", "menu_code");

  for (let i = 1; i < snapshotData.length; i++) {
    const sPeriod = snapshotData[i][idxSnapPeriod] ? snapshotData[i][idxSnapPeriod].toString().trim() : "";
    const sCode = snapshotData[i][idxSnapCode] ? snapshotData[i][idxSnapCode].toString().trim().toLowerCase() : "";
    if (sPeriod && sCode) {
      snapshotMap.set(`${sPeriod}|${sCode}`, i + 1);
    }
  }

  const maxColCount = schemaMap["MENU_SNAPSHOT"].columns.length;
  let addedCount = 0;
  let updatedCount = 0;
  const rowsToAppend = [];

  // 6. Duyệt MENU, tính toán giá trị
  for (let i = 1; i < menuData.length; i++) {
    const row = menuData[i];
    const mCode = row[idxMenuCode] ? row[idxMenuCode].toString().trim() : "";
    if (!mCode) continue;

    const mCodeLower = mCode.toLowerCase();
    const mName = idxMenuName !== -1 ? row[idxMenuName] : "";
    const mUom = idxMenuUom !== -1 ? row[idxMenuUom] : "";
    const subRate = idxMenuSubRate !== -1 ? (Number(row[idxMenuSubRate]) || 0) : 0;
    const spiceRate = idxMenuSpiceRate !== -1 ? (Number(row[idxMenuSpiceRate]) || 0) : 0;

    const mainCost = bomMainCostMap.get(mCodeLower) || 0;
    const subCost = mainCost * subRate;
    const spiceCost = mainCost * spiceRate;
    const totalCost = mainCost + subCost + spiceCost;

    const snapshotRow = new Array(maxColCount).fill("");
    snapshotRow[schemaGetColIndex(schemaMap, "MENU_SNAPSHOT", "period")] = targetPeriod;
    snapshotRow[schemaGetColIndex(schemaMap, "MENU_SNAPSHOT", "menu_code")] = mCode;
    snapshotRow[schemaGetColIndex(schemaMap, "MENU_SNAPSHOT", "menu_name")] = mName;
    snapshotRow[schemaGetColIndex(schemaMap, "MENU_SNAPSHOT", "menu_uom")] = mUom;
    snapshotRow[schemaGetColIndex(schemaMap, "MENU_SNAPSHOT", "main_material_cost")] = Math.round(mainCost * 100) / 100;
    snapshotRow[schemaGetColIndex(schemaMap, "MENU_SNAPSHOT", "sub_material_rate")] = subRate;
    snapshotRow[schemaGetColIndex(schemaMap, "MENU_SNAPSHOT", "spice_rate")] = spiceRate;
    snapshotRow[schemaGetColIndex(schemaMap, "MENU_SNAPSHOT", "total_cost")] = Math.round(totalCost * 100) / 100;

    const key = `${targetPeriod}|${mCodeLower}`;
    if (snapshotMap.has(key)) {
      const rowIndex = snapshotMap.get(key);
      snapshotSheet.getRange(rowIndex, 1, 1, maxColCount).setValues([snapshotRow]);
      updatedCount++;
    } else {
      rowsToAppend.push(snapshotRow);
      addedCount++;
    }
  }

  if (rowsToAppend.length > 0) {
    snapshotSheet.getRange(snapshotSheet.getLastRow() + 1, 1, rowsToAppend.length, maxColCount).setValues(rowsToAppend);
  }

  ui.alert(
    "✅ Chốt Snapshot Thành Công", 
    `Kỳ báo cáo: [${targetPeriod}]\n• Thêm mới: ${addedCount} món\n• Cập nhật: ${updatedCount} món`, 
    ui.ButtonSet.OK
  );
}
