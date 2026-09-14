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
