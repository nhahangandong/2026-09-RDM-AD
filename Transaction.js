/**
 * ============================================================================
 * MODULE: Transaction.gs
 * MỤC ĐÍCH: Xử lý logic nghiệp vụ cho các giao dịch TRANSACTION.
 * Quy tắc đặt tên: transactionActionEntity (CamelCase)
 * ============================================================================
 */

/**
 * Core: Tra cứu và điền mã SKU (item_code) từ sheet MAPPING quay trở lại các sheet giao dịch.
 * Khóa tra cứu Composite Key: raw_name + source_group
 */
function transactionAssignItemCodeToAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. LẤY CẤU TRÚC SCHEMA
  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi", "Không thể đọc cấu trúc từ sheet 'SCHEMA'!", ui.ButtonSet.OK);
    return;
  }

  const mappingSheetName = schemaGetSheetName(schemaMap, "MAPPING");
  const mappingSheet = mappingSheetName ? ss.getSheetByName(mappingSheetName) : null;
  if (!mappingSheet) {
    ui.alert("⚠️ Thông báo", "Không tìm thấy sheet MAPPING!", ui.ButtonSet.OK);
    return;
  }

  const idxMappingRaw    = schemaGetColIndex(schemaMap, "MAPPING", "raw_name");
  const idxMappingSource = schemaGetColIndex(schemaMap, "MAPPING", "source");
  const idxMappingCode   = schemaGetColIndex(schemaMap, "MAPPING", "item_code");

  if (idxMappingRaw === -1 || idxMappingCode === -1) {
    ui.alert("⚠️ Lỗi SCHEMA", "Sheet MAPPING thiếu col_key 'raw_name' hoặc 'item_code'!", ui.ButtonSet.OK);
    return;
  }

  // 2. LẤY MAP ÁNH XẠ NHÓM NGUỒN TỪ TBL_RAW_SOURCE_GROUP
  const dynamicSourceGroupMap = getSourceGroupMap_(schemaMap, ss);

  // 3. DỰNG MAP TRA CỨU MÃ SKU THEO COMPOSITE KEY (raw_name + source_group)
  const mappingData = mappingSheet.getDataRange().getValues();
  const codeLookupMap = new Map();

  for (let i = 1; i < mappingData.length; i++) {
    const rawVal = mappingData[i][idxMappingRaw] ? mappingData[i][idxMappingRaw].toString().trim().toLowerCase() : "";
    const sourceVal = idxMappingSource !== -1 && mappingData[i][idxMappingSource] ? mappingData[i][idxMappingSource].toString().trim().toLowerCase() : "";
    const codeVal = mappingData[i][idxMappingCode] ? mappingData[i][idxMappingCode].toString().trim() : "";

    if (rawVal && codeVal) {
      // Tìm nhóm nguồn tương ứng từ bảng tbl_raw_source_group.
      // Nếu không khai báo trong bảng thì fallback giữ nguyên giá trị sourceVal.
      const sourceGroup = dynamicSourceGroupMap.get(sourceVal) || sourceVal;
      
      const compositeKey = `${rawVal}|${sourceGroup}`;
      codeLookupMap.set(compositeKey, codeVal);
    }
  }

  if (codeLookupMap.size === 0) {
    ui.alert("ℹ️ Thông báo", "Sheet MAPPING chưa có dữ liệu Mã SKU hợp lệ!", ui.ButtonSet.OK);
    return;
  }

  let totalUpdatedSheets = 0;
  let totalUpdatedRows = 0;

  // 4. DUYỆT TẤT CẢ CÁC SHEET GIAO DỊCH RAW
  Object.keys(schemaMap).forEach(schemaName => {
    // Bỏ qua các bảng hệ thống
    if (schemaName === "MAPPING" || schemaName === "SCHEMA" || schemaName === "RAW_SOURCE_GROUP" || schemaName === "ITEM_MASTER") return;

    const idxSourceRaw  = schemaGetColIndex(schemaMap, schemaName, "raw_name");
    const idxSourceCode = schemaGetColIndex(schemaMap, schemaName, "item_code");

    if (idxSourceRaw === -1 || idxSourceCode === -1) return;

    const targetSheetName = schemaGetSheetName(schemaMap, schemaName);
    const targetSheet = targetSheetName ? ss.getSheetByName(targetSheetName) : null;
    if (!targetSheet) return;

    const lastRow = targetSheet.getLastRow();
    if (lastRow <= 1) return;

    const rangeRaw = targetSheet.getRange(2, idxSourceRaw + 1, lastRow - 1, 1).getValues();
    const rangeCode = targetSheet.getRange(2, idxSourceCode + 1, lastRow - 1, 1).getValues();

    // Tra cứu source_group cho Sheet hiện tại theo schema_name hoặc sheet_name
    const currentSchemaLower = schemaName.toLowerCase();
    const currentSheetLower  = targetSheetName.toLowerCase();
    const currentSourceGroup = dynamicSourceGroupMap.get(currentSchemaLower) || dynamicSourceGroupMap.get(currentSheetLower) || currentSchemaLower;

    let sheetUpdatedCount = 0;

    for (let i = 0; i < rangeRaw.length; i++) {
      const rawVal = rangeRaw[i][0] ? rangeRaw[i][0].toString().trim().toLowerCase() : "";
      if (!rawVal) continue;

      const targetCompositeKey = `${rawVal}|${currentSourceGroup}`;

      if (codeLookupMap.has(targetCompositeKey)) {
        const matchedCode = codeLookupMap.get(targetCompositeKey);
        const currentCode = rangeCode[i][0] ? rangeCode[i][0].toString().trim() : "";

        if (currentCode !== matchedCode) {
          rangeCode[i][0] = matchedCode;
          sheetUpdatedCount++;
        }
      }
    }

    // 5. GHI DỮ LIỆU CỘT ITEM_CODE XUỐNG SHEET
    if (sheetUpdatedCount > 0) {
      const colCode1Based = idxSourceCode + 1;
      targetSheet.getRange(2, colCode1Based, rangeCode.length, 1).setValues(rangeCode);

      totalUpdatedSheets++;
      totalUpdatedRows += sheetUpdatedCount;
    }
  });

  // 6. THÔNG BÁO KẾT QUẢ
  if (totalUpdatedRows > 0) {
    ui.alert(
      "✅ Hoàn thành",
      `Đã cập nhật thành công ${totalUpdatedRows} dòng mã SKU trên ${totalUpdatedSheets} sheet giao dịch.`,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      "ℹ️ Thông báo",
      "Dữ liệu mã SKU trên các sheet giao dịch đã khớp hoàn toàn, không có dòng nào cần cập nhật mới.",
      ui.ButtonSet.OK
    );
  }
}


/**
 * Hàm hỗ trợ tự động lấy hoặc tính toán Period (YYYYMM)
 * @param {Date|string} transDate - Ngày chứng từ
 * @param {string|number} manualPeriod - Giá trị kỳ người dùng tự gõ (nếu có)
 * @returns {string} Kỳ hạch toán dạng "YYYYMM"
 */
function resolvePeriod_(transDate, manualPeriod) {
  if (manualPeriod && manualPeriod.toString().trim() !== '') {
    return manualPeriod.toString().trim();
  }
  
  if (transDate instanceof Date && !isNaN(transDate.getTime())) {
    const year = transDate.getFullYear();
    const month = String(transDate.getMonth() + 1).padStart(2, '0');
    return `${year}${month}`;
  }
  
  return '';
}

/**
 * Tạo chuỗi Composite PK từ trans_no, line_no, item_code
 * @param {string} transNo - Số chứng từ
 * @param {number|string} lineNo - STT dòng
 * @param {string} itemCode - Mã SKU
 * @returns {string} Chuỗi PK duy nhất
 */
function generateCompositeKey_(transNo, lineNo, itemCode) {
  const cleanTransNo = transNo ? transNo.toString().trim() : '';
  const cleanLine = lineNo ? lineNo.toString().trim() : '0';
  const cleanCode = itemCode ? itemCode.toString().trim() : '';
  
  return `${cleanTransNo}_${cleanLine}_${cleanCode}`;
}
