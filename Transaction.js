/**
 * ============================================================================
 * MODULE: Transaction.gs
 * MỤC ĐÍCH: Tra cứu và điền mã SKU (item_code) từ sheet MAPPING quay trở lại
 * các sheet giao dịch khai báo trong SCHEMA.
 * ============================================================================
 */

/**
 * Entry Point (Menu UI): Tra cứu và điền mã SKU (item_code) từ sheet MAPPING
 * quay trở lại tất cả các sheet giao dịch có khai báo col_key 'raw_name' và 'item_code'.
 */
function runAssignItemCodeToTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. GỌI CORE API LẤY LƯỢC ĐỒ SCHEMA
  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi", "Không thể đọc cấu trúc từ sheet 'SCHEMA'!", ui.ButtonSet.OK);
    return;
  }

  // Lấy sheet MAPPING vật lý qua schema_name
  const mappingSheetName = schemaGetSheetName(schemaMap, "MAPPING");
  const mappingSheet = mappingSheetName ? ss.getSheetByName(mappingSheetName) : null;
  if (!mappingSheet) {
    ui.alert("⚠️ Thông báo", "Không tìm thấy sheet tương ứng với SCHEMA 'MAPPING'!", ui.ButtonSet.OK);
    return;
  }

  // Tra cứu vị trí cột trong MAPPING bằng schema_name
  const idxMappingRaw = schemaGetColIndex(schemaMap, "MAPPING", "raw_name");
  const idxMappingSource = schemaGetColIndex(schemaMap, "MAPPING", "source");
  const idxMappingCode = schemaGetColIndex(schemaMap, "MAPPING", "item_code");

  if (idxMappingRaw === -1 || idxMappingSource === -1 || idxMappingCode === -1) {
    ui.alert("⚠️ Lỗi SCHEMA", "MAPPING trong SCHEMA chưa khai báo đủ col_key: 'raw_name', 'source', 'item_code'!", ui.ButtonSet.OK);
    return;
  }

  // 2. DỰNG MAP TRA CỨU ĐỘC LẬP (Dictionary Lookup)
  // Key: "raw_name_lowercase|schema_name_lowercase" -> Value: "item_code"
  const mappingData = mappingSheet.getDataRange().getValues();
  const codeLookupMap = new Map();

  for (let i = 1; i < mappingData.length; i++) {
    const rawVal = mappingData[i][idxMappingRaw] ? mappingData[i][idxMappingRaw].toString().trim().toLowerCase() : "";
    const sourceVal = mappingData[i][idxMappingSource] ? mappingData[i][idxMappingSource].toString().trim().toLowerCase() : "";
    const codeVal = mappingData[i][idxMappingCode] ? mappingData[i][idxMappingCode].toString().trim() : "";

    if (rawVal && sourceVal && codeVal) {
      codeLookupMap.set(`${rawVal}|${sourceVal}`, codeVal);
    }
  }

  if (codeLookupMap.size === 0) {
    ui.alert("ℹ️ Thông báo", "Sheet 'MAPPING' chưa có dữ liệu Mã SKU (item_code) hợp lệ để điền!", ui.ButtonSet.OK);
    return;
  }

  let totalUpdatedSheets = 0;
  let totalUpdatedRows = 0;

  // 3. DUYỆT QUA TẤT CẢ CÁC SCHEMA CÓ CẢ 'raw_name' VÀ 'item_code'
  Object.keys(schemaMap).forEach(schemaName => {
    if (schemaName === "MAPPING" || schemaName === "SCHEMA") return;

    // Kiểm tra xem schema có đủ 2 cột raw_name và item_code không
    const idxSourceRaw = schemaGetColIndex(schemaMap, schemaName, "raw_name");
    const idxSourceCode = schemaGetColIndex(schemaMap, schemaName, "item_code");

    if (idxSourceRaw === -1 || idxSourceCode === -1) return; // Bỏ qua nếu thiếu 1 trong 2 cột

    // Lấy tên sheet vật lý từ schema_name
    const targetSheetName = schemaGetSheetName(schemaMap, schemaName);
    const targetSheet = targetSheetName ? ss.getSheetByName(targetSheetName) : null;
    if (!targetSheet) return;

    const sourceData = targetSheet.getDataRange().getValues();
    if (sourceData.length <= 1) return;

    let sheetUpdatedCount = 0;

    for (let i = 1; i < sourceData.length; i++) {
      const rawVal = sourceData[i][idxSourceRaw] ? sourceData[i][idxSourceRaw].toString().trim().toLowerCase() : "";
      if (!rawVal) continue;

      // Tra cứu theo key: raw_name + schemaName
      const lookupKey = `${rawVal}|${schemaName.toLowerCase()}`;

      // Nếu tìm thấy mã SKU trong Map lookup
      if (codeLookupMap.has(lookupKey)) {
        const matchedCode = codeLookupMap.get(lookupKey);
        const currentCode = sourceData[i][idxSourceCode] ? sourceData[i][idxSourceCode].toString().trim() : "";

        // Chỉ cập nhật nếu mã khác hoặc chưa có
        if (currentCode !== matchedCode) {
          sourceData[i][idxSourceCode] = matchedCode;
          sheetUpdatedCount++;
        }
      }
    }

    // 4. CẬP NHẬT GHI HÀNG LOẠT VÀO SHEET (CHỈ CẬP NHẬT CỘT item_code)
    if (sheetUpdatedCount > 0) {
      const codeColumnValues = sourceData.slice(1).map(row => [row[idxSourceCode]]);
      const colCode1Based = idxSourceCode + 1; // Đổi 0-based index sang 1-based index cho Range

      targetSheet.getRange(2, colCode1Based, codeColumnValues.length, 1).setValues(codeColumnValues);

      totalUpdatedSheets++;
      totalUpdatedRows += sheetUpdatedCount;
    }
  });

  // 5. THÔNG BÁO KẾT QUẢ
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
  // 1. Nếu người dùng đã gõ chủ động (VD: "202603" hoặc 202603)
  if (manualPeriod && manualPeriod.toString().trim() !== '') {
    return manualPeriod.toString().trim();
  }
  
  // 2. Nếu để trống, tự động lấy YYYYMM từ trans_date
  if (transDate instanceof Date && !isNaN(transDate.getTime())) {
    const year = transDate.getFullYear();
    const month = String(transDate.getMonth() + 1).padStart(2, '0');
    return `${year}${month}`;
  }
  
  return '';
}

/**
 * Tao chuoi Composite PK tu trans_no, line_no, item_code
 * @param {string} transNo - So chung tu
 * @param {number|string} lineNo - STT dong
 * @param {string} itemCode - Ma SKU
 * @returns {string} Chuoi PK duy nhat (VD: "260305-02_1_GASPC4501")
 */
function generateCompositeKey_(transNo, lineNo, itemCode) {
  const cleanTransNo = transNo ? transNo.toString().trim() : '';
  const cleanLine = lineNo ? lineNo.toString().trim() : '0';
  const cleanCode = itemCode ? itemCode.toString().trim() : '';
  
  return `${cleanTransNo}_${cleanLine}_${cleanCode}`;
}