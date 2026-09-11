/**
 * ============================================================================
 * MODULE: MappingSync.gs
 * MỤC ĐÍCH: Xử lý quét dữ liệu raw_name, sinh mã SKU/nhãn unmapped và đồng bộ.
 * ============================================================================
 */

/**
 * Entry Point (Menu UI): Quét dữ liệu thô (raw_name) từ tất cả các sheet giao dịch
 * có khai báo col_key 'raw_name' trong SCHEMA và đẩy dữ liệu mới vào bảng MAPPING.
 */
function runSyncMappingFromTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi", "Không thể đọc cấu trúc từ sheet 'SCHEMA'!", ui.ButtonSet.OK);
    return;
  }

  // 1. Lấy Sheet MAPPING vật lý qua schema_name
  const mappingSheetName = schemaGetSheetName(schemaMap, "MAPPING");
  const mappingSheet = mappingSheetName ? ss.getSheetByName(mappingSheetName) : null;
  if (!mappingSheet) {
    ui.alert("⚠️ Thông báo", "Không tìm thấy sheet tương ứng với SCHEMA 'MAPPING'!", ui.ButtonSet.OK);
    return;
  }

  // 2. Tra cứu cột bằng schema_name "MAPPING"
  const idxMappingRaw = schemaGetColIndex(schemaMap, "MAPPING", "raw_name");
  const idxMappingSource = schemaGetColIndex(schemaMap, "MAPPING", "source");

  if (idxMappingRaw === -1 || idxMappingSource === -1) {
    ui.alert("⚠️ Lỗi SCHEMA", "MAPPING trong SCHEMA chưa khai báo col_key 'raw_name' hoặc 'source'!", ui.ButtonSet.OK);
    return;
  }

  const maxColCount = schemaMap["MAPPING"].columns.length;
  const mappingData = mappingSheet.getDataRange().getValues();
  const existingKeys = new Set();

  for (let i = 1; i < mappingData.length; i++) {
    const rawVal = mappingData[i][idxMappingRaw] ? mappingData[i][idxMappingRaw].toString().trim() : "";
    const sourceVal = mappingData[i][idxMappingSource] ? mappingData[i][idxMappingSource].toString().trim() : "";
    if (rawVal && sourceVal) {
      existingKeys.add(`${rawVal.toLowerCase()}|${sourceVal.toLowerCase()}`);
    }
  }

  const newRowsToAppend = [];

  // 3. Đã sửa: Lặp qua từng schemaName (Key của schemaMap)
  Object.keys(schemaMap).forEach(schemaName => {
    // Bỏ qua chính MAPPING và SCHEMA
    if (schemaName === "MAPPING" || schemaName === "SCHEMA") return;

    // Tìm index cột raw_name theo schemaName
    const idxSourceRaw = schemaGetColIndex(schemaMap, schemaName, "raw_name");
    if (idxSourceRaw === -1) return;

    // Lấy tên sheet vật lý thông qua schemaGetSheetName
    const targetSheetName = schemaGetSheetName(schemaMap, schemaName);
    const targetSheet = targetSheetName ? ss.getSheetByName(targetSheetName) : null;
    if (!targetSheet) return;

    const sourceData = targetSheet.getDataRange().getValues();
    if (sourceData.length <= 1) return;

    const localSeen = new Set();

    for (let i = 1; i < sourceData.length; i++) {
      const rawVal = sourceData[i][idxSourceRaw];
      if (!rawVal) continue;

      const rawName = rawVal.toString().trim();
      if (!rawName) continue;

      // Đã sửa: Composite Key dùng schemaName cố định làm Nguồn
      const compositeKey = `${rawName.toLowerCase()}|${schemaName.toLowerCase()}`;

      if (!existingKeys.has(compositeKey) && !localSeen.has(compositeKey)) {
        localSeen.add(compositeKey);
        existingKeys.add(compositeKey);

        const newRow = new Array(maxColCount).fill("");
        newRow[idxMappingRaw] = rawName;
        newRow[idxMappingSource] = schemaName; // Đã sửa: Lưu schemaName cố định thay vì targetSheetName

        newRowsToAppend.push(newRow);
      }
    }
  });

  if (newRowsToAppend.length > 0) {
    const lastRow = mappingSheet.getLastRow();
    mappingSheet.getRange(lastRow + 1, 1, newRowsToAppend.length, maxColCount).setValues(newRowsToAppend);

    ui.alert(
      "✅ Hoàn thành",
      `Đã quét từ các sheet giao dịch và thêm mới ${newRowsToAppend.length} item(s) vào sheet 'MAPPING'.`,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      "ℹ️ Thông báo",
      "Tất cả tên thô (raw_name) từ các sheet giao dịch đều đã tồn tại trong sheet 'MAPPING'.",
      ui.ButtonSet.OK
    );
  }
}
// ================================================================================
// Entry End
// ================================================================================


/**
 * Entry Point (Menu UI): Tự động sinh mã SKU cho các dòng trong MAPPING.
 * CẶP KHÓA CHÍNH (PK): raw_name + source
 * 
 * QUY TẮC BẮT BỘC:
 * - Không gán nhãn vào cột item_name.
 * - Nếu item_name để trống -> item_code tự động gán nhãn "🔴 unmapped".
 * - Nếu item_name được nhập -> Tự động đổi item_code thành mã SKU chuẩn từ item_name.
 * - Cú pháp mã SKU: PREFIX_[ITEM_NAME KHÔNG DẤU]01
 *   + PREFIX: 'MON_' (nguồn SALES), 'NVL_' (nguồn TRANSACTION/STOCKTAKE)
 *   + ITEM_NAME: Chữ hoa, viết liền, không dấu tiếng Việt
 *   + HẬU TỐ: Cố định '01'
 */
function runAutoGenerateSKUForMapping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. LẤY LƯỢC ĐỒ SCHEMA
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
  const idxMappingRaw    = schemaGetColIndex(schemaMap, "MAPPING", "raw_name");
  const idxMappingSource = schemaGetColIndex(schemaMap, "MAPPING", "source");
  const idxMappingCode   = schemaGetColIndex(schemaMap, "MAPPING", "item_code");
  const idxMappingName   = schemaGetColIndex(schemaMap, "MAPPING", "item_name");

  if (idxMappingRaw === -1 || idxMappingCode === -1 || idxMappingSource === -1) {
    ui.alert("⚠️ Lỗi SCHEMA", "MAPPING trong SCHEMA chưa khai báo đủ col_key 'raw_name', 'source' hoặc 'item_code'!", ui.ButtonSet.OK);
    return;
  }

  const mappingData = mappingSheet.getDataRange().getValues();
  if (mappingData.length <= 1) {
    ui.alert("ℹ️ Thông báo", "Sheet MAPPING không có dữ liệu để xử lý!", ui.ButtonSet.OK);
    return;
  }

  // 2. TẬP HỢP CÁC CẶP KHÓA CHÍNH PK (raw_name + source) ĐỂ BẢO TỒN TỰ ĐỒNG BỘ
  const primaryKeyMap = new Map();
  for (let i = 1; i < mappingData.length; i++) {
    const rawVal = mappingData[i][idxMappingRaw] ? mappingData[i][idxMappingRaw].toString().trim().toLowerCase() : "";
    const sourceVal = mappingData[i][idxMappingSource] ? mappingData[i][idxMappingSource].toString().trim().toLowerCase() : "";
    const codeVal = mappingData[i][idxMappingCode] ? mappingData[i][idxMappingCode].toString().trim() : "";

    if (rawVal && sourceVal) {
      const pk = `${rawVal}|${sourceVal}`;
      primaryKeyMap.set(pk, codeVal);
    }
  }

  // 3. THỰC HIỆN XỬ LÝ GÁN NHÃN VÀ SINH MÃ SKU
  let updatedCount = 0;

  for (let i = 1; i < mappingData.length; i++) {
    const rawVal = mappingData[i][idxMappingRaw] ? mappingData[i][idxMappingRaw].toString().trim() : "";
    const sourceVal = mappingData[i][idxMappingSource] ? mappingData[i][idxMappingSource].toString().trim().toUpperCase() : "";
    const currentCode = mappingData[i][idxMappingCode] ? mappingData[i][idxMappingCode].toString().trim() : "";

    if (!rawVal || !sourceVal) continue;

    const pkKey = `${rawVal.toLowerCase()}|${sourceVal.toLowerCase()}`;

    // Lấy giá trị item_name (giữ nguyên nếu trống)
    let nameVal = "";
    if (idxMappingName !== -1) {
      nameVal = mappingData[i][idxMappingName] ? mappingData[i][idxMappingName].toString().trim() : "";
    }

    // TH 1: item_name VẪN ĐỂ TRỐNG -> Gán nhãn "🔴 unmapped" vào item_code
    if (!nameVal) {
      if (currentCode !== "🔴 unmapped") {
        mappingData[i][idxMappingCode] = "🔴 unmapped";
        primaryKeyMap.set(pkKey, "🔴 unmapped");
        updatedCount++;
      }
      continue;
    }

    // TH 2: CÓ ITEM_NAME CHUẨN -> Kiểm tra sinh/cập nhật mã SKU từ item_name
    const isCodeEmptyOrUnmapped = !currentCode || 
                                  currentCode === "🔴 unmapped" || 
                                  currentCode.toUpperCase() === "UNMAPPED";

    if (isCodeEmptyOrUnmapped) {
      // Xác định Prefix theo nguồn
      let prefix = "NVL_";
      if (sourceVal === "SALES") {
        prefix = "MON_";
      } else if (sourceVal === "TRANSACTION" || sourceVal === "STOCKTAKE") {
        prefix = "NVL_";
      } else {
        prefix = "SKU_";
      }

      // Tạo mã DỰA TRÊN ITEM_NAME (Chữ hoa, viết liền, không dấu)
      const cleanNameCode = removeVietnameseTones_(nameVal)
        .toUpperCase()
        .replace(/[^A-Z0-9]/g, ""); // Chỉ giữ chữ cái và số

      const generatedCode = `${prefix}${cleanNameCode}01`;

      if (currentCode !== generatedCode) {
        mappingData[i][idxMappingCode] = generatedCode;
        primaryKeyMap.set(pkKey, generatedCode);
        updatedCount++;
      }
    }
  }

  // 4. GHI DỮ LIỆU HÀNG LOẠT TRỞ LẠI SHEET MAPPING
  if (updatedCount > 0) {
    const maxColCount = schemaMap["MAPPING"].columns.length;
    mappingSheet.getRange(1, 1, mappingData.length, maxColCount).setValues(mappingData);

    ui.alert(
      "✅ Hoàn thành",
      `Đã cập nhật nhãn '🔴 unmapped' / sinh mã SKU cho ${updatedCount} dòng theo cặp khóa PK trong sheet MAPPING.`,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      "ℹ️ Thông báo",
      "Tất cả các dòng theo cặp khóa trong MAPPING đã có nhãn hoặc mã SKU hợp lệ.",
      ui.ButtonSet.OK
    );
  }
}
// END FUNC
