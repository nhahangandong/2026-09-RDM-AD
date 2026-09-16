// /**
//  * ============================================================================
//  * MODULE: MappingSync.gs
//  * MỤC ĐÍCH: Xử lý quét dữ liệu raw_name, sinh mã SKU/nhãn unmapped và đồng bộ.
//  * ============================================================================
//  */

//  /**
//  * Entry Point (Menu UI): Quét dữ liệu thô (raw_name) từ tất cả các sheet giao dịch
//  * có khai báo col_key 'raw_name' trong SCHEMA và đẩy dữ liệu mới vào bảng MAPPING.
//  */
// function runSyncMappingFromTransactions() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const ui = SpreadsheetApp.getUi();

//   const schemaMap = schemaGetMap();
//   if (!schemaMap) {
//     ui.alert("⚠️ Lỗi", "Không thể đọc cấu trúc từ sheet 'SCHEMA'!", ui.ButtonSet.OK);
//     return;
//   }

//   // 1. Lấy Sheet MAPPING vật lý qua schema_name
//   const mappingSheetName = schemaGetSheetName(schemaMap, "MAPPING");
//   const mappingSheet = mappingSheetName ? ss.getSheetByName(mappingSheetName) : null;
//   if (!mappingSheet) {
//     ui.alert("⚠️ Thông báo", "Không tìm thấy sheet tương ứng với SCHEMA 'MAPPING'!", ui.ButtonSet.OK);
//     return;
//   }

//   // 2. Tra cứu cột bằng schema_name "MAPPING"
//   const idxMappingRaw    = schemaGetColIndex(schemaMap, "MAPPING", "raw_name");
//   const idxMappingSource = schemaGetColIndex(schemaMap, "MAPPING", "source");
//   const idxMappingName   = schemaGetColIndex(schemaMap, "MAPPING", "item_name");
//   const idxMappingCode   = schemaGetColIndex(schemaMap, "MAPPING", "item_code");

//   if (idxMappingRaw === -1 || idxMappingSource === -1) {
//     ui.alert("⚠️ Lỗi SCHEMA", "MAPPING trong SCHEMA chưa khai báo col_key 'raw_name' hoặc 'source'!", ui.ButtonSet.OK);
//     return;
//   }

//   const maxColCount = schemaMap["MAPPING"].columns.length;
//   const mappingData = mappingSheet.getDataRange().getValues();

//   // =========================================================================
//   // [BỔ SUNG] 2.1. LỌC VÀ LÀM SẠCH DỮ LIỆU CŨ TRONG MAPPING THEO SOURCE_GROUP
//   // =========================================================================
//   const uniqueMappingRowsMap = new Map(); // Key: raw_name|source_group -> Row Data
//   const existingKeys = new Set();
//   let duplicateOldRowsCount = 0;

//   for (let i = 1; i < mappingData.length; i++) {
//     const row = mappingData[i];
//     const rawVal = row[idxMappingRaw] ? row[idxMappingRaw].toString().trim() : "";
//     const sourceVal = row[idxMappingSource] ? row[idxMappingSource].toString().trim() : "";

//     if (!rawVal || !sourceVal) continue;

//     const sourceGroup = getSourceGroup_(sourceVal);
//     const pkGroupKey = `${rawVal.toLowerCase()}|${sourceGroup.toLowerCase()}`;

//     const itemName = idxMappingName !== -1 && row[idxMappingName] ? row[idxMappingName].toString().trim() : "";
//     const itemCode = idxMappingCode !== -1 && row[idxMappingCode] ? row[idxMappingCode].toString().trim() : "";

//     if (!uniqueMappingRowsMap.has(pkGroupKey)) {
//       uniqueMappingRowsMap.set(pkGroupKey, row);
//     } else {
//       // ĐÃ TỒN TẠI DÒNG TRÙNG CŨ THEO SOURCE_GROUP
//       duplicateOldRowsCount++;
//       const existingRow = uniqueMappingRowsMap.get(pkGroupKey);
//       const existingName = idxMappingName !== -1 && existingRow[idxMappingName] ? existingRow[idxMappingName].toString().trim() : "";

//       // ƯU TIÊN GIỮ DÒNG ĐÃ MAP ITEM_NAME/MÃ SKU THẬT (Khác trống và khác unmapped)
//       const hasRealMapping = itemName && !itemName.toLowerCase().includes("unmapped");
//       const existingHasReal = existingName && !existingName.toLowerCase().includes("unmapped");

//       if (hasRealMapping && !existingHasReal) {
//         uniqueMappingRowsMap.set(pkGroupKey, row); // Ghi đè lấy dòng đã được map tay chuẩn
//       }
//     }

//     existingKeys.add(pkGroupKey);
//   }

//   // Nếu phát hiện dữ liệu cũ có trùng lặp theo source_group -> Ghi đè làm sạch sheet MAPPING ngay
//   if (duplicateOldRowsCount > 0) {
//     const cleanedRows = Array.from(uniqueMappingRowsMap.values());
//     mappingSheet.getRange(2, 1, mappingSheet.getLastRow() - 1, maxColCount).clearContent();
//     if (cleanedRows.length > 0) {
//       mappingSheet.getRange(2, 1, cleanedRows.length, maxColCount).setValues(cleanedRows);
//     }
//   }

//   // =========================================================================
//   // 3. QUÉT THÊM ITEM MỚI TỪ CÁC SHEET GIAO DỊCH DỰA TRÊN EXISTINGKEYS ĐÃ LỌC
//   // =========================================================================
//   const newRowsToAppend = [];

//   Object.keys(schemaMap).forEach(schemaName => {
//     if (schemaName === "MAPPING" || schemaName === "SCHEMA") return;

//     const idxSourceRaw = schemaGetColIndex(schemaMap, schemaName, "raw_name");
//     if (idxSourceRaw === -1) return;

//     const targetSheetName = schemaGetSheetName(schemaMap, schemaName);
//     const targetSheet = targetSheetName ? ss.getSheetByName(targetSheetName) : null;
//     if (!targetSheet) return;

//     const sourceData = targetSheet.getDataRange().getValues();
//     if (sourceData.length <= 1) return;

//     const localSeen = new Set();
//     const sourceGroup = getSourceGroup_(schemaName);

//     for (let i = 1; i < sourceData.length; i++) {
//       const rawVal = sourceData[i][idxSourceRaw];
//       if (!rawVal) continue;

//       const rawName = rawVal.toString().trim();
//       if (!rawName) continue;

//       const compositeKey = `${rawName.toLowerCase()}|${sourceGroup.toLowerCase()}`;

//       if (!existingKeys.has(compositeKey) && !localSeen.has(compositeKey)) {
//         localSeen.add(compositeKey);
//         existingKeys.add(compositeKey);

//         const newRow = new Array(maxColCount).fill("");
//         newRow[idxMappingRaw] = rawName;
//         newRow[idxMappingSource] = schemaName;

//         newRowsToAppend.push(newRow);
//       }
//     }
//   });

//   // Ghi thêm dòng mới
//   if (newRowsToAppend.length > 0) {
//     const lastRow = mappingSheet.getLastRow();
//     mappingSheet.getRange(lastRow + 1, 1, newRowsToAppend.length, maxColCount).setValues(newRowsToAppend);
//   }

//   // Thông báo kết quả rõ ràng
//   let msg = [];
//   if (duplicateOldRowsCount > 0) {
//     msg.push(`🧹 Đã dọn dẹp ${duplicateOldRowsCount} dòng dữ liệu lặp cũ trong MAPPING.`);
//   }
//   if (newRowsToAppend.length > 0) {
//     msg.push(`✅ Đã thêm mới ${newRowsToAppend.length} item(s) mới vào 'MAPPING'.`);
//   }
//   if (msg.length === 0) {
//     msg.push("ℹ️ Dữ liệu MAPPING đã chuẩn hóa hoàn toàn và không có item mới.");
//   }

//   ui.alert("Kết quả Đồng bộ", msg.join("\n"), ui.ButtonSet.OK);
// }


// /**
//  * Private Helper: Chuyển đổi schemaName thành Nhóm nghiệp vụ (source_group)
//  * dùng riêng cho việc tạo khóa lọc trùng PK (Composite Key).
//  */
// function getSourceGroup_(schemaName) {
//   if (!schemaName) return "UNKNOWN";
//   const nameUpper = schemaName.toString().trim().toUpperCase();
  
//   if (nameUpper === "TRANSACTION" || nameUpper === "STOCKTAKE" || nameUpper === "STOCK_TAKE") {
//     return "INVENTORY";
//   }
//   if (nameUpper === "SALES") {
//     return "SALES";
//   }
//   return nameUpper;
// }
// //===========================================================================================================

// /**
//  * Entry Point (Menu UI): Tự động sinh mã SKU cho các dòng trong MAPPING.
//  * CẶP KHÓA CHÍNH (PK): raw_name + source
//  */
// function runAutoGenerateSKUForMapping() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const ui = SpreadsheetApp.getUi();

//   // 1. LẤY LƯỢC ĐỒ SCHEMA
//   const schemaMap = schemaGetMap();
//   if (!schemaMap) {
//     ui.alert("⚠️ Lỗi", "Không thể đọc cấu trúc từ sheet 'SCHEMA'!", ui.ButtonSet.OK);
//     return;
//   }

//   // Lấy sheet MAPPING vật lý qua schema_name
//   const mappingSheetName = schemaGetSheetName(schemaMap, "MAPPING");
//   const mappingSheet = mappingSheetName ? ss.getSheetByName(mappingSheetName) : null;
//   if (!mappingSheet) {
//     ui.alert("⚠️ Thông báo", "Không tìm thấy sheet tương ứng với SCHEMA 'MAPPING'!", ui.ButtonSet.OK);
//     return;
//   }

//   // Tra cứu vị trí cột trong MAPPING bằng schema_name
//   const idxMappingRaw    = schemaGetColIndex(schemaMap, "MAPPING", "raw_name");
//   const idxMappingSource = schemaGetColIndex(schemaMap, "MAPPING", "source");
//   const idxMappingCode   = schemaGetColIndex(schemaMap, "MAPPING", "item_code");
//   const idxMappingName   = schemaGetColIndex(schemaMap, "MAPPING", "item_name");

//   if (idxMappingRaw === -1 || idxMappingCode === -1 || idxMappingSource === -1) {
//     ui.alert("⚠️ Lỗi SCHEMA", "MAPPING trong SCHEMA chưa khai báo đủ col_key 'raw_name', 'source' hoặc 'item_code'!", ui.ButtonSet.OK);
//     return;
//   }

//   const mappingData = mappingSheet.getDataRange().getValues();
//   if (mappingData.length <= 1) {
//     ui.alert("ℹ️ Thông báo", "Sheet MAPPING không có dữ liệu để xử lý!", ui.ButtonSet.OK);
//     return;
//   }

//   // 2. TẬP HỢP CÁC CẶP KHÓA CHÍNH PK (raw_name + source_group)
//   const primaryKeyMap = new Map();
//   for (let i = 1; i < mappingData.length; i++) {
//     const rawVal = mappingData[i][idxMappingRaw] ? mappingData[i][idxMappingRaw].toString().trim().toLowerCase() : "";
//     const sourceVal = mappingData[i][idxMappingSource] ? mappingData[i][idxMappingSource].toString().trim() : "";
//     const codeVal = mappingData[i][idxMappingCode] ? mappingData[i][idxMappingCode].toString().trim() : "";

//     if (rawVal && sourceVal) {
//       const sourceGroup = getSourceGroup_(sourceVal);
//       const pk = `${rawVal}|${sourceGroup.toLowerCase()}`;
//       primaryKeyMap.set(pk, codeVal);
//     }
//   }

//   // 3. THỰC HIỆN XỬ LÝ GÁN NHÃN VÀ SINH MÃ SKU
//   let updatedCount = 0;

//   for (let i = 1; i < mappingData.length; i++) {
//     const rawVal = mappingData[i][idxMappingRaw] ? mappingData[i][idxMappingRaw].toString().trim() : "";
//     const sourceVal = mappingData[i][idxMappingSource] ? mappingData[i][idxMappingSource].toString().trim().toUpperCase() : "";
//     const currentCode = mappingData[i][idxMappingCode] ? mappingData[i][idxMappingCode].toString().trim() : "";

//     if (!rawVal || !sourceVal) continue;

//     const sourceGroup = getSourceGroup_(sourceVal);
//     const pkKey = `${rawVal.toLowerCase()}|${sourceGroup.toLowerCase()}`;

//     // Lấy giá trị item_name (giữ nguyên nếu trống)
//     let nameVal = "";
//     if (idxMappingName !== -1) {
//       nameVal = mappingData[i][idxMappingName] ? mappingData[i][idxMappingName].toString().trim() : "";
//     }

//     // TH 1: item_name VẪN ĐỂ TRỐNG -> Gán nhãn "🔴 unmapped" vào item_code
//     if (!nameVal) {
//       if (currentCode !== "🔴 unmapped") {
//         mappingData[i][idxMappingCode] = "🔴 unmapped";
//         primaryKeyMap.set(pkKey, "🔴 unmapped");
//         updatedCount++;
//       }
//       continue;
//     }

//     // TH 2: CÓ ITEM_NAME CHUẨN -> Kiểm tra sinh/cập nhật mã SKU từ item_name
//     const isCodeEmptyOrUnmapped = !currentCode || 
//                                   currentCode === "🔴 unmapped" || 
//                                   currentCode.toUpperCase() === "UNMAPPED";

//     if (isCodeEmptyOrUnmapped) {
//       // Xác định Prefix theo nguồn
//       let prefix = "NVL_";
//       if (sourceGroup === "SALES") {
//         prefix = "MON_";
//       } else if (sourceGroup === "INVENTORY") {
//         prefix = "NVL_";
//       } else {
//         prefix = "SKU_";
//       }

//       // Tạo mã DỰA TRÊN ITEM_NAME (Chữ hoa, viết liền, không dấu)
//       const cleanNameCode = removeVietnameseTones_(nameVal)
//         .toUpperCase()
//         .replace(/[^A-Z0-9]/g, ""); // Chỉ giữ chữ cái và số

//       const generatedCode = `${prefix}${cleanNameCode}01`;

//       if (currentCode !== generatedCode) {
//         mappingData[i][idxMappingCode] = generatedCode;
//         primaryKeyMap.set(pkKey, generatedCode);
//         updatedCount++;
//       }
//     }
//   }

//   // 4. GHI DỮ LIỆU HÀNG LOẠT TRỞ LẠI SHEET MAPPING
//   if (updatedCount > 0) {
//     const maxColCount = schemaMap["MAPPING"].columns.length;
//     mappingSheet.getRange(1, 1, mappingData.length, maxColCount).setValues(mappingData);

//     ui.alert(
//       "✅ Hoàn thành",
//       `Đã cập nhật nhãn '🔴 unmapped' / sinh mã SKU cho ${updatedCount} dòng trong sheet MAPPING.`,
//       ui.ButtonSet.OK
//     );
//   } else {
//     ui.alert(
//       "ℹ️ Thông báo",
//       "Tất cả các dòng trong MAPPING đã có nhãn hoặc mã SKU hợp lệ.",
//       ui.ButtonSet.OK
//     );
//   }
// }
// // ============= END FUNCTION ==================//


  
// /**
//  * Hàm tự động gợi ý / điền hàng loạt item_name cho các dòng chưa map trong MAPPING.
//  * Dữ liệu quy tắc đọc ĐỘNG từ sheet AUTO_MAP_RULE via SCHEMA.
//  */
// function runAutoSuggestMappingNames() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const ui = SpreadsheetApp.getUi();

//   const schemaMap = schemaGetMap();
//   if (!schemaMap) {
//     ui.alert("⚠️ Lỗi", "Không thể đọc cấu trúc từ sheet 'SCHEMA'!", ui.ButtonSet.OK);
//     return;
//   }

//   // 1. LẤY DỮ LIỆU BẢNG MAPPING
//   const mappingSheetName = schemaGetSheetName(schemaMap, "MAPPING");
//   const mappingSheet = mappingSheetName ? ss.getSheetByName(mappingSheetName) : null;
//   if (!mappingSheet) {
//     ui.alert("⚠️ Thông báo", "Không tìm thấy sheet tương ứng với SCHEMA 'MAPPING'!", ui.ButtonSet.OK);
//     return;
//   }

//   const idxMappingRaw    = schemaGetColIndex(schemaMap, "MAPPING", "raw_name");
//   const idxMappingSource = schemaGetColIndex(schemaMap, "MAPPING", "source");
//   const idxMappingName   = schemaGetColIndex(schemaMap, "MAPPING", "item_name");

//   if (idxMappingRaw === -1 || idxMappingName === -1 || idxMappingSource === -1) {
//     ui.alert("⚠️ Lỗi SCHEMA", "MAPPING chưa khai báo đủ col_key 'raw_name', 'source' hoặc 'item_name'!", ui.ButtonSet.OK);
//     return;
//   }

//   const mappingData = mappingSheet.getDataRange().getValues();
//   if (mappingData.length <= 1) return;

//   // 2. ĐỌC ĐỘNG CÁC QUY TẮC TỪ SHEET AUTO_MAP_RULE VIA SCHEMA
//   const autoRuleSheetName = schemaGetSheetName(schemaMap, "AUTO_MAP_RULE");
//   const autoRuleSheet     = autoRuleSheetName ? ss.getSheetByName(autoRuleSheetName) : null;
//   const rules = [];

//   if (autoRuleSheet) {
//     const idxRuleKeywords = schemaGetColIndex(schemaMap, "AUTO_MAP_RULE", "keywords");
//     const idxRuleGroup    = schemaGetColIndex(schemaMap, "AUTO_MAP_RULE", "source_group");
//     const idxRuleTarget   = schemaGetColIndex(schemaMap, "AUTO_MAP_RULE", "target_item_name");

//     if (idxRuleKeywords !== -1 && idxRuleTarget !== -1) {
//       const ruleData = autoRuleSheet.getDataRange().getValues();
//       for (let i = 1; i < ruleData.length; i++) {
//         const kwStr  = ruleData[i][idxRuleKeywords] ? ruleData[i][idxRuleKeywords].toString().trim() : "";
//         const group  = idxRuleGroup !== -1 && ruleData[i][idxRuleGroup] ? ruleData[i][idxRuleGroup].toString().trim().toUpperCase() : "ALL";
//         const target = ruleData[i][idxRuleTarget] ? ruleData[i][idxRuleTarget].toString().trim() : "";

//         if (kwStr && target) {
//           const kwList = kwStr.split(",").map(k => k.trim().toLowerCase()).filter(k => k);
//           rules.push({ keywords: kwList, sourceGroup: group || "ALL", targetName: target });
//         }
//       }
//     }
//   }

//   // 3. TẠO TỪ ĐIỂN MẪU TỪ CÁC DÒNG ĐÃ MAP TAY TRƯỚC ĐÓ (Machine Learning đơn giản)
//   const learnedNameMap = new Map(); // Key: normalized_raw|source_group -> Value: item_name
//   for (let i = 1; i < mappingData.length; i++) {
//     const rawVal    = mappingData[i][idxMappingRaw] ? mappingData[i][idxMappingRaw].toString().trim() : "";
//     const sourceVal = mappingData[i][idxMappingSource] ? mappingData[i][idxMappingSource].toString().trim() : "";
//     const nameVal   = mappingData[i][idxMappingName] ? mappingData[i][idxMappingName].toString().trim() : "";

//     if (rawVal && sourceVal && nameVal && !nameVal.toLowerCase().includes("unmapped")) {
//       const sourceGroup   = getSourceGroup_(sourceVal);
//       const normalizedKey = `${cleanSpecWords_(rawVal)}|${sourceGroup.toLowerCase()}`;
//       if (!learnedNameMap.has(normalizedKey)) {
//         learnedNameMap.set(normalizedKey, nameVal);
//       }
//     }
//   }

//   // 4. TIẾN HÀNH ĐIỀN TỰ ĐỘNG CÁC DÒNG ĐANG TRỐNG ITEM_NAME
//   let filledCount = 0;

//   for (let i = 1; i < mappingData.length; i++) {
//     const rawVal    = mappingData[i][idxMappingRaw] ? mappingData[i][idxMappingRaw].toString().trim() : "";
//     const sourceVal = mappingData[i][idxMappingSource] ? mappingData[i][idxMappingSource].toString().trim() : "";
//     let currentName = mappingData[i][idxMappingName] ? mappingData[i][idxMappingName].toString().trim() : "";

//     // Bỏ qua nếu dòng rỗng hoặc đã có item_name
//     if (!rawVal || !sourceVal || currentName) continue;

//     const sourceGroup = getSourceGroup_(sourceVal);
//     const rawLower    = rawVal.toLowerCase();

//     // CHIẾN LƯỢC A: MATCH THEO CẤU HÌNH TỪ KHÓA DYNAMIC TRONG AUTO_MAP_RULE
//     let matchedTargetName = null;
//     for (const rule of rules) {
//       if (rule.sourceGroup !== "ALL" && rule.sourceGroup !== sourceGroup) continue;

//       // Kiểm tra nếu raw_name chứa bất kỳ từ khóa nào trong danh sách
//       const isMatch = rule.keywords.some(kw => rawLower.includes(kw));
//       if (isMatch) {
//         matchedTargetName = rule.targetName;
//         break; // Lấy quy tắc khớp đầu tiên
//       }
//     }

//     if (matchedTargetName) {
//       mappingData[i][idxMappingName] = matchedTargetName;
//       filledCount++;
//       continue;
//     }

//     // CHIẾN LƯỢC B: HỌC TỪ CÁC DÒNG ĐÃ ĐƯỢC MAP TAY TƯƠNG TỰ
//     const normalizedKey = `${cleanSpecWords_(rawVal)}|${sourceGroup.toLowerCase()}`;
//     if (learnedNameMap.has(normalizedKey)) {
//       mappingData[i][idxMappingName] = learnedNameMap.get(normalizedKey);
//       filledCount++;
//     }
//   }

//   // 5. CẬP NHẬT LẠI SHEET MAPPING
//   if (filledCount > 0) {
//     const maxColCount = schemaMap["MAPPING"].columns.length;
//     mappingSheet.getRange(1, 1, mappingData.length, maxColCount).setValues(mappingData);
//     ui.alert("✅ Hoàn thành", `Đã tự động gợi ý và điền 'item_name' cho ${filledCount} dòng trong MAPPING!`, ui.ButtonSet.OK);
//   } else {
//     ui.alert("ℹ️ Thông báo", "Không có dòng nào phù hợp với các quy tắc gợi ý tự động.", ui.ButtonSet.OK);
//   }
// }


/**
 * ============================================================================
 * MODULE: MappingSync.gs (Refactored & Structured)
 * MỤC ĐÍCH: Quản lý ánh xạ (Mapping), đồng bộ dữ liệu thô, gợi ý tên chuẩn 
 * và tự động sinh mã SKU cho hệ thống nhập liệu.
 * ============================================================================
 */

// ============================================================================
// I. ENTRY POINTS (GỌI TỪ MENU UI GOOGLE SHEETS)
// ============================================================================

/**
 * 1. Quét dữ liệu thô (raw_name) từ các sheet giao dịch để đưa vào bảng MAPPING.
 */
function runSyncMappingFromTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi", "Không thể đọc cấu trúc từ sheet 'SCHEMA'!", ui.ButtonSet.OK);
    return;
  }

  const mappingSheet = getMappingSheet_(ss, schemaMap);
  if (!mappingSheet) return;

  const idxMappingRaw    = schemaGetColIndex(schemaMap, "MAPPING", "raw_name");
  const idxMappingSource = schemaGetColIndex(schemaMap, "MAPPING", "source");
  const idxMappingName   = schemaGetColIndex(schemaMap, "MAPPING", "item_name");
  const idxMappingCode   = schemaGetColIndex(schemaMap, "MAPPING", "item_code");

  if (idxMappingRaw === -1 || idxMappingSource === -1) {
    ui.alert("⚠️ Lỗi SCHEMA", "MAPPING trong SCHEMA chưa khai báo đủ col_key 'raw_name' hoặc 'source'!", ui.ButtonSet.OK);
    return;
  }

  const maxColCount = schemaMap["MAPPING"].columns.length;
  const mappingData = mappingSheet.getDataRange().getValues();

  // Bước A: Làm sạch dữ liệu cũ trùng lặp trong MAPPING[cite: 2]
  const { uniqueRowsMap, duplicateCount, existingKeys } = cleanExistingMappingData_(mappingData, idxMappingRaw, idxMappingSource, idxMappingName, idxMappingCode);

  if (duplicateCount > 0) {
    const cleanedRows = Array.from(uniqueRowsMap.values());
    mappingSheet.getRange(2, 1, mappingSheet.getLastRow() - 1, maxColCount).clearContent();
    if (cleanedRows.length > 0) {
      mappingSheet.getRange(2, 1, cleanedRows.length, maxColCount).setValues(cleanedRows);
    }
  }

  // Bước B: Quét item mới từ các sheet giao dịch[cite: 2]
  const newRowsToAppend = scanNewTransactions_(ss, schemaMap, existingKeys, maxColCount);

  if (newRowsToAppend.length > 0) {
    const lastRow = mappingSheet.getLastRow();
    mappingSheet.getRange(lastRow + 1, 1, newRowsToAppend.length, maxColCount).setValues(newRowsToAppend);
  }

  // Bước C: Tổng hợp thông báo kết quả[cite: 2]
  showSyncResultMessage_(ui, duplicateCount, newRowsToAppend.length);
}


/**
 * 2. Tự động gợi ý / điền hàng loạt item_name dựa trên AUTO_MAP_RULE và lịch sử[cite: 2].
 */
function runAutoSuggestMappingNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi", "Không thể đọc cấu trúc từ sheet 'SCHEMA'!", ui.ButtonSet.OK);
    return;
  }

  const mappingSheet = getMappingSheet_(ss, schemaMap);
  if (!mappingSheet) return;

  const idxMappingRaw    = schemaGetColIndex(schemaMap, "MAPPING", "raw_name");
  const idxMappingSource = schemaGetColIndex(schemaMap, "MAPPING", "source");
  const idxMappingName   = schemaGetColIndex(schemaMap, "MAPPING", "item_name");

  if (idxMappingRaw === -1 || idxMappingName === -1 || idxMappingSource === -1) {
    ui.alert("⚠️ Lỗi SCHEMA", "MAPPING chưa khai báo đủ cột cơ bản!", ui.ButtonSet.OK);
    return;
  }

  const mappingData = mappingSheet.getDataRange().getValues();
  if (mappingData.length <= 1) return;

  // Tải quy tắc và từ điển học tập
  const rules = loadAutoMapRules_(ss, schemaMap);
  const learnedNameMap = buildLearnedNameDictionary_(mappingData, idxMappingRaw, idxMappingSource, idxMappingName);

  let filledCount = 0;

  for (let i = 1; i < mappingData.length; i++) {
    const rawVal    = mappingData[i][idxMappingRaw] ? mappingData[i][idxMappingRaw].toString().trim() : "";
    const sourceVal = mappingData[i][idxMappingSource] ? mappingData[i][idxMappingSource].toString().trim() : "";
    const currentName = mappingData[i][idxMappingName] ? mappingData[i][idxMappingName].toString().trim() : "";

    if (!rawVal || !sourceVal || currentName) continue;

    const sourceGroup = getSourceGroup_(sourceVal);
    const rawLower    = rawVal.toLowerCase();

    // Chiến lược 1: Khớp quy tắc từ AUTO_MAP_RULE[cite: 2]
    let matchedTargetName = null;
    for (const rule of rules) {
      if (rule.sourceGroup !== "ALL" && rule.sourceGroup !== sourceGroup) continue;
      if (rule.keywords.some(kw => rawLower.includes(kw))) {
        matchedTargetName = rule.targetName;
        break;
      }
    }

    if (matchedTargetName) {
      mappingData[i][idxMappingName] = matchedTargetName;
      filledCount++;
      continue;
    }

    // Chiến lược 2: Học từ các dòng đã map tay trước đó[cite: 2]
    const normalizedKey = `${cleanSpecWords_(rawVal)}|${sourceGroup.toLowerCase()}`;
    if (learnedNameMap.has(normalizedKey)) {
      mappingData[i][idxMappingName] = learnedNameMap.get(normalizedKey);
      filledCount++;
    }
  }

  if (filledCount > 0) {
    const maxColCount = schemaMap["MAPPING"].columns.length;
    mappingSheet.getRange(1, 1, mappingData.length, maxColCount).setValues(mappingData);
    ui.alert("✅ Hoàn thành", `Đã tự động gợi ý và điền 'item_name' cho ${filledCount} dòng trong MAPPING!`, ui.ButtonSet.OK);
  } else {
    ui.alert("ℹ️ Thông báo", "Không có dòng nào phù hợp với các quy tắc gợi ý tự động.", ui.ButtonSet.OK);
  }
}


/**
 * 3. Tự động sinh mã SKU hoặc gán nhãn '🔴 unmapped' cho sheet MAPPING[cite: 2].
 */
function runAutoGenerateSKUForMapping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi", "Không thể đọc cấu trúc từ sheet 'SCHEMA'!", ui.ButtonSet.OK);
    return;
  }

  const mappingSheet = getMappingSheet_(ss, schemaMap);
  if (!mappingSheet) return;

  const idxMappingRaw    = schemaGetColIndex(schemaMap, "MAPPING", "raw_name");
  const idxMappingSource = schemaGetColIndex(schemaMap, "MAPPING", "source");
  const idxMappingCode   = schemaGetColIndex(schemaMap, "MAPPING", "item_code");
  const idxMappingName   = schemaGetColIndex(schemaMap, "MAPPING", "item_name");

  if (idxMappingRaw === -1 || idxMappingCode === -1 || idxMappingSource === -1) {
    ui.alert("⚠️ Lỗi SCHEMA", "MAPPING chưa khai báo đủ cột khóa!", ui.ButtonSet.OK);
    return;
  }

  const mappingData = mappingSheet.getDataRange().getValues();
  if (mappingData.length <= 1) {
    ui.alert("ℹ️ Thông báo", "Sheet MAPPING không có dữ liệu để xử lý!", ui.ButtonSet.OK);
    return;
  }

  let updatedCount = 0;

  for (let i = 1; i < mappingData.length; i++) {
    const rawVal = mappingData[i][idxMappingRaw] ? mappingData[i][idxMappingRaw].toString().trim() : "";
    let sourceVal = mappingData[i][idxMappingSource] ? mappingData[i][idxMappingSource].toString().trim().toUpperCase() : "";
    
    // Fallback an toàn nếu cột nguồn bị trống[cite: 2]
    if (!sourceVal) sourceVal = "INVENTORY";

    const currentCode = mappingData[i][idxMappingCode] ? mappingData[i][idxMappingCode].toString().trim() : "";
    if (!rawVal) continue;

    const sourceGroup = getSourceGroup_(sourceVal);
    let nameVal = idxMappingName !== -1 && mappingData[i][idxMappingName] ? mappingData[i][idxMappingName].toString().trim() : "";

    // TRƯỜNG HỢP 1: Chưa có tên chuẩn -> Gán nhãn "🔴 unmapped"[cite: 2]
    if (!nameVal) {
      if (currentCode !== "🔴 unmapped") {
        mappingData[i][idxMappingCode] = "🔴 unmapped";
        updatedCount++;
      }
      continue;
    }

    // TRƯỜNG HỢP 2: Đã có tên chuẩn -> Kiểm tra và sinh mã SKU[cite: 2]
    const isCodeEmptyOrUnmapped = !currentCode || 
                                  currentCode === "" ||
                                  currentCode === "🔴 unmapped" ||  
                                  currentCode.toUpperCase() === "UNMAPPED";

    if (isCodeEmptyOrUnmapped) {
      const prefix = determineSkuPrefix_(sourceGroup);
      const cleanNameCode = removeVietnameseTones_(nameVal)
        .toUpperCase()
        .replace(/[^A-Z0-9]/g, "");

      const generatedCode = `${prefix}${cleanNameCode}01`;

      if (currentCode !== generatedCode) {
        mappingData[i][idxMappingCode] = generatedCode;
        updatedCount++;
      }
    }
  }

  if (updatedCount > 0) {
    const maxColCount = schemaMap["MAPPING"].columns.length;
    mappingSheet.getRange(1, 1, mappingData.length, maxColCount).setValues(mappingData);
    ui.alert("✅ Hoàn thành", `Đã cập nhật nhãn / sinh mã SKU cho ${updatedCount} dòng trong MAPPING.`, ui.ButtonSet.OK);
  } else {
    ui.alert("ℹ️ Thông báo", "Tất cả các dòng trong MAPPING đã có nhãn hoặc mã SKU hợp lệ.", ui.ButtonSet.OK);
  }
}


// ============================================================================
// II. PRIVATE HELPER FUNCTIONS (HÀM BỔ TRỢ NỘI BỘ)
// ============================================================================

/** Lấy sheet MAPPING vật lý an toàn */
function getMappingSheet_(ss, schemaMap) {
  const mappingSheetName = schemaGetSheetName(schemaMap, "MAPPING");
  const mappingSheet = mappingSheetName ? ss.getSheetByName(mappingSheetName) : null;
  if (!mappingSheet) {
    SpreadsheetApp.getUi().alert("⚠️ Thông báo", "Không tìm thấy sheet tương ứng với SCHEMA 'MAPPING'!", SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }
  return mappingSheet;
}

/** Quy đổi tên schema thành nhóm nghiệp vụ chuẩn[cite: 2] */
function getSourceGroup_(schemaName) {
  if (!schemaName) return "UNKNOWN";
  const nameUpper = schemaName.toString().trim().toUpperCase();
  
  if (nameUpper === "TRANSACTION" || nameUpper === "STOCKTAKE" || nameUpper === "STOCK_TAKE") {
    return "INVENTORY";
  }
  if (nameUpper === "SALES") {
    return "SALES";
  }
  return nameUpper;
}

/** Xác định tiền tố SKU dựa trên nhóm nguồn */
function determineSkuPrefix_(sourceGroup) {
  if (sourceGroup === "SALES") return "MON_";
  if (sourceGroup === "INVENTORY") return "NVL_";
  return "SKU_";
}

/** Làm sạch và loại bỏ các dòng trùng lặp cũ trong MAPPING[cite: 2] */
function cleanExistingMappingData_(mappingData, idxRaw, idxSource, idxName, idxCode) {
  const uniqueRowsMap = new Map();
  const existingKeys = new Set();
  let duplicateCount = 0;

  for (let i = 1; i < mappingData.length; i++) {
    const row = mappingData[i];
    const rawVal = row[idxRaw] ? row[idxRaw].toString().trim() : "";
    const sourceVal = row[idxSource] ? row[idxSource].toString().trim() : "";

    if (!rawVal || !sourceVal) continue;

    const sourceGroup = getSourceGroup_(sourceVal);
    const pkGroupKey = `${rawVal.toLowerCase()}|${sourceGroup.toLowerCase()}`;

    const itemName = idxName !== -1 && row[idxName] ? row[idxName].toString().trim() : "";

    if (!uniqueRowsMap.has(pkGroupKey)) {
      uniqueRowsMap.set(pkGroupKey, row);
    } else {
      duplicateCount++;
      const existingRow = uniqueRowsMap.get(pkGroupKey);
      const existingName = idxName !== -1 && existingRow[idxName] ? existingRow[idxName].toString().trim() : "";

      const hasRealMapping = itemName && !itemName.toLowerCase().includes("unmapped");
      const existingHasReal = existingName && !existingName.toLowerCase().includes("unmapped");

      if (hasRealMapping && !existingHasReal) {
        uniqueRowsMap.set(pkGroupKey, row);
      }
    }
    existingKeys.add(pkGroupKey);
  }

  return { uniqueRowsMap, duplicateCount, existingKeys };
}


/** Quét các item mới từ đúng 3 nguồn giao dịch cho phép: TRANSACTION, STOCKTAKE, SALES */
function scanNewTransactions_(ss, schemaMap, existingKeys, maxColCount) {
  const newRowsToAppend = [];

  // DANH SÁCH TRẮNG (Whitelist) các schema được phép cấp dữ liệu
  const allowedSourceSchemas = ["TRANSACTION", "STOCKTAKE", "SALES"];

  Object.keys(schemaMap).forEach(schemaName => {
    const schemaUpper = schemaName.toUpperCase();
    
    // Chỉ xử lý nếu schema nằm trong danh sách cho phép
    if (!allowedSourceSchemas.includes(schemaUpper)) return;

    const idxSourceRaw = schemaGetColIndex(schemaMap, schemaName, "raw_name");
    if (idxSourceRaw === -1) return;

    const targetSheetName = schemaGetSheetName(schemaMap, schemaName);
    const targetSheet = targetSheetName ? ss.getSheetByName(targetSheetName) : null;
    if (!targetSheet) return;

    const sourceData = targetSheet.getDataRange().getValues();
    if (sourceData.length <= 1) return;

    const localSeen = new Set();
    const sourceGroup = getSourceGroup_(schemaName);

    for (let i = 1; i < sourceData.length; i++) {
      const rawVal = sourceData[i][idxSourceRaw];
      if (!rawVal) continue;

      const rawName = rawVal.toString().trim();
      if (!rawName) continue;

      const compositeKey = `${rawName.toLowerCase()}|${sourceGroup.toLowerCase()}`;

      if (!existingKeys.has(compositeKey) && !localSeen.has(compositeKey)) {
        localSeen.add(compositeKey);
        existingKeys.add(compositeKey);

        const newRow = new Array(maxColCount).fill("");
        
        const idxMapRawCol = schemaGetColIndex(schemaMap, "MAPPING", "raw_name");
        const idxMapSourceCol = schemaGetColIndex(schemaMap, "MAPPING", "source");
        
        if (idxMapRawCol !== -1) newRow[idxMapRawCol] = rawName;
        if (idxMapSourceCol !== -1) newRow[idxMapSourceCol] = schemaName; // Ghi nhận đúng tên nguồn gốc

        newRowsToAppend.push(newRow);
      }
    }
  });

  return newRowsToAppend;
}

/** Tải danh sách quy tắc tự động từ sheet AUTO_MAP_RULE[cite: 2] */
function loadAutoMapRules_(ss, schemaMap) {
  const autoRuleSheetName = schemaGetSheetName(schemaMap, "AUTO_MAP_RULE");
  const autoRuleSheet = autoRuleSheetName ? ss.getSheetByName(autoRuleSheetName) : null;
  const rules = [];

  if (autoRuleSheet) {
    const idxRuleKeywords = schemaGetColIndex(schemaMap, "AUTO_MAP_RULE", "keywords");
    const idxRuleGroup    = schemaGetColIndex(schemaMap, "AUTO_MAP_RULE", "source_group");
    const idxRuleTarget   = schemaGetColIndex(schemaMap, "AUTO_MAP_RULE", "target_item_name");

    if (idxRuleKeywords !== -1 && idxRuleTarget !== -1) {
      const ruleData = autoRuleSheet.getDataRange().getValues();
      for (let i = 1; i < ruleData.length; i++) {
        const kwStr  = ruleData[i][idxRuleKeywords] ? ruleData[i][idxRuleKeywords].toString().trim() : "";
        const group  = idxRuleGroup !== -1 && ruleData[i][idxRuleGroup] ? ruleData[i][idxRuleGroup].toString().trim().toUpperCase() : "ALL";
        const target = ruleData[i][idxRuleTarget] ? ruleData[i][idxRuleTarget].toString().trim() : "";

        if (kwStr && target) {
          const kwList = kwStr.split(",").map(k => k.trim().toLowerCase()).filter(k => k);
          rules.push({ keywords: kwList, sourceGroup: group || "ALL", targetName: target });
        }
      }
    }
  }
  return rules;
}

/** Xây dựng từ điển học tập từ các item đã map[cite: 2] */
function buildLearnedNameDictionary_(mappingData, idxRaw, idxSource, idxName) {
  const learnedNameMap = new Map();
  for (let i = 1; i < mappingData.length; i++) {
    const rawVal    = mappingData[i][idxRaw] ? mappingData[i][idxRaw].toString().trim() : "";
    const sourceVal = mappingData[i][idxSource] ? mappingData[i][idxSource].toString().trim() : "";
    const nameVal   = mappingData[i][idxName] ? mappingData[i][idxName].toString().trim() : "";

    if (rawVal && sourceVal && nameVal && !nameVal.toLowerCase().includes("unmapped")) {
      const sourceGroup   = getSourceGroup_(sourceVal);
      const normalizedKey = `${cleanSpecWords_(rawVal)}|${sourceGroup.toLowerCase()}`;
      if (!learnedNameMap.has(normalizedKey)) {
        learnedNameMap.set(normalizedKey, nameVal);
      }
    }
  }
  return learnedNameMap;
}

/** Hiển thị thông báo kết quả đồng bộ[cite: 2] */
function showSyncResultMessage_(ui, duplicateCount, newCount) {
  let msg = [];
  if (duplicateCount > 0) {
    msg.push(`🧹 Đã dọn dẹp ${duplicateCount} dòng dữ liệu lặp cũ trong MAPPING.`);
  }
  if (newCount > 0) {
    msg.push(`✅ Đã thêm mới ${newCount} item(s) mới vào 'MAPPING'.`);
  }
  if (msg.length === 0) {
    msg.push("ℹ️ Dữ liệu MAPPING đã chuẩn hóa hoàn toàn và không có item mới.");
  }

  ui.alert("Kết quả Đồng bộ", msg.join("\n"), ui.ButtonSet.OK);
}
