// /**
//  * ============================================================================
//  * MODULE: SchemaCore.gs
//  * MỤC ĐÍCH: Cung cấp các API hệ thống dùng chung để tra cứu Lược đồ (SCHEMA).
//  * QUY CHUẨN ĐẶT TÊN: 
//  *   - Global API: schema[Action][Entity] (VD: schemaGetMap)
//  *   - Private Helper: [functionName]_ (VD: parseRowData_)
//  * ============================================================================
//  */

// /**
//  * Global API: Đọc toàn bộ cấu trúc SCHEMA và caching dạng Map object.
//  * @return {Object|null} Schema Map object
//  */
// function schemaGetMap() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const schemaSheet = ss.getSheetByName("SCHEMA");
//   if (!schemaSheet) return null;

//   const data = schemaSheet.getDataRange().getValues();
//   if (data.length <= 1) return null;

//   const headerRow = data[0].map(h => h ? h.toString().trim().toLowerCase() : "");
//   const idxSheetName = headerRow.indexOf("schema_name") !== -1 ? headerRow.indexOf("schema_name") : headerRow.indexOf("sheet_name");
//   const idxColIndex = headerRow.indexOf("col_index");
//   const idxColKey = headerRow.indexOf("col_key");
//   const idxColHeader = headerRow.indexOf("col_header");
//   const idxDataType = headerRow.indexOf("data_type");

//   const schemaMap = {};

//   for (let i = 1; i < data.length; i++) {
//     const row = data[i];
//     const sheetName = row[idxSheetName] ? row[idxSheetName].toString().trim() : "";
//     const colIndex = parseInt(row[idxColIndex], 10);
//     const colKey = row[idxColKey] ? row[idxColKey].toString().trim() : "";
//     const colHeader = row[idxColHeader] ? row[idxColHeader].toString().trim() : "";
//     const dataType = idxDataType !== -1 && row[idxDataType] ? row[idxDataType].toString().trim() : "String";

//     if (!sheetName || isNaN(colIndex) || !colKey) continue;

//     if (!schemaMap[sheetName]) {
//       schemaMap[sheetName] = {
//         byKey: {},
//         columns: []
//       };
//     }

//     const colObj = {
//       colIndex: colIndex, // 1-based index
//       colKey: colKey,
//       colHeader: colHeader,
//       dataType: dataType
//     };

//     schemaMap[sheetName].byKey[colKey] = colObj;
//     schemaMap[sheetName].columns.push(colObj);
//   }

//   // Sort các cột theo col_index
//   Object.keys(schemaMap).forEach(s => {
//     schemaMap[s].columns.sort((a, b) => a.colIndex - b.colIndex);
//   });

//   return schemaMap;
// }

// /**
//  * Global API: Lấy 0-based Array Index của một col_key trong sheet tương ứng.
//  * @param {Object} schemaMap - Schema Map từ schemaGetMap()
//  * @param {string} sheetName - Tên sheet trong SCHEMA
//  * @param {string} colKey - Mã khóa cột
//  * @return {number} Index mảng (0-based), -1 nếu không tìm thấy
//  */
// function schemaGetColIndex(schemaMap, sheetName, colKey) {
//   if (!schemaMap || !schemaMap[sheetName] || !schemaMap[sheetName].byKey[colKey]) {
//     return -1;
//   }
//   return schemaMap[sheetName].byKey[colKey].colIndex - 1;
// }

// /**
//  * Global API: Lấy thông tin chi tiết của một cột theo col_key.
//  * @param {Object} schemaMap - Schema Map từ schemaGetMap()
//  * @param {string} sheetName - Tên sheet
//  * @param {string} colKey - Mã khóa cột
//  * @return {Object|null} { colIndex, colKey, colHeader, dataType }
//  */
// function schemaGetColMeta(schemaMap, sheetName, colKey) {
//   if (!schemaMap || !schemaMap[sheetName] || !schemaMap[sheetName].byKey[colKey]) {
//     return null;
//   }
//   return schemaMap[sheetName].byKey[colKey];
// }



// // /**
// //  * Core Helper: Đọc toàn bộ cấu trúc SCHEMA và caching dạng Map object.
// //  * Sử dụng chung cho toàn hệ thống để tra cứu vị trí cột theo col_key hoặc sheet_name.
// //  * @return {Object} Schema Map structure
// //  */
// // function getSchemaMap() {
// //   const ss = SpreadsheetApp.getActiveSpreadsheet();
// //   const schemaSheet = ss.getSheetByName("SCHEMA");
// //   if (!schemaSheet) return null;

// //   const data = schemaSheet.getDataRange().getValues();
// //   if (data.length <= 1) return null;

// //   // Lấy dòng header của SCHEMA để map đúng vị trí các cột khai báo meta
// //   const headerRow = data[0].map(h => h ? h.toString().trim().toLowerCase() : "");
// //   const idxSheetName = headerRow.indexOf("schema_name") !== -1 ? headerRow.indexOf("schema_name") : headerRow.indexOf("sheet_name");
// //   const idxColIndex = headerRow.indexOf("col_index");
// //   const idxColKey = headerRow.indexOf("col_key");
// //   const idxColHeader = headerRow.indexOf("col_header");
// //   const idxDataType = headerRow.indexOf("data_type");

// //   const schemaMap = {};

// //   for (let i = 1; i < data.length; i++) {
// //     const row = data[i];
// //     const sheetName = row[idxSheetName] ? row[idxSheetName].toString().trim() : "";
// //     const colIndex = parseInt(row[idxColIndex], 10);
// //     const colKey = row[idxColKey] ? row[idxColKey].toString().trim() : "";
// //     const colHeader = row[idxColHeader] ? row[idxColHeader].toString().trim() : "";
// //     const dataType = idxDataType !== -1 && row[idxDataType] ? row[idxDataType].toString().trim() : "String";

// //     if (!sheetName || isNaN(colIndex) || !colKey) continue;

// //     if (!schemaMap[sheetName]) {
// //       schemaMap[sheetName] = {
// //         byKey: {},    // Tra cứu theo col_key: schemaMap['SALES'].byKey['raw_name'] -> { colIndex, colHeader }
// //         columns: []   // Danh sách cột đã sort theo col_index
// //       };
// //     }

// //     const colObj = {
// //       colIndex: colIndex,   // 1-based index (1, 2, 3...)
// //       colKey: colKey,
// //       colHeader: colHeader,
// //       dataType: dataType
// //     };

// //     schemaMap[sheetName].byKey[colKey] = colObj;
// //     schemaMap[sheetName].columns.push(colObj);
// //   }

// //   // Sort danh sách cột theo col_index tăng dần
// //   Object.keys(schemaMap).forEach(s => {
// //     schemaMap[s].columns.sort((a, b) => a.colIndex - b.colIndex);
// //   });

// //   return schemaMap;
// // }

// // /**
// //  * Core Helper: Lấy 0-based Array Index của một col_key trong một sheet cụ thể.
// //  * @param {Object} schemaMap - Object trả về từ getSchemaMap()
// //  * @param {string} sheetName - Tên sheet (VD: "SALES", "MAPPING")
// //  * @param {string} colKey - Khóa cột (VD: "raw_name", "item_code")
// //  * @return {number} Index trong mảng (0-based), trả về -1 nếu không tìm thấy
// //  */
// // function getColumnIndexByKey(schemaMap, sheetName, colKey) {
// //   if (!schemaMap || !schemaMap[sheetName] || !schemaMap[sheetName].byKey[colKey]) {
// //     return -1;
// //   }
// //   return schemaMap[sheetName].byKey[colKey].colIndex - 1;
// // }


// // /**
// //  * Lấy Map vị trí cột (1-based index) linh hoạt theo col_key
// //  * @param {string} targetSheetName - Tên sheet cần lấy (VD: 'TRANSACTION')
// //  * @returns {Object} Ví dụ: { trans_no: 1, trans_date: 2, line_no: 5, ... }
// //  */
// // function getColumnIndices(targetSheetName) {
// //   const ss = SpreadsheetApp.getActiveSpreadsheet();
// //   const schemaSheet = ss.getSheetByName('SCHEMA');
// //   const targetSheet = ss.getSheetByName(targetSheetName);
  
// //   if (!schemaSheet) throw new Error("Khong tim thay sheet SCHEMA");

// //   // 1. Đọc dữ liệu từ SCHEMA
// //   const schemaData = schemaSheet.getDataRange().getValues();
// //   // Giả định SCHEMA cột: [0] sheet_name, [1] col_index, [2] col_key, [3] col_header
// //   const schemaMap = new Map();
  
// //   for (let i = 1; i < schemaData.length; i++) {
// //     const [sheetName, colIdx, colKey, colHeader] = schemaData[i];
// //     if (sheetName === targetSheetName && colKey) {
// //       schemaMap.set(colHeader.toString().trim(), {
// //         colKey: colKey.toString().trim(),
// //         defaultIndex: Number(colIdx)
// //       });
// //     }
// //   }

// //   const colIndices = {};

// //   // 2. Nếu sheet mục tiêu đã tồn tại, đọc trực tiếp dòng Header (Dòng 1) để lấy vị trí thực tế
// //   if (targetSheet && targetSheet.getLastColumn() > 0) {
// //     const headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    
// //     headers.forEach((headerText, index) => {
// //       const cleanHeader = headerText.toString().trim();
// //       if (schemaMap.has(cleanHeader)) {
// //         const key = schemaMap.get(cleanHeader).colKey;
// //         colIndices[key] = index + 1; // Index dạng 1-based (Cột A = 1)
// //       }
// //     });
// //   }

// //   // 3. Với các key chưa tìm thấy trên Header thực tế, fallback dùng col_index trong SCHEMA
// //   schemaMap.forEach((info) => {
// //     if (!colIndices[info.colKey]) {
// //       colIndices[info.colKey] = info.defaultIndex;
// //     }
// //   });

// //   return colIndices;
// // }

// // /**
// //  * Lấy danh sách cột đã sắp xếp theo col_index từ SCHEMA
// //  * @param {string} targetSheetName - Tên sheet cần lấy (VD: 'TRANSACTION')
// //  * @returns {Array<Object>} Mảng cấu trúc cột [{ colIndex, colKey, colHeader, dataType }, ...]
// //  */
// // function getSchemaColumns(targetSheetName) {
// //   const ss = SpreadsheetApp.getActiveSpreadsheet();
// //   const schemaSheet = ss.getSheetByName('SCHEMA');
// //   if (!schemaSheet) throw new Error("Không tìm thấy sheet SCHEMA!");

// //   const schemaData = schemaSheet.getDataRange().getValues();
// //   const columns = [];

// //   // Lặp từ dòng 2 (bỏ qua Header của SCHEMA)
// //   // Giả định SCHEMA có thứ tự cột: [0] sheet_name, [1] col_index, [2] col_key, [3] col_header, [4] data_type
// //   for (let i = 1; i < schemaData.length; i++) {
// //     const [sheetName, colIdx, colKey, colHeader, dataType] = schemaData[i];
// //     if (sheetName === targetSheetName && colKey) {
// //       columns.push({
// //         colIndex: Number(colIdx),
// //         colKey: colKey.toString().trim(),
// //         colHeader: colHeader.toString().trim(),
// //         dataType: dataType ? dataType.toString().trim() : 'String'
// //       });
// //     }
// //   }

// //   // Sắp xếp tăng dần theo col_index
// //   return columns.sort((a, b) => a.colIndex - b.colIndex);
// // }

// // /**
// //  * Tự động khởi tạo hoặc format lại Sheet theo cấu trúc trong SCHEMA
// //  * @param {string} sheetName - Tên sheet cần setup (VD: 'TRANSACTION')
// //  */
// // function setupSheetFromSchema(sheetName) {
// //   const ss = SpreadsheetApp.getActiveSpreadsheet();
// //   let sheet = ss.getSheetByName(sheetName);
  
// //   // Tạo sheet nếu chưa tồn tại
// //   if (!sheet) {
// //     sheet = ss.insertSheet(sheetName);
// //   }

// //   const columns = getSchemaColumns(sheetName);
// //   if (columns.length === 0) {
// //     Logger.log(`Không tìm thấy khai báo cấu trúc cho sheet: ${sheetName}`);
// //     return;
// //   }

// //   // Mảng chứa tên tiêu đề theo thứ tự col_index
// //   const headers = columns.map(col => col.colHeader);

// //   // Ghi dòng Header (Row 1)
// //   sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

// //   // Format cơ bản cho Header
// //   const headerRange = sheet.getRange(1, 1, 1, headers.length);
// //   headerRange.setBackground('#202124')
// //              .setFontColor('#ffffff')
// //              .setFontWeight('bold')
// //              .setHorizontalAlignment('center');

// //   sheet.setFrozenRows(1); // Cố định dòng tiêu đề
// //   Logger.log(`Khởi tạo thành công sheet [${sheetName}] với ${headers.length} cột.`);
// // }

// // /**
// //  * Hàm hỗ trợ tự động lấy hoặc tính toán Period (YYYYMM)
// //  * @param {Date|string} transDate - Ngày chứng từ
// //  * @param {string|number} manualPeriod - Giá trị kỳ người dùng tự gõ (nếu có)
// //  * @returns {string} Kỳ hạch toán dạng "YYYYMM"
// //  */
// // function resolvePeriod(transDate, manualPeriod) {
// //   // 1. Nếu người dùng đã gõ chủ động (VD: "202603" hoặc 202603)
// //   if (manualPeriod && manualPeriod.toString().trim() !== '') {
// //     return manualPeriod.toString().trim();
// //   }
  
// //   // 2. Nếu để trống, tự động lấy YYYYMM từ trans_date
// //   if (transDate instanceof Date && !isNaN(transDate.getTime())) {
// //     const year = transDate.getFullYear();
// //     const month = String(transDate.getMonth() + 1).padStart(2, '0');
// //     return `${year}${month}`;
// //   }
  
// //   return '';
// // }

// // /**
// //  * Tao chuoi Composite PK tu trans_no, line_no, item_code
// //  * @param {string} transNo - So chung tu
// //  * @param {number|string} lineNo - STT dong
// //  * @param {string} itemCode - Ma SKU
// //  * @returns {string} Chuoi PK duy nhat (VD: "260305-02_1_GASPC4501")
// //  */
// // function generateCompositeKey(transNo, lineNo, itemCode) {
// //   const cleanTransNo = transNo ? transNo.toString().trim() : '';
// //   const cleanLine = lineNo ? lineNo.toString().trim() : '0';
// //   const cleanCode = itemCode ? itemCode.toString().trim() : '';
  
// //   return `${cleanTransNo}_${cleanLine}_${cleanCode}`;
// // }