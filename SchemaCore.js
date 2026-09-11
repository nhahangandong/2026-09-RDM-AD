/**
 * ============================================================================
 * MODULE: SchemaCore.gs
 * MỤC ĐÍCH: Cung cấp các API hệ thống dùng chung để tra cứu Lược đồ (SCHEMA).
 * QUY CHUẨN ĐẶT TÊN: 
 *   - Global API: schema[Action][Entity] (VD: schemaGetMap)
 *   - Private Helper: [functionName]_ (VD: parseRowData_)
 * ============================================================================
 */

/**
 * Global API: Đọc toàn bộ cấu trúc SCHEMA và caching dạng Map object.
 * @return {Object|null} Schema Map object
 */
function schemaGetMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schemaSheet = ss.getSheetByName("SCHEMA");
  if (!schemaSheet) return null;

  const data = schemaSheet.getDataRange().getValues();
  if (data.length <= 1) return null;

  const headerRow = data[0].map(h =>
    h ? h.toString().trim().toLowerCase() : ""
  );

  const idxSchemaName = headerRow.indexOf("schema_name");
  const idxSheetName = headerRow.indexOf("sheet_name");
  const idxColIndex = headerRow.indexOf("col_index");
  const idxColKey = headerRow.indexOf("col_key");
  const idxColHeader = headerRow.indexOf("col_header");
  const idxDataType = headerRow.indexOf("data_type");

  // schema_name và sheet_name là bắt buộc
  if (idxSchemaName === -1 || idxSheetName === -1) {
    throw new Error(
      "SCHEMA phải có schema_name và sheet_name."
    );
  }

  if (idxColIndex === -1 || idxColKey === -1) {
    throw new Error(
      "SCHEMA phải có col_index và col_key."
    );
  }

  const schemaMap = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const schemaName = row[idxSchemaName]
      ? row[idxSchemaName].toString().trim()
      : "";

    const sheetName = row[idxSheetName]
      ? row[idxSheetName].toString().trim()
      : "";

    const colIndex = parseInt(row[idxColIndex], 10);

    const colKey = row[idxColKey]
      ? row[idxColKey].toString().trim()
      : "";

    const colHeader = idxColHeader !== -1 && row[idxColHeader]
      ? row[idxColHeader].toString().trim()
      : "";

    const dataType = idxDataType !== -1 && row[idxDataType]
      ? row[idxDataType].toString().trim()
      : "String";

    if (
      !schemaName ||
      !sheetName ||
      isNaN(colIndex) ||
      !colKey
    ) {
      continue;
    }

    if (!schemaMap[schemaName]) {
      schemaMap[schemaName] = {
        schemaName: schemaName,
        sheetName: sheetName,
        byKey: {},
        columns: []
      };
    }

    const schema = schemaMap[schemaName];

    // Phát hiện schema_name -> nhiều sheet_name
    if (schema.sheetName !== sheetName) {
      throw new Error(
        `Schema "${schemaName}" có nhiều sheet_name: ` +
        `"${schema.sheetName}" và "${sheetName}".`
      );
    }

    const colObj = {
      colIndex: colIndex,
      colKey: colKey,
      colHeader: colHeader,
      dataType: dataType
    };

    schema.byKey[colKey] = colObj;
    schema.columns.push(colObj);
  }

  Object.keys(schemaMap).forEach(schemaName => {
    schemaMap[schemaName].columns.sort(
      (a, b) => a.colIndex - b.colIndex
    );
  });

  return schemaMap;
}

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

/**
 * Global API: Lấy Tên sheet vật lý (sheet_name) từ schema_name.
 * @param {Object} schemaMap - Schema Map từ schemaGetMap()
 * @param {string} schemaName - Mã schema logic (VD: 'MENU', 'ITEM_MASTER')
 * @return {string|null} Tên sheet thực tế trên Google Sheets
 */
function schemaGetSheetName(schemaMap, schemaName) {
  if (!schemaMap || !schemaMap[schemaName]) return null;
  return schemaMap[schemaName].sheetName;
}

/**
 * Global API: Lấy 0-based Array Index của một col_key trong sheet tương ứng.
 * @param {Object} schemaMap - Schema Map từ schemaGetMap()
 * @param {string} sheetName - Tên sheet trong SCHEMA
 * @param {string} colKey - Mã khóa cột
 * @return {number} Index mảng (0-based), -1 nếu không tìm thấy
 */
// function schemaGetColIndex(schemaMap, sheetName, colKey) {
//   if (!schemaMap || !schemaMap[sheetName] || !schemaMap[sheetName].byKey[colKey]) {
//     return -1;
//   }
//   return schemaMap[sheetName].byKey[colKey].colIndex - 1;
// }
function schemaGetColIndex(schemaMap, schemaName, colKey) {
  if (
    !schemaMap ||
    !schemaMap[schemaName] ||
    !schemaMap[schemaName].byKey[colKey]
  ) {
    return -1;
  }

  return schemaMap[schemaName].byKey[colKey].colIndex - 1;
}
/**
 * Global API: Lấy thông tin chi tiết của một cột theo col_key.
 * @param {Object} schemaMap - Schema Map từ schemaGetMap()
 * @param {string} sheetName - Tên sheet
 * @param {string} colKey - Mã khóa cột
 * @return {Object|null} { colIndex, colKey, colHeader, dataType }
 */
// function schemaGetColMeta(schemaMap, sheetName, colKey) {
//   if (!schemaMap || !schemaMap[sheetName] || !schemaMap[sheetName].byKey[colKey]) {
//     return null;
//   }
//   return schemaMap[sheetName].byKey[colKey];
// }
function schemaGetColMeta(schemaMap, schemaName, colKey) {
  if (
    !schemaMap ||
    !schemaMap[schemaName] ||
    !schemaMap[schemaName].byKey[colKey]
  ) {
    return null;
  }

  return schemaMap[schemaName].byKey[colKey];
}
// ========================================================================================




/**
 * Tự động đồng bộ Sheet và Google Sheets Table từ SCHEMA.
 * Tự động xóa Table cũ (nếu có) trước khi tạo Table mới, KHÔNG CẦN XÓA SHEET THỦ CÔNG.
 * Tu dong tao sheet, truyen dung Ten Cot tu SCHEMA vao GSheet Table va dat ten Table.
 */
function setupSheetsFromSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  const schemaSheet = ss.getSheetByName("SCHEMA");
  const ui = SpreadsheetApp.getUi();

  if (!schemaSheet) {
    ui.alert("⚠️ Thông báo", "Không tìm thấy sheet 'SCHEMA'!", ui.ButtonSet.OK);
    return;
  }

  const schemaData = schemaSheet.getDataRange().getValues();
  if (schemaData.length <= 1) {
    ui.alert("⚠️ Thông báo", "Sheet 'SCHEMA' đang trống!", ui.ButtonSet.OK);
    return;
  }

  // 1. Gom nhóm danh sách cột theo từng sheet_name từ SCHEMA (BỎ QUA SHEET SCHEMA VÀ MAPPING NẾU CẦN)
  const mapSchema = {};
  for (let i = 1; i < schemaData.length; i++) {
    const row = schemaData[i];
    const sheetName = row[0] ? row[0].toString().trim() : "";
    const colIndex = parseInt(row[1], 10);
    const colHeader = row[3] ? row[3].toString().trim() : "";

    if (sheetName && sheetName.toUpperCase() !== "SCHEMA" && !isNaN(colIndex) && colHeader) {
      if (!mapSchema[sheetName]) {
        mapSchema[sheetName] = [];
      }
      mapSchema[sheetName].push({
        index: colIndex,
        header: colHeader
      });
    }
  }

  // 2. Lấy danh sách các Table ĐÃ TỒN TẠI để KHÔNG XÓA NHỮNG BẢNG ĐANG CÓ DỮ LIỆU
  const existingTableSheetIds = new Set();
  try {
    const spreadsheetInfo = Sheets.Spreadsheets.get(spreadsheetId, { includeGridData: false });
    spreadsheetInfo.sheets.forEach(s => {
      if (s.tables && s.tables.length > 0) {
        existingTableSheetIds.add(s.properties.sheetId);
      }
    });
  } catch (err) {
    Logger.log("Chưa thể lấy danh sách Tables: " + err.message);
  }

  const addTableRequests = [];
  let createdSheetsCount = 0;
  let processedTablesCount = 0;

  // 3. Lặp qua các Sheet khai báo trong SCHEMA
  Object.keys(mapSchema).forEach(sheetName => {
    let targetSheet = ss.getSheetByName(sheetName);

    // Nếu chưa có Sheet thì tạo mới
    if (!targetSheet) {
      targetSheet = ss.insertSheet(sheetName);
      createdSheetsCount++;
    }

    const sheetId = targetSheet.getSheetId();

    // NẾU SHEET NÀY ĐÃ CÓ TABLE RỒI -> BỎ QUA KHÔNG CAN THIỆP ĐỂ BẢO VỆ DỮ LIỆU
    if (existingTableSheetIds.has(sheetId)) {
      return;
    }

    // Sắp xếp các cột theo col_index
    const columns = mapSchema[sheetName].sort((a, b) => a.index - b.index);
    const headers = columns.map(col => col.header);

    let lastRow = targetSheet.getLastRow();

    // Nếu Sheet hoàn toàn trống (mới tạo) -> Ghi Header vào dòng 1
    if (lastRow === 0) {
      targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      lastRow = 1;
    }

    // Phạm vi dòng của Table: Từ dòng 1 tới dòng dữ liệu cuối cùng (tối thiểu là dòng 2)
    const endRowIndex = Math.max(lastRow, 2);
    const tableName = `tbl_${sheetName.toLowerCase().trim().replace(/[^a-z0-9_]/g, "_")}`;

    // Tạo Request addTable mà KHÔNG XÓA DỮ LIỆU CŨ
    addTableRequests.push({
      addTable: {
        table: {
          name: tableName,
          range: {
            sheetId: sheetId,
            startRowIndex: 0,
            endRowIndex: endRowIndex,
            startColumnIndex: 0,
            endColumnIndex: headers.length
          }
        }
      }
    });

    processedTablesCount++;
  });

  // 4. Thực thi Batch Update tạo Table mới cho các Sheet chưa có Table
  if (addTableRequests.length > 0) {
    try {
      Sheets.Spreadsheets.batchUpdate({ requests: addTableRequests }, spreadsheetId);
    } catch (e) {
      Logger.log("Lỗi thực thi API Table: " + e.message);
      ui.alert("⚠️ Lỗi API Table", e.message, ui.ButtonSet.OK);
      return;
    }
  }

  ui.alert(
    "✅ Thành công",
    `Đã hoàn tất đồng bộ Table từ SCHEMA (An toàn 100% dữ liệu):\n- Số sheet mới tạo: ${createdSheetsCount}\n- Số sheet đã tạo Bảng mới: ${processedTablesCount}`,
    ui.ButtonSet.OK
  );
}
// function setupSheetsFromSchema() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const spreadsheetId = ss.getId();
//   const schemaSheet = ss.getSheetByName("SCHEMA");
//   const ui = SpreadsheetApp.getUi();
  
//   if (!schemaSheet) {
//     ui.alert("⚠️ Thong bao", "Khong tim thay sheet 'SCHEMA'!", ui.ButtonSet.OK);
//     return;
//   }

//   const schemaData = schemaSheet.getDataRange().getValues();
//   if (schemaData.length <= 1) {
//     ui.alert("⚠️ Thong bao", "Sheet 'SCHEMA' dang trong!", ui.ButtonSet.OK);
//     return;
//   }

//   // 1. Gom nhom danh sach cot theo tung sheet_name tu SCHEMA
//   const mapSchema = {};
//   for (let i = 1; i < schemaData.length; i++) {
//     const row = schemaData[i];
//     const sheetName = row[0] ? row[0].toString().trim() : "";
//     const colIndex = parseInt(row[1], 10);
//     const colHeader = row[3] ? row[3].toString().trim() : "";

//     if (sheetName && !isNaN(colIndex) && colHeader) {
//       if (!mapSchema[sheetName]) {
//         mapSchema[sheetName] = [];
//       }
//       mapSchema[sheetName].push({
//         index: colIndex,
//         header: colHeader
//       });
//     }
//   }

//   // 2. Lay danh sach cac Table hien co de xoa truoc khi tao lai
//   let existingTablesMap = {};
//   try {
//     const spreadsheetInfo = Sheets.Spreadsheets.get(spreadsheetId, { includeGridData: false });
//     spreadsheetInfo.sheets.forEach(s => {
//       if (s.tables && s.tables.length > 0) {
//         existingTablesMap[s.properties.sheetId] = s.tables.map(tbl => tbl.tableId);
//       }
//     });
//   } catch (err) {
//     Logger.log("Chua the lay thong tin Tables: " + err.message);
//   }

//   const requests = [];
//   let createdSheetsCount = 0;
//   let processedTablesCount = 0;

//   // 3. Lap qua tung Sheet khai bao trong SCHEMA
//   Object.keys(mapSchema).forEach(sheetName => {
//     let targetSheet = ss.getSheetByName(sheetName);

//     if (!targetSheet) {
//       targetSheet = ss.insertSheet(sheetName);
//       createdSheetsCount++;
//     }

//     const sheetId = targetSheet.getSheetId();

//     // Xoa Table cu neu da ton tai tren sheet
//     if (existingTablesMap[sheetId]) {
//       existingTablesMap[sheetId].forEach(tableId => {
//         requests.push({
//           deleteTable: {
//             tableId: tableId
//           }
//         });
//       });
//     }

//     // Sap xep cac cot theo col_index
//     const columns = mapSchema[sheetName].sort((a, b) => a.index - b.index);
//     const headers = columns.map(col => col.header);

//     // Ghi Header vao dong 1
//     const headerRange = targetSheet.getRange(1, 1, 1, headers.length);
//     headerRange.setValues([headers]);

//     // Format danh sach columnProperties de dinh danh Header cho Table API
//     const columnProperties = headers.map(h => ({ name: h }));

//     const tableName = `tbl_${sheetName.toLowerCase().trim()}`;

//     // Tao Request addTable voi columnProperties chuan
//     requests.push({
//       addTable: {
//         table: {
//           name: tableName,
//           range: {
//             sheetId: sheetId,
//             startRowIndex: 0,          // Dong 1
//             endRowIndex: 2,            // Khởi tạo đến dòng 2 (Header + 1 dòng dữ liệu)
//             startColumnIndex: 0,
//             endColumnIndex: headers.length
//           },
//           columnProperties: columnProperties // Định danh chính xác tiêu đề cột cho Table
//         }
//       }
//     });

//     processedTablesCount++;
//   });

//   // 4. Thuc thi Batch Update
//   if (requests.length > 0) {
//     try {
//       Sheets.Spreadsheets.batchUpdate({ requests: requests }, spreadsheetId);
//     } catch (e) {
//       Logger.log("Loi thuc thi API Table: " + e.message);
//     }
//   }

//   ui.alert(
//     "✅ Thanh cong",
//     `Da dong bo thanh cong GSheet Table tu SCHEMA:\n- So sheet moi duoc tao: ${createdSheetsCount}\n- So sheet da cap nhat Table: ${processedTablesCount}`,
//     ui.ButtonSet.OK
//   );
// }
