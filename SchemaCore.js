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

