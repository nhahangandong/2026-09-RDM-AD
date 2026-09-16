class ConfigLoader {
  /**
   * Tải toàn bộ cấu hình theo module từ bảng SYSTEM_CONFIG dựa vào Schema Map
   * @param {string} moduleName - Tên module cần lấy cấu hình (VD: "THEORETICAL")
   * @returns {Map<string, Array<string>>} - Trả về Map chứa các key và danh sách giá trị tương ứng
   */
  static getConfigMap(moduleName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const schemaMap = schemaGetMap();
    if (!schemaMap) return new Map();

    const sheetName = schemaGetSheetName(schemaMap, "SYSTEM_CONFIG");
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return new Map();

    const data = sheet.getDataRange().getValues();
    const getIdx = (cKey) => schemaGetColIndex(schemaMap, "SYSTEM_CONFIG", cKey);
    
    const idxModule = getIdx("module_name");
    const idxKey    = getIdx("config_key");
    const idxVal    = getIdx("config_value");
    const idxStatus = getIdx("status");

    const configMap = new Map();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const mod = idxModule !== -1 ? String(row[idxModule]).trim().toUpperCase() : "";
      const key = idxKey !== -1 ? String(row[idxKey]).trim().toUpperCase() : "";
      const val = idxVal !== -1 ? String(row[idxVal]).trim() : "";
      const status = idxStatus !== -1 ? String(row[idxStatus]).trim().toUpperCase() : "ACTIVE";

      // Chỉ lấy các cấu hình khớp module, đang ACTIVE và có khóa
      if (mod === moduleName.toUpperCase() && status === "ACTIVE" && key) {
        if (!configMap.has(key)) {
          configMap.set(key, []);
        }
        configMap.get(key).push(val);
      }
    }
    return configMap;
  }
}
