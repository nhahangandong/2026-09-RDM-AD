/**
 * Module RouteEngine: Xử lý tra cứu tuyến giao dịch động qua SCHEMA.
 * Luồng xử lý: Transaction (codes) -> Location Map (types) -> Route Map (rules)
 */
class RouteEngine {

  /**
   * Tải bản đồ tra cứu: location_code -> location_type từ SCHEMA "LOCATION_MAP"
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
   * @param {Object} schemaMap
   * @returns {Map<string, string>} Map(location_code => location_type)
   */
  static getLocationTypeMap(ss, schemaMap) {
    const SCHEMA_KEY = "LOCATION_MAP";
    const sheetName = schemaGetSheetName(schemaMap, SCHEMA_KEY);
    const sheet = sheetName ? ss.getSheetByName(sheetName) : null;
    if (!sheet) return new Map();

    const data = sheet.getDataRange().getValues();
    const idxCode = schemaGetColIndex(schemaMap, SCHEMA_KEY, "location_code");
    const idxType = schemaGetColIndex(schemaMap, SCHEMA_KEY, "location_type");

    const map = new Map();
    for (let i = 1; i < data.length; i++) {
      const code = cleanCodeValue_(data[i][idxCode]);
      const type = cleanCodeValue_(data[i][idxType]);
      if (code) map.set(code, type.toUpperCase());
    }
    return map;
  }

  /**
   * Tải bản đồ quy tắc tuyến: "FROM_TYPE->TO_TYPE" -> Quy tắc từ SCHEMA "ROUTE_MAP"
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
   * @param {Object} schemaMap
   * @returns {Map<string, {routeCode: string, transactionType: string, isInventory: boolean}>}
   */
  static getRouteRulesMap(ss, schemaMap) {
    const SCHEMA_KEY = "ROUTE_MAP";
    const sheetName = schemaGetSheetName(schemaMap, SCHEMA_KEY);
    const sheet = sheetName ? ss.getSheetByName(sheetName) : null;
    if (!sheet) return new Map();

    const data = sheet.getDataRange().getValues();

    // Tự động nhận diện từ from_type/from_code và to_type/to_code
    let idxFromType = schemaGetColIndex(schemaMap, SCHEMA_KEY, "from_type");
    if (idxFromType === -1) idxFromType = schemaGetColIndex(schemaMap, SCHEMA_KEY, "from_code");

    let idxToType = schemaGetColIndex(schemaMap, SCHEMA_KEY, "to_type");
    if (idxToType === -1) idxToType = schemaGetColIndex(schemaMap, SCHEMA_KEY, "to_code");

    const idxCode = schemaGetColIndex(schemaMap, SCHEMA_KEY, "route_code");

    let idxTxType = schemaGetColIndex(schemaMap, SCHEMA_KEY, "trans_type");
    if (idxTxType === -1) idxTxType = schemaGetColIndex(schemaMap, SCHEMA_KEY, "transaction_type");

    const idxIsInv = schemaGetColIndex(schemaMap, SCHEMA_KEY, "is_inventory");

    const map = new Map();
    for (let i = 1; i < data.length; i++) {
      const fromType = cleanCodeValue_(data[i][idxFromType]).toUpperCase();
      const toType   = cleanCodeValue_(data[i][idxToType]).toUpperCase();

      if (fromType && toType) {
        const key = `${fromType}->${toType}`;
        const isInvVal = data[i][idxIsInv];
        const isTrueInv = (isInvVal === true || String(isInvVal).toUpperCase() === "TRUE" || isInvVal === 1);

        map.set(key, {
          routeCode: idxCode !== -1 ? cleanCodeValue_(data[i][idxCode]) : "",
          transactionType: idxTxType !== -1 ? cleanCodeValue_(data[i][idxTxType]).toUpperCase() : "",
          isInventory: isTrueInv
        });
      }
    }
    return map;
  }

  /**
   * Xác định Route từ 2 mã địa điểm (fromCode, toCode)
   * 1. Tra cứu location_code -> location_type từ LOCATION_MAP
   * 2. Tra cứu cặp (from_type -> to_type) từ ROUTE_MAP
   */
  static getRoute(ss, schemaMap, fromCode, toCode) {
    const locationTypeMap = RouteEngine.getLocationTypeMap(ss, schemaMap);
    const routeRulesMap   = RouteEngine.getRouteRulesMap(ss, schemaMap);

    const fromClean = cleanCodeValue_(fromCode);
    const toClean   = cleanCodeValue_(toCode);

    // Chuyển đổi Mã địa điểm thành Loại địa điểm
    const fromType  = (locationTypeMap.get(fromClean) || "UNKNOWN").toUpperCase();
    const toType    = (locationTypeMap.get(toClean) || "UNKNOWN").toUpperCase();

    // Tra cứu quy tắc theo Cặp loại địa điểm
    const routeKey  = `${fromType}->${toType}`;
    
    return routeRulesMap.get(routeKey) || {
      routeCode: "UNKNOWN_ROUTE",
      transactionType: "UNKNOWN",
      isInventory: false
    };
  }
}
