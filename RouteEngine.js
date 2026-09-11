/**
 * Module RouteEngine: Xử lý tra cứu tuyến giao dịch động qua SCHEMA.
 */
const RouteEngine = {

  /**
   * Tải bản đồ tra cứu: location_code -> location_type
   */
  getLocationTypeMap: function(ss, schemaMap) {
    const sheetName = schemaGetSheetName(schemaMap, "LOCATION");
    const sheet = sheetName ? ss.getSheetByName(sheetName) : null;
    if (!sheet) return new Map();

    const data = sheet.getDataRange().getValues();
    const idxCode = schemaGetColIndex(schemaMap, "LOCATION", "location_code");
    const idxType = schemaGetColIndex(schemaMap, "LOCATION", "location_type");

    const map = new Map();
    for (let i = 1; i < data.length; i++) {
      const code = cleanCodeValue_(data[i][idxCode]);
      const type = cleanCodeValue_(data[i][idxType]);
      if (code) map.set(code, type);
    }
    return map;
  },

  /**
   * Tải bản đồ quy tắc: "FROM_TYPE->TO_TYPE" -> Thông tin tuyến
   */
  getRouteRulesMap: function(ss, schemaMap) {
    const sheetName = schemaGetSheetName(schemaMap, "ROUTE_MAP");
    const sheet = sheetName ? ss.getSheetByName(sheetName) : null;
    if (!sheet) return new Map();

    const data = sheet.getDataRange().getValues();
    const idxFromType = schemaGetColIndex(schemaMap, "ROUTE_MAP", "from_type");
    const idxToType   = schemaGetColIndex(schemaMap, "ROUTE_MAP", "to_type");
    const idxCode     = schemaGetColIndex(schemaMap, "ROUTE_MAP", "route_code");
    const idxTxType   = schemaGetColIndex(schemaMap, "ROUTE_MAP", "transaction_type");
    const idxIsInv    = schemaGetColIndex(schemaMap, "ROUTE_MAP", "is_inventory");

    const map = new Map();
    for (let i = 1; i < data.length; i++) {
      const fromType = cleanCodeValue_(data[i][idxFromType]);
      const toType   = cleanCodeValue_(data[i][idxToType]);

      if (fromType && toType) {
        const key = `${fromType}->${toType}`;
        const isInvVal = data[i][idxIsInv];
        const isTrueInv = (isInvVal === true || isInvVal.toString().toUpperCase() === "TRUE" || isInvVal === 1);

        map.set(key, {
          routeCode: cleanCodeValue_(data[i][idxCode]),
          transactionType: cleanCodeValue_(data[i][idxTxType]),
          isInventory: isTrueInv
        });
      }
    }
    return map;
  },

  /**
   * Xác định Route từ 2 mã địa điểm (fromCode, toCode)
   */
  getRoute: function(ss, schemaMap, fromCode, toCode) {
    const locationTypeMap = this.getLocationTypeMap(ss, schemaMap);
    const routeRulesMap   = this.getRouteRulesMap(ss, schemaMap);

    const fromType = locationTypeMap.get(fromCode) || "UNKNOWN";
    const toType   = locationTypeMap.get(toCode) || "UNKNOWN";

    const key = `${fromType}->${toType}`;
    return routeRulesMap.get(key) || {
      routeCode: "UNKNOWN_ROUTE",
      transactionType: "UNKNOWN",
      isInventory: false
    };
  }
};
