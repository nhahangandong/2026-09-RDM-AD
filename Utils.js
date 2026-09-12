/**
 * ============================================================================
 * FILE: Utils.gs
 * MỤC ĐÍCH: Chứa các hàm Helper Dùng chung / Private cho hệ thống.
 * Quy tắc đặt tên: functionName_ (có hậu tố underscore)
 * ============================================================================
 */

/**
 * Làm sạch và chuẩn hóa chuỗi dữ liệu đầu vào.
 * @param {*} val 
 * @returns {string}
 */
function cleanCodeValue_(val) {
  if (val === null || val === undefined) return "";
  return val.toString().trim();
}


/**
 * Helper Private: Bỏ dấu tiếng Việt và ký tự đặc biệt
 */
function removeVietnameseTones_(str) {
  if (!str) return "";
  str = str.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, "a");
  str = str.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, "e");
  str = str.replace(/ì|í|ị|ỉ|ĩ/g, "i");
  str = str.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g, "o");
  str = str.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, "u");
  str = str.replace(/ỳ|ý|ỵ|ỷ|ỹ/g, "y");
  str = str.replace(/đ/g, "d");
  str = str.replace(/À|Á|Ạ|Ả|Ã|Â|Ầ|Ấ|Ậ|Ẩ|Ẫ|Ă|Ằ|Ắ|Ặ|Ẳ|Ẵ/g, "A");
  str = str.replace(/È|É|Ẹ|Ẻ|Ẽ|Ê|Ề|Ế|Ệ|Ể|Ễ/g, "E");
  str = str.replace(/Ì|Í|Ị|Ỉ|Ĩ/g, "I");
  str = str.replace(/Ò|Ó|Ọ|Ỏ|Õ|Ô|Ồ|Ố|Ộ|Ổ|Ỗ|Ơ|Ờ|Ớ|Ợ|Ở|Ỡ/g, "O");
  str = str.replace(/Ù|Ú|Ụ|Ủ|Ũ|Ư|Ừ|Ứ|Ự|Ử|Ữ/g, "U");
  str = str.replace(/Ỳ|Ý|Ỵ|Ỷ|Ỹ/g, "Y");
  str = str.replace(/Đ/g, "D");
  return str;
}


/**
 * Chuyển đổi chuỗi pattern (chứa dấu đại diện *) thành RegEx động
 * Ví dụ: "*thịt*" -> /^.*thịt.*$/i
 */
function patternToRegex_(pattern) {
  if (!pattern || pattern.trim() === "*" || pattern.trim() === "") {
    return /^.*$/i;
  }

  const trimmed = pattern.trim();
  
  // Escape các ký tự đặc biệt trong RegEx (ngoại trừ dấu *)
  const escaped = trimmed.replace(/([.+?^=!:${}()|\[\]\/\\])/g, "\\$1");
  
  // Thay thế * thành .* và bọc trong mỏ neo ^ $
  const regexStr = "^" + escaped.replace(/\*/g, ".*") + "$";
  
  return new RegExp(regexStr, "i");
}


/**
 * Helper: Đọc dữ liệu từ sheet RAW_SOURCE_GROUP để tạo Map ánh xạ Nguồn/Sheet -> Nhóm nguồn.
 * @param {Object} schemaMap 
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss 
 * @returns {Map<string, string>} Map tra cứu Key (schema_name/sheet_name chữ thường) -> Value (source_group chữ thường)
 */
function getSourceGroupMap_(schemaMap, ss) {
  const sourceGroupMap = new Map();
  
  // Kiểm tra xem schema RAW_SOURCE_GROUP có tồn tại không
  if (!schemaMap["RAW_SOURCE_GROUP"]) {
    return sourceGroupMap;
  }

  const sheetName = schemaGetSheetName(schemaMap, "RAW_SOURCE_GROUP");
  const sheet = sheetName ? ss.getSheetByName(sheetName) : null;
  if (!sheet) return sourceGroupMap;

  const idxSchemaName  = schemaGetColIndex(schemaMap, "RAW_SOURCE_GROUP", "schema_name");
  const idxSheetName   = schemaGetColIndex(schemaMap, "RAW_SOURCE_GROUP", "sheet_name");
  const idxSourceGroup = schemaGetColIndex(schemaMap, "RAW_SOURCE_GROUP", "source_group");

  if (idxSourceGroup === -1) return sourceGroupMap;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const groupVal = data[i][idxSourceGroup] ? data[i][idxSourceGroup].toString().trim().toLowerCase() : "";
    if (!groupVal) continue;

    // Ánh xạ theo schema_name (ví dụ: TRANSACTION, STOCKTAKE, SALES)
    if (idxSchemaName !== -1 && data[i][idxSchemaName]) {
      const schemaKey = data[i][idxSchemaName].toString().trim().toLowerCase();
      sourceGroupMap.set(schemaKey, groupVal);
    }

    // Ánh xạ theo sheet_name (ví dụ: Mua hàng, Kiểm kê, Bán hàng)
    if (idxSheetName !== -1 && data[i][idxSheetName]) {
      const sheetKey = data[i][idxSheetName].toString().trim().toLowerCase();
      sourceGroupMap.set(sheetKey, groupVal);
    }
  }

  return sourceGroupMap;
}

/**
 * Helper: Tra cứu xem địa điểm Xuất -> Nhập có thuộc tuyến TÍNH TỒN KHO (is_inventory = true) hay không
 * Update: không còn cần sử dụng (sử dụng hàm nội tuyến)
 */
// function checkIsStockRouteByLocation_(fromCode, toCode, ss, schemaMap, validStockRoutes) {
//   if (!fromCode || !toCode) return false;

//   const locSheetName = schemaGetSheetName(schemaMap, "LOCATION_MAP");
//   const locSheet = locSheetName ? ss.getSheetByName(locSheetName) : null;
  
//   if (!locSheet) return true;

//   const locData = locSheet.getDataRange().getValues();
//   const colLocCode = schemaGetColIndex(schemaMap, "LOCATION_MAP", "location_code");
//   const colLocType = schemaGetColIndex(schemaMap, "LOCATION_MAP", "location_type");

//   let fromType = "";
//   let toType = "";

//   if (colLocCode !== -1 && colLocType !== -1) {
//     for (let i = 1; i < locData.length; i++) {
//       const code = cleanCodeValue_(locData[i][colLocCode]);
//       const type = cleanCodeValue_(locData[i][colLocType]).toUpperCase();
      
//       if (code === fromCode) fromType = type;
//       if (code === toCode) toType = type;
//     }
//   }

//   // 1. Loại bỏ các địa điểm thuộc nhóm Chi phí (EXPENSE) hoặc Kho ảo (VIRTUAL -> VIRTUAL)
//   if (fromType === "EXPENSE" || toType === "EXPENSE") return false;
//   if (fromType === "VIRTUAL" && toType === "VIRTUAL") return false;

//   // 2. Tra cứu chính xác theo col_key 'is_inventory' trong tbl_route_map
//   const routeSheetName = schemaGetSheetName(schemaMap, "ROUTE_MAP");
//   const routeSheet = routeSheetName ? ss.getSheetByName(routeSheetName) : null;
  
//   if (routeSheet && fromType && toType) {
//     const routeData = routeSheet.getDataRange().getValues();
//     const colFromType  = schemaGetColIndex(schemaMap, "ROUTE_MAP", "from_code");
//     const colToType    = schemaGetColIndex(schemaMap, "ROUTE_MAP", "to_code");
//     const colIsInventory = schemaGetColIndex(schemaMap, "ROUTE_MAP", "is_inventory"); // ◄--- Đã đổi thành is_inventory

//     if (colFromType !== -1 && colToType !== -1 && colIsInventory !== -1) {
//       for (let i = 1; i < routeData.length; i++) {
//         const fT = cleanCodeValue_(routeData[i][colFromType]).toUpperCase();
//         const tT = cleanCodeValue_(routeData[i][colToType]).toUpperCase();
        
//         if (fT === fromType && tT === toType) {
//           return routeData[i][colIsInventory] === true || String(routeData[i][colIsInventory]).toUpperCase() === "TRUE";
//         }
//       }
//     }
//   }

//   return true;
// }


