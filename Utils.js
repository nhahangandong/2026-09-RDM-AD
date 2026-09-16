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


/**
 * Tính kỳ liền trước theo mặt nạ YYYY-MM (Ví dụ: '2026-03' -> '2026-02')
 * @param {string} strPeriod - Kỳ hiện tại dạng YYYY-MM
 * @returns {string} Kỳ trước dạng YYYY-MM
 */
function getPreviousPeriod_(strPeriod) {
  const parts = String(strPeriod).trim().split('-');
  if (parts.length !== 2) return strPeriod;

  let year = parseInt(parts[0], 10);
  let month = parseInt(parts[1], 10);

  if (month === 1) {
    year -= 1;
    month = 12;
  } else {
    month -= 1;
  }
  return `${year}-${String(month).padStart(2, '0')}`;
}

/**
 * Parse giá trị ngày bất kỳ về chuỗi định dạng YYYY-MM
 * @param {*} rawNgay - Giá trị ngày đầu vào (Date object hoặc String)
 * @param {string} tz - Múi giờ của Spreadsheet
 * @returns {string} Chuỗi kỳ dạng YYYY-MM
 */
function parsePeriod_(rawNgay, tz) {
  if (!rawNgay) return "";

  if (rawNgay instanceof Date) {
    return Utilities.formatDate(rawNgay, tz, "yyyy-MM");
  }

  const str = String(rawNgay).trim();

  // Kiểm tra định dạng YYYY-MM
  if (/^\d{4}-\d{2}$/.test(str)) {
    return str;
  }

  // Nếu là dạng dd/MM/yyyy
  if (str.length >= 10 && str.includes('/')) {
    const p = str.split('/');
    return `${p[2].substring(0, 4)}-${p[1].padStart(2, '0')}`;
  }

  const d = new Date(str);
  return !isNaN(d.getTime()) ? Utilities.formatDate(d, tz, "yyyy-MM") : "";
}


/**
 * Hàm phụ trợ sinh danh sách các kỳ dạng YYYY-MM từ khoảng [Start, End]
 */
function generatePeriodRange_(startStr, endStr) {
  const result = [];
  const startParts = String(startStr).trim().split('-');
  const endParts = String(endStr).trim().split('-');

  if (startParts.length !== 2 || endParts.length !== 2) return [];

  let sYear = parseInt(startParts[0], 10);
  let sMonth = parseInt(startParts[1], 10);
  let eYear = parseInt(endParts[0], 10);
  let eMonth = parseInt(endParts[1], 10);

  if (isNaN(sYear) || isNaN(sMonth) || isNaN(eYear) || isNaN(eMonth)) return [];

  let currentVal = sYear * 12 + sMonth;
  let targetLimit = eYear * 12 + eMonth;

  if (currentVal > targetLimit) return [];

  while (currentVal <= targetLimit) {
    const monthStr = sMonth < 10 ? '0' + sMonth : sMonth;
    result.push(`${sYear}-${monthStr}`);

    sMonth++;
    if (sMonth > 12) {
      sMonth = 1;
      sYear++;
    }
    currentVal = sYear * 12 + sMonth;
  }

  return result;
}


/**
 * Helper tính kỳ kế tiếp YYYY-MM trong Utils hoặc cục bộ
 */
function getNextPeriod_(strPeriod) {
  const parts = String(strPeriod).trim().split('-');
  if (parts.length !== 2) return strPeriod;

  let year = parseInt(parts[0], 10);
  let month = parseInt(parts[1], 10);

  if (month === 12) {
    year += 1;
    month = 1;
  } else {
    month += 1;
  }
  return `${year}-${String(month).padStart(2, '0')}`;
}


/**
 * Helper tự động cập nhật số dư đầu kỳ cho kỳ kế tiếp (LOCATION_ITEM_BALANCE)
 * Đặt trong Utils.gs, không dùng static.
 */
function updateNextPeriodBalance_(ss, schemaMap, nextPeriod, newRows) {
  const S_BAL = "LOCATION_ITEM_BALANCE";
  const sheetName = schemaGetSheetName(schemaMap, S_BAL);
  let sBal = ss.getSheetByName(sheetName);
  if (!sBal) return; // Nếu chưa tạo sheet balance thì bỏ qua

  let dataBal = sBal.getDataRange().getValues();
  const cols = schemaMap[S_BAL].columns;
  const headers = cols.map(c => c.colHeader || c.colKey);

  if (dataBal.length === 0 || dataBal[0][0] === "") {
    dataBal = [headers];
  }

  const idxP = schemaGetColIndex(schemaMap, S_BAL, "period");

  // Lọc bỏ các dòng cũ của kỳ nextPeriod để ghi đè mới hoàn toàn số liệu chốt sổ sạch
  const filteredBal = [dataBal[0]];
  for (let i = 1; i < dataBal.length; i++) {
    const p = String(dataBal[i][idxP] || "").trim().replace(/\.0$/, '');
    if (p !== nextPeriod) {
      filteredBal.push(dataBal[i]);
    }
  }

  // Thêm các dòng số dư mới vào
  newRows.forEach(r => filteredBal.push(r));

  sBal.clearContents();
  sBal.getRange(1, 1, filteredBal.length, cols.length).setValues(filteredBal);
}




/**
 * Hàm phụ trợ chuẩn hóa giá trị kỳ về dạng YYYY-MM
 */
function formatPeriodStandard_(val) {
  if (!val) return "";
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    return `${y}-${m}`;
  }
  const str = String(val).trim();
  // Nếu là dạng 202608 -> đổi thành 2026-08
  if (/^\d{6}$/.test(str)) {
    return `${str.substring(0, 4)}-${str.substring(4, 6)}`;
  }
  return str;
}


/**
 * Private Helper: Lọc bỏ các từ dung tích/quy cách rác để gom nhóm các biến thể tên gần giống nhau
 */
function cleanSpecWords_(str) {
  if (!str) return "";
  return str.toLowerCase()
    .replace(/\b(\d+)\s*(g|kg|ml|l|gr|pax|set)\b/gi, "")
    .replace(/\b(hộp|túi|chai|thùng|bao|block|lốc|khay|can)\b/gi, "")
    .replace(/[^a-zA-Z0-9àáâãèéêìíòóôõùúăđĩũơưăạảấầẩẫậắnằẳẵặẹẻẽềềểễệỉịọỏốồổỗộớờởỡợụủứừửữựỳỵỷỹ\s]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}




/**
 * Hàm tiện ích tự động gán số thứ tự (STT) theo nhóm mã giao dịch (trans_no)
 * @param {Sheet} sheet - Đối tượng Sheet cần xử lý
 * @param {number} colIndexTransNo - Chỉ số cột chứa mã giao dịch (1-based index, VD: cột A = 1)
 * @param {number} colIndexStt - Chỉ số cột cần ghi STT nhóm (1-based index)
 */
function assignGroupSttByTransNo(sheet, colIndexTransNo, colIndexStt) {
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length <= 1) return; // Không có dữ liệu hoặc chỉ có tiêu đề

  const countMap = new Map(); // Dùng để đếm số lần xuất hiện của mỗi trans_no
  const sttColumnData = [];   // Mảng chứa giá trị STT để ghi ngược lại sheet

  // Bắt đầu duyệt từ dòng thứ 2 (bỏ qua dòng tiêu đề index = 0)
  for (let i = 1; i < values.length; i++) {
    const transNo = values[i][colIndexTransNo - 1] ? values[i][colIndexTransNo - 1].toString().trim() : "";
    
    if (transNo) {
      // Tăng số đếm cho mã giao dịch hiện tại
      const currentCount = (countMap.get(transNo) || 0) + 1;
      countMap.set(transNo, currentCount);
      sttColumnData.push([currentCount]); // Gán số thứ tự trong nhóm (1, 2, 3...)
    } else {
      sttColumnData.push([""]); // Nếu dòng trống mã giao dịch thì bỏ trống STT
    }
  }

  // Ghi dữ liệu STT hàng loạt (Batch update) vào cột STT từ dòng thứ 2 trở xuống
  if (sttColumnData.length > 0) {
    sheet.getRange(2, colIndexStt, sttColumnData.length, 1).setValues(sttColumnData);
  }
}




