/**
 * [Module EXPENSE] Đồng bộ và tổng hợp dữ liệu chi phí từ sheet TRANSACTION.
 * Có tham chiếu ROUTE_MAP và LOCATION để lọc đúng các dòng giao dịch có nhãn/loại là EXPENSE.
 * @param {string} targetPeriod Kỳ hạch toán cần xử lý (VD: '2026-08')
 */
function expenseSyncFromTransaction(targetPeriod) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi Schema", "Không thể đọc cấu trúc SCHEMA!", ui.ButtonSet.OK);
    return;
  }
  
  const transSheetName = schemaGetSheetName(schemaMap, "TRANSACTION");
  const expenseSheetName = schemaGetSheetName(schemaMap, "EXPENSE") || "EXPENSE";
  const routeMapSheetName = schemaGetSheetName(schemaMap, "ROUTE_MAP") || "ROUTE_MAP";
  const locationSheetName = schemaGetSheetName(schemaMap, "LOCATION") || "LOCATION";
  
  const transSheet = transSheetName ? ss.getSheetByName(transSheetName) : null;
  const routeSheet = ss.getSheetByName(routeMapSheetName);
  const locationSheet = ss.getSheetByName(locationSheetName);
  let expenseSheet = ss.getSheetByName(expenseSheetName);
  
  if (!transSheet) {
    ui.alert("⚠️ Lỗi", "Không tìm thấy sheet TRANSACTION theo SCHEMA!", ui.ButtonSet.OK);
    return;
  }
  
  // Nếu chưa có sheet EXPENSE thì tạo mới và tự động tạo tiêu đề từ Schema
  if (!expenseSheet) {
    expenseSheet = ss.insertSheet(expenseSheetName);
    const headers = schemaMap["EXPENSE"].columns.map(c => c.col_header);
    expenseSheet.appendRow(headers);
  }

  // 1. Xác định danh sách các điểm đích (to_type) hoặc mã tuyến thuộc loại EXPENSE từ ROUTE_MAP và LOCATION
  const expenseToTypes = new Set();
  
  if (locationSheet && schemaMap["LOCATION"]) {
    const locData = locationSheet.getDataRange().getValues();
    const idxLocType = schemaGetColIndex(schemaMap, "LOCATION", "location_type");
    const idxLocCode = schemaGetColIndex(schemaMap, "LOCATION", "location_code");

    if (idxLocType !== -1 && idxLocCode !== -1) {
      for (let i = 1; i < locData.length; i++) {
        const lType = locData[i][idxLocType] ? locData[i][idxLocType].toString().trim().toUpperCase() : "";
        const lCode = locData[i][idxLocCode] ? locData[i][idxLocCode].toString().trim().toLowerCase() : "";
        if (lType === "EXPENSE" && lCode) {
          expenseToTypes.add(lCode);
        }
      }
    }
  }

  // Fallback hoặc bổ sung check từ ROUTE_MAP nếu cần thiết (ví dụ kiểm tra to_type trong ROUTE_MAP)
  if (routeSheet && schemaMap["ROUTE_MAP"]) {
    const routeData = routeSheet.getDataRange().getValues();
    const idxRouteToType = schemaGetColIndex(schemaMap, "ROUTE_MAP", "to_type");
    
    if (idxRouteToType !== -1) {
      for (let i = 1; i < routeData.length; i++) {
        const rToType = routeData[i][idxRouteToType] ? routeData[i][idxRouteToType].toString().trim().toLowerCase() : "";
        if (rToType) {
          expenseToTypes.add(rToType);
        }
      }
    }
  }

  // 2. Lấy vị trí cột bên TRANSACTION
  const idxTransPeriod   = schemaGetColIndex(schemaMap, "TRANSACTION", "period");
  const idxTransToType   = schemaGetColIndex(schemaMap, "TRANSACTION", "to_type");   // Hoặc to_code tùy cấu trúc TRANSACTION của bạn
  const idxTransToCode   = schemaGetColIndex(schemaMap, "TRANSACTION", "to_code");
  const idxTransRawName  = schemaGetColIndex(schemaMap, "TRANSACTION", "raw_name");
  const idxTransItemCode = schemaGetColIndex(schemaMap, "TRANSACTION", "item_code");
  const idxTransUnit     = schemaGetColIndex(schemaMap, "TRANSACTION", "input_unit");
  const idxTransQty      = schemaGetColIndex(schemaMap, "TRANSACTION", "quantity");
  const idxTransTotalAmt = schemaGetColIndex(schemaMap, "TRANSACTION", "total_amount");
  const idxTransTaxRate  = schemaGetColIndex(schemaMap, "TRANSACTION", "tax_rate");
  const idxTransNetAmt   = schemaGetColIndex(schemaMap, "TRANSACTION", "net_amount");

  if (idxTransPeriod === -1 || idxTransToCode === -1 || idxTransItemCode === -1) {
    ui.alert("⚠️ Lỗi Schema", "Sheet TRANSACTION thiếu các cột bắt buộc (period, to_code, item_code)!", ui.ButtonSet.OK);
    return;
  }

  const transData = transSheet.getDataRange().getValues();
  if (transData.length <= 1) {
    ui.alert("ℹ️ Thông báo", "Sheet TRANSACTION không có dữ liệu.", ui.ButtonSet.OK);
    return;
  }

  // 3. Duyệt và lọc dữ liệu TRANSACTION chỉ lấy dòng EXPENSE dựa trên ROUTE_MAP / LOCATION
  const expenseMap = new Map();
  
  for (let i = 1; i < transData.length; i++) {
    const row = transData[i];
    const period = row[idxTransPeriod] ? row[idxTransPeriod].toString().trim() : "";
    
    // Lọc theo kỳ nếu có chỉ định
    if (targetPeriod && period !== targetPeriod) continue;
    
    // Kiểm tra xem dòng này có phải là EXPENSE dựa vào to_type / to_code khớp với ROUTE_MAP / LOCATION không
    const toTypeVal = idxTransToType !== -1 ? row[idxTransToType].toString().trim().toLowerCase() : "";
    const toCodeVal = row[idxTransToCode] ? row[idxTransToCode].toString().trim().toLowerCase() : "";
    
    const isExpenseRow = expenseToTypes.has(toTypeVal) || expenseToTypes.has(toCodeVal) || toTypeVal.includes("expense") || toCodeVal.includes("cp_");
    
    if (!isExpenseRow) continue; // Bỏ qua nếu không phải dòng chi phí

    const deptCode = row[idxTransToCode] ? row[idxTransToCode].toString().trim() : "";
    const itemCode = row[idxTransItemCode] ? row[idxTransItemCode].toString().trim() : "";
    const rawName  = idxTransRawName !== -1 ? row[idxTransRawName] : "";
    const buyUnit  = idxTransUnit !== -1 ? row[idxTransUnit] : "";
    
    const qty      = idxTransQty !== -1 ? (Number(row[idxTransQty]) || 0) : 0;
    const totalAmt = idxTransTotalAmt !== -1 ? (Number(row[idxTransTotalAmt]) || 0) : 0;
    const netAmt   = idxTransNetAmt !== -1 ? (Number(row[idxTransNetAmt]) || 0) : 0;
    const taxRate  = idxTransTaxRate !== -1 ? row[idxTransTaxRate] : 0;
    
    // Tạo khóa duy nhất để gom nhóm
    const key = `${period}|${deptCode}|${itemCode}|${buyUnit}`;
    
    if (expenseMap.has(key)) {
      const existing = expenseMap.get(key);
      existing.buy_qty += qty;
      existing.total_amt += totalAmt;
      existing.net_amt += netAmt;
    } else {
      expenseMap.set(key, {
        period: period,
        department_code: deptCode,
        raw_name: rawName,
        item_code: itemCode,
        buy_unit: buyUnit,
        buy_qty: qty,
        total_amt: totalAmt,
        tax_rate: taxRate,
        net_amt: netAmt
      });
    }
  }

  // 4. Chuẩn bị mảng ghi ra sheet EXPENSE
  const maxColCount = schemaMap["EXPENSE"].columns.length;
  const rowsToInsert = [];
  
  expenseMap.forEach((val) => {
    const expRow = new Array(maxColCount).fill("");
    
    expRow[schemaGetColIndex(schemaMap, "EXPENSE", "period")] = val.period;
    expRow[schemaGetColIndex(schemaMap, "EXPENSE", "department_code")] = val.department_code;
    expRow[schemaGetColIndex(schemaMap, "EXPENSE", "raw_name")] = val.raw_name;
    expRow[schemaGetColIndex(schemaMap, "EXPENSE", "item_code")] = val.item_code;
    expRow[schemaGetColIndex(schemaMap, "EXPENSE", "buy_unit")] = val.buy_unit;
    expRow[schemaGetColIndex(schemaMap, "EXPENSE", "buy_qty")] = Math.round(val.buy_qty * 1000) / 1000;
    expRow[schemaMap, "EXPENSE", "total_amt"] = Math.round(val.total_amt * 100) / 100; // Sửa index an toàn bên dưới
    expRow[schemaGetColIndex(schemaMap, "EXPENSE", "total_amt")] = Math.round(val.total_amt * 100) / 100;
    expRow[schemaGetColIndex(schemaMap, "EXPENSE", "tax_rate")] = val.tax_rate;
    expRow[schemaGetColIndex(schemaMap, "EXPENSE", "net_amt")] = Math.round(val.net_amt * 100) / 100;
    
    rowsToInsert.push(expRow);
  });

  // 5. Ghi dữ liệu hàng loạt xuống sheet EXPENSE
  if (rowsToInsert.length > 0) {
    if (expenseSheet.getLastRow() > 1) {
      expenseSheet.getRange(2, 1, expenseSheet.getLastRow() - 1, maxColCount).clearContent();
    }
    expenseSheet.getRange(2, 1, rowsToInsert.length, maxColCount).setValues(rowsToInsert);
  }

  ui.alert(
    "✅ Thành công", 
    `Đã lọc và tổng hợp thành công ${rowsToInsert.length} dòng chi phí (EXPENSE) theo ROUTE_MAP cho kỳ [${targetPeriod || "Tất cả"}]!`, 
    ui.ButtonSet.OK
  );
}
