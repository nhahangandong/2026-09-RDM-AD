/**
 * ============================================================================
 * MODULE: Transaction.gs
 * MỤC ĐÍCH: Xử lý logic nghiệp vụ cho các giao dịch TRANSACTION.
 * Quy tắc đặt tên: transactionActionEntity (CamelCase)
 * ============================================================================
 */

/**
 * Core: Tra cứu và điền mã SKU (item_code) từ sheet MAPPING quay trở lại các sheet giao dịch.
 * Khóa tra cứu Composite Key: raw_name + source_group
 */
function transactionAssignItemCodeToAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. LẤY CẤU TRÚC SCHEMA
  const schemaMap = schemaGetMap();
  if (!schemaMap) {
    ui.alert("⚠️ Lỗi", "Không thể đọc cấu trúc từ sheet 'SCHEMA'!", ui.ButtonSet.OK);
    return;
  }

  const mappingSheetName = schemaGetSheetName(schemaMap, "MAPPING");
  const mappingSheet = mappingSheetName ? ss.getSheetByName(mappingSheetName) : null;
  if (!mappingSheet) {
    ui.alert("⚠️ Thông báo", "Không tìm thấy sheet MAPPING!", ui.ButtonSet.OK);
    return;
  }

  const idxMappingRaw    = schemaGetColIndex(schemaMap, "MAPPING", "raw_name");
  const idxMappingSource = schemaGetColIndex(schemaMap, "MAPPING", "source");
  const idxMappingCode   = schemaGetColIndex(schemaMap, "MAPPING", "item_code");

  if (idxMappingRaw === -1 || idxMappingCode === -1) {
    ui.alert("⚠️ Lỗi SCHEMA", "Sheet MAPPING thiếu col_key 'raw_name' hoặc 'item_code'!", ui.ButtonSet.OK);
    return;
  }

  // 2. LẤY MAP ÁNH XẠ NHÓM NGUỒN TỪ TBL_RAW_SOURCE_GROUP
  const dynamicSourceGroupMap = getSourceGroupMap_(schemaMap, ss);

  // 3. DỰNG MAP TRA CỨU MÃ SKU THEO COMPOSITE KEY (raw_name + source_group)
  const mappingData = mappingSheet.getDataRange().getValues();
  const codeLookupMap = new Map();

  for (let i = 1; i < mappingData.length; i++) {
    const rawVal = mappingData[i][idxMappingRaw] ? mappingData[i][idxMappingRaw].toString().trim().toLowerCase() : "";
    const sourceVal = idxMappingSource !== -1 && mappingData[i][idxMappingSource] ? mappingData[i][idxMappingSource].toString().trim().toLowerCase() : "";
    const codeVal = mappingData[i][idxMappingCode] ? mappingData[i][idxMappingCode].toString().trim() : "";

    if (rawVal && codeVal) {
      const sourceGroup = dynamicSourceGroupMap.get(sourceVal) || sourceVal;
      const compositeKey = `${rawVal}|${sourceGroup}`;
      codeLookupMap.set(compositeKey, codeVal);
    }
  }

  if (codeLookupMap.size === 0) {
    ui.alert("ℹ️ Thông báo", "Sheet MAPPING chưa có dữ liệu Mã SKU hợp lệ!", ui.ButtonSet.OK);
    return;
  }

  let totalUpdatedSheets = 0;
  let totalUpdatedRows = 0;
  
  // [CẢI TIẾN] Mảng lưu chi tiết kết quả từng sheet để thông báo rõ ràng
  const detailReport = [];

  // 4. DUYỆT TẤT CẢ CÁC SHEET GIAO DỊCH RAW
  Object.keys(schemaMap).forEach(schemaName => {
    if (schemaName === "MAPPING" || schemaName === "SCHEMA" || schemaName === "RAW_SOURCE_GROUP" || schemaName === "ITEM_MASTER") return;

    // CHỈ CHO PHÉP 3 NGUỒN GIAO DỊCH CHÍNH THEO YÊU CẦU TRƯỚC ĐÓ
    const schemaUpper = schemaName.toUpperCase();
    if (schemaUpper !== "TRANSACTION" && schemaUpper !== "STOCKTAKE" && schemaUpper !== "STOCK_TAKE" && schemaUpper !== "SALES") {
      return;
    }

    const idxSourceRaw  = schemaGetColIndex(schemaMap, schemaName, "raw_name");
    const idxSourceCode = schemaGetColIndex(schemaMap, schemaName, "item_code");

    if (idxSourceRaw === -1 || idxSourceCode === -1) return;

    const targetSheetName = schemaGetSheetName(schemaMap, schemaName);
    const targetSheet = targetSheetName ? ss.getSheetByName(targetSheetName) : null;
    if (!targetSheet) return;

    const lastRow = targetSheet.getLastRow();
    if (lastRow <= 1) return;

    const rangeRaw = targetSheet.getRange(2, idxSourceRaw + 1, lastRow - 1, 1).getValues();
    const rangeCode = targetSheet.getRange(2, idxSourceCode + 1, lastRow - 1, 1).getValues();

    const currentSchemaLower = schemaName.toLowerCase();
    const currentSheetLower  = targetSheetName.toLowerCase();
    const currentSourceGroup = dynamicSourceGroupMap.get(currentSchemaLower) || dynamicSourceGroupMap.get(currentSheetLower) || currentSchemaLower;

    let sheetUpdatedCount = 0;

    for (let i = 0; i < rangeRaw.length; i++) {
      const rawVal = rangeRaw[i][0] ? rangeRaw[i][0].toString().trim().toLowerCase() : "";
      if (!rawVal) continue;

      const targetCompositeKey = `${rawVal}|${currentSourceGroup}`;

      if (codeLookupMap.has(targetCompositeKey)) {
        const matchedCode = codeLookupMap.get(targetCompositeKey);
        const currentCode = rangeCode[i][0] ? rangeCode[i][0].toString().trim() : "";

        if (currentCode !== matchedCode) {
          rangeCode[i][0] = matchedCode;
          sheetUpdatedCount++;
        }
      }
    }

    // 5. GHI DỮ LIỆU CỘT ITEM_CODE XUỐNG SHEET
    if (sheetUpdatedCount > 0) {
      const colCode1Based = idxSourceCode + 1;
      targetSheet.getRange(2, colCode1Based, rangeCode.length, 1).setValues(rangeCode);

      totalUpdatedSheets++;
      totalUpdatedRows += sheetUpdatedCount;
      
      // Ghi nhận chi tiết cho từng sheet
      detailReport.py ? null : detailReport.push(`• Sheet [${targetSheetName}] (${schemaName}): Cập nhật ${sheetUpdatedCount} dòng.`);
      // Hoặc viết gọn:
      detailReport.push(`• Sheet "${targetSheetName}": ${sheetUpdatedCount} dòng`);
    }
  });

  // 6. THÔNG BÁO KẾT QUẢ CHI TIẾT VÀ RÕ RÀNG
  if (totalUpdatedRows > 0) {
    let message = `✅ Đã cập nhật tổng cộng ${totalUpdatedRows} dòng mã SKU trên ${totalUpdatedSheets} sheet giao dịch:\n\n`;
    message += detailReport.join("\n");
    
    ui.alert("Kết quả cập nhật mã SKU", message, ui.ButtonSet.OK);
  } else {
    ui.alert(
      "ℹ️ Thông báo",
      "Dữ liệu mã SKU trên các sheet giao dịch đã khớp hoàn toàn, không có dòng nào cần cập nhật mới.",
      ui.ButtonSet.OK
    );
  }
}


/**
 * Hàm hỗ trợ tự động lấy hoặc tính toán Period (YYYYMM)
 * @param {Date|string} transDate - Ngày chứng từ
 * @param {string|number} manualPeriod - Giá trị kỳ người dùng tự gõ (nếu có)
 * @returns {string} Kỳ hạch toán dạng "YYYYMM"
 */
function resolvePeriod_(transDate, manualPeriod) {
  if (manualPeriod && manualPeriod.toString().trim() !== '') {
    return manualPeriod.toString().trim();
  }
  
  if (transDate instanceof Date && !isNaN(transDate.getTime())) {
    const year = transDate.getFullYear();
    const month = String(transDate.getMonth() + 1).padStart(2, '0');
    return `${year}${month}`;
  }
  
  return '';
}

/**
 * Tạo chuỗi Composite PK từ trans_no, line_no, item_code
 * @param {string} transNo - Số chứng từ
 * @param {number|string} lineNo - STT dòng
 * @param {string} itemCode - Mã SKU
 * @returns {string} Chuỗi PK duy nhất
 */
function generateCompositeKey_(transNo, lineNo, itemCode) {
  const cleanTransNo = transNo ? transNo.toString().trim() : '';
  const cleanLine = lineNo ? lineNo.toString().trim() : '0';
  const cleanCode = itemCode ? itemCode.toString().trim() : '';
  
  return `${cleanTransNo}_${cleanLine}_${cleanCode}`;
}


/**
 * ============================================================================
 * MODULE: TransactionEngine.gs
 * MỤC ĐÍCH: Tra cứu ITEM_MASTER, LOCATION_MAP, ROUTE_MAP để tự động tính toán
 *           và gán Category, ĐVT, Kỳ HT, Số lượng/Đơn giá chuẩn cho các sheet.
 * ============================================================================
 */

function transactionFillCalculatedColumnsAllSheets() {
  // CỜ CẤU HÌNH CỤC BỘ: true = cho phép cập nhật/đè lại cột Category sau mỗi lần chạy.
  // Chuyển sang false khi hệ thống đã chuẩn hóa xong nhóm dữ liệu.
  const ENABLE_CATEGORY_UPDATE = true;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schemaMap = schemaGetMap();
  if (!schemaMap) throw new Error("Không thể tải cấu trúc Schema Map.");
  const tz = ss.getSpreadsheetTimeZone() || "GMT+7";

  // --------------------------------------------------------------------------
  // 1. NẠP ITEM_MASTER (item_code -> {base_unit, category})
  // --------------------------------------------------------------------------
  const SCHEMA_ITEM = "ITEM_MASTER";
  const sheetNameItem = schemaGetSheetName(schemaMap, SCHEMA_ITEM);
  const sItem = ss.getSheetByName(sheetNameItem);
  const mapItemMaster = {};

  if (sItem) {
    const dataItem = sItem.getDataRange().getValues();
    const idxItemCode = schemaGetColIndex(schemaMap, SCHEMA_ITEM, "item_code");
    const idxBaseUnit = schemaGetColIndex(schemaMap, SCHEMA_ITEM, "base_unit");
    const idxCategory = schemaGetColIndex(schemaMap, SCHEMA_ITEM, "category");

    for (let i = 1; i < dataItem.length; i++) {
      const row = dataItem[i];
      const code = cleanCodeValue_(row[idxItemCode]);
      if (code) {
        mapItemMaster[code] = {
          base_unit: String(row[idxBaseUnit] || "").trim(),
          category: String(row[idxCategory] || "").trim()
        };
      }
    }
  }

  // --------------------------------------------------------------------------
  // 2. NẠP LOCATION_MAP (location_code -> location_type)
  // --------------------------------------------------------------------------
  const SCHEMA_LOC = "LOCATION_MAP";
  const sheetNameLoc = schemaGetSheetName(schemaMap, SCHEMA_LOC);
  const sLoc = ss.getSheetByName(sheetNameLoc);
  const mapLocation = {};

  if (sLoc) {
    const dataLoc = sLoc.getDataRange().getValues();
    const idxLocCode = schemaGetColIndex(schemaMap, SCHEMA_LOC, "location_code");
    const idxLocType = schemaGetColIndex(schemaMap, SCHEMA_LOC, "location_type");

    for (let i = 1; i < dataLoc.length; i++) {
      const row = dataLoc[i];
      const code = idxLocCode !== -1 ? cleanCodeValue_(row[idxLocCode]) : "";
      const locType = idxLocType !== -1 ? String(row[idxLocType] || "").trim().toUpperCase() : "";
      if (code) {
        mapLocation[code] = { type: locType };
      }
    }
  }

  // --------------------------------------------------------------------------
  // 3. NẠP ROUTE_MAP (from_code + to_code -> trans_type)
  // --------------------------------------------------------------------------
  const SCHEMA_ROUTE = "ROUTE_MAP";
  const sheetNameRoute = schemaGetSheetName(schemaMap, SCHEMA_ROUTE);
  const sRoute = ss.getSheetByName(sheetNameRoute);
  const mapRoute = {};

  if (sRoute) {
    const dataRoute = sRoute.getDataRange().getValues();
    const idxRouteFrom = schemaGetColIndex(schemaMap, SCHEMA_ROUTE, "from_code");
    const idxRouteTo = schemaGetColIndex(schemaMap, SCHEMA_ROUTE, "to_code");
    const idxTransType = schemaGetColIndex(schemaMap, SCHEMA_ROUTE, "trans_type");

    for (let i = 1; i < dataRoute.length; i++) {
      const row = dataRoute[i];
      const fCode = idxRouteFrom !== -1 ? cleanCodeValue_(row[idxRouteFrom]) : "";
      const tCode = idxRouteTo !== -1 ? cleanCodeValue_(row[idxRouteTo]) : "";
      const tType = idxTransType !== -1 ? String(row[idxTransType] || "").trim().toUpperCase() : "";
      
      if (fCode && tCode) {
        mapRoute[`${fCode}_${tCode}`] = tType;
      }
    }
  }

  // --------------------------------------------------------------------------
  // 4. XỬ LÝ SHEET TRANSACTION
  // --------------------------------------------------------------------------
  const SCHEMA_TRANS = "TRANSACTION";
  const sheetNameTrans = schemaGetSheetName(schemaMap, SCHEMA_TRANS);
  const sTrans = ss.getSheetByName(sheetNameTrans);

  if (sTrans) {
    const dataTrans = sTrans.getDataRange().getValues();
    if (dataTrans.length > 1) {
      const headers = dataTrans[0].map(h => String(h).trim().toLowerCase());
      const getIdxT = (colKey) => schemaGetColIndex(schemaMap, SCHEMA_TRANS, colKey);

      let idxDate = getIdxT("trans_date");
      if (idxDate === -1) idxDate = getIdxT("date");
      if (idxDate === -1) idxDate = headers.findIndex(h => h.includes("ngày"));

      let idxPeriod = getIdxT("period");
      if (idxPeriod === -1) idxPeriod = headers.findIndex(h => h.includes("kỳ"));

      let idxFromCode = getIdxT("from_code");
      if (idxFromCode === -1) idxFromCode = headers.findIndex(h => h.includes("nơi xuất") || h.includes("nguồn"));

      let idxToCode = getIdxT("to_code");
      if (idxToCode === -1) idxToCode = headers.findIndex(h => h.includes("nơi nhận") || h.includes("đích"));

      const idxItemCode = getIdxT("item_code");
      const idxQty = getIdxT("quantity");
      const idxFactor = getIdxT("conversion_factor") !== -1 ? getIdxT("conversion_factor") : getIdxT("factor");
      const idxTotalAmt = getIdxT("total_amount");
      const idxTaxRate = getIdxT("tax_rate");
      const idxUnit = getIdxT("input_unit") !== -1 ? getIdxT("input_unit") : getIdxT("unit");

      const idxBaseQty = getIdxT("base_quantity");
      const idxBaseUnit = getIdxT("base_unit");
      const idxNetAmt = getIdxT("net_amount");
      const idxTaxAmt = getIdxT("tax_amount");
      const idxUnitPrice = getIdxT("unit_price");
      const idxNetUnitPrice = getIdxT("net_unit_price");

      let idxCat = getIdxT("category");
      if (idxCat === -1) idxCat = headers.findIndex(h => h.includes("nhóm") || h.includes("phân loại"));

      for (let i = 1; i < dataTrans.length; i++) {
        const row = dataTrans[i];
        const rawDate = idxDate !== -1 ? row[idxDate] : null;

        // Điền Kỳ HT (YYYY-MM)
        if (idxPeriod !== -1 && rawDate) {
          row[idxPeriod] = parsePeriod_(rawDate, tz);
        }

        const fromCode = idxFromCode !== -1 ? cleanCodeValue_(row[idxFromCode]) : "";
        const toCode = idxToCode !== -1 ? cleanCodeValue_(row[idxToCode]) : "";
        const itemCode = idxItemCode !== -1 ? cleanCodeValue_(row[idxItemCode]) : "";
        
        const qty = idxQty !== -1 ? (Number(row[idxQty]) || 0) : 0;
        const rawFactor = idxFactor !== -1 ? row[idxFactor] : null;
        const factor = (rawFactor !== "" && rawFactor !== null && !isNaN(rawFactor) && Number(rawFactor) > 0) ? Number(rawFactor) : 1;
        const totalAmt = idxTotalAmt !== -1 ? (Number(row[idxTotalAmt]) || 0) : 0;
        const taxRate = idxTaxRate !== -1 ? (Number(row[idxTaxRate]) || 0) : 0;
        const rawUnit = idxUnit !== -1 ? String(row[idxUnit] || "").trim() : "";

        const master = mapItemMaster[itemCode];
        const toLocInfo = mapLocation[toCode];
        const routeTransType = mapRoute[`${fromCode}_${toCode}`] || "";

        // Gán ĐVT quy đổi chuẩn
        if (idxBaseUnit !== -1) {
          row[idxBaseUnit] = (master && master.base_unit) ? master.base_unit : rawUnit;
        }

        // --- CƠ CHẾ GÁN CATEGORY CHO TRANSACTION ---
        if (idxCat !== -1) {
          const currentCat = String(row[idxCat] || "").trim();

          if (ENABLE_CATEGORY_UPDATE || !currentCat) {
            // Kiểm tra Tuyến Chi phí (Nơi nhận CP_*, location_type = EXPENSE, hoặc trans_type = EXPENSE)
            const isExpenseRoute = toCode.startsWith("CP_") || 
                                   (toLocInfo && toLocInfo.type === "EXPENSE") || 
                                   (routeTransType === "EXPENSE");

            if (isExpenseRoute) {
              row[idxCat] = "CHI PHÍ / DỊCH VỤ";
            } else if (master && master.category) {
              row[idxCat] = master.category;
            } else {
              row[idxCat] = "Chưa phân nhóm";
            }
          }
        }

        // Tính toán các thông số tài chính & số lượng quy đổi
        if (idxBaseQty !== -1) row[idxBaseQty] = qty * factor;

        const netAmt = taxRate > 0 ? totalAmt / (1 + taxRate) : totalAmt;
        const taxAmt = totalAmt - netAmt;
        if (idxNetAmt !== -1) row[idxNetAmt] = netAmt;
        if (idxTaxAmt !== -1) row[idxTaxAmt] = taxAmt;

        if (idxUnitPrice !== -1) row[idxUnitPrice] = qty > 0 ? totalAmt / qty : 0;
        if (idxNetUnitPrice !== -1) row[idxNetUnitPrice] = qty > 0 ? netAmt / qty : 0;
      }

      sTrans.getRange(1, 1, dataTrans.length, dataTrans[0].length).setValues(dataTrans);
    }
  }

  // --------------------------------------------------------------------------
  // 5. XỬ LÝ SHEET STOCKTAKE (Tồn kho thuần túy)
  // --------------------------------------------------------------------------
  const SCHEMA_STOCK = "STOCKTAKE";
  const sheetNameStock = schemaGetSheetName(schemaMap, SCHEMA_STOCK);
  const sStock = ss.getSheetByName(sheetNameStock);

  if (sStock) {
    const dataStock = sStock.getDataRange().getValues();
    if (dataStock.length > 1) {
      const headers = dataStock[0].map(h => String(h).trim().toLowerCase());
      const getIdxS = (colKey) => schemaGetColIndex(schemaMap, SCHEMA_STOCK, colKey);

      let idxDate = getIdxS("trans_date");
      if (idxDate === -1) idxDate = getIdxS("date");
      if (idxDate === -1) idxDate = headers.findIndex(h => h.includes("ngày"));

      let idxPeriod = getIdxS("period");
      if (idxPeriod === -1) idxPeriod = headers.findIndex(h => h.includes("kỳ"));

      const idxItemCode = getIdxS("item_code");
      const idxQty = getIdxS("quantity");
      const idxFactor = getIdxS("conversion_factor") !== -1 ? getIdxS("conversion_factor") : getIdxS("factor");
      const idxUnit = getIdxS("input_unit") !== -1 ? getIdxS("input_unit") : getIdxS("unit");

      const idxBaseQty = getIdxS("base_quantity");
      const idxBaseUnit = getIdxS("base_unit");

      let idxCat = getIdxS("category");
      if (idxCat === -1) idxCat = headers.findIndex(h => h.includes("nhóm") || h.includes("phân loại"));

      for (let i = 1; i < dataStock.length; i++) {
        const row = dataStock[i];
        const rawDate = idxDate !== -1 ? row[idxDate] : null;

        if (idxPeriod !== -1 && rawDate) {
          row[idxPeriod] = parsePeriod_(rawDate, tz);
        }

        const itemCode = idxItemCode !== -1 ? cleanCodeValue_(row[idxItemCode]) : "";
        const qty = idxQty !== -1 ? (Number(row[idxQty]) || 0) : 0;
        const rawFactor = idxFactor !== -1 ? row[idxFactor] : null;
        const factor = (rawFactor !== "" && rawFactor !== null && !isNaN(rawFactor) && Number(rawFactor) > 0) ? Number(rawFactor) : 1;
        const rawUnit = idxUnit !== -1 ? String(row[idxUnit] || "").trim() : "";

        const master = mapItemMaster[itemCode];

        if (idxBaseUnit !== -1) {
          row[idxBaseUnit] = (master && master.base_unit) ? master.base_unit : rawUnit;
        }

        // Category cho Stocktake thuần túy theo Master Data
        if (idxCat !== -1) {
          const currentCat = String(row[idxCat] || "").trim();
          if (ENABLE_CATEGORY_UPDATE || !currentCat) {
            row[idxCat] = (master && master.category) ? master.category : "Chưa phân nhóm";
          }
        }

        if (idxBaseQty !== -1) row[idxBaseQty] = qty * factor;
      }

      sStock.getRange(1, 1, dataStock.length, dataStock[0].length).setValues(dataStock);
    }
  }
}
