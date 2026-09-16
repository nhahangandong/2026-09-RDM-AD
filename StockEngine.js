/**
 * Module StockEngine: Tính toán Tổng hợp Tồn kho (Số lượng & Giá trị) chuẩn Schema-driven.
 * Hỗ trợ cơ chế Waterfall Pricing (Lùi kỳ lịch sử và Fallback Cost Price) và chạy Batch không ngắt quãng.
 */
class StockEngine {

  /**
   * Tính toán và Upsert dữ liệu tổng hợp tồn kho theo Kỳ (YYYY-MM) - Số lượng
   * @param {string|number} periodTarget - Kỳ tính toán (Ví dụ: '2026-07', '2026-08')
   */
  static calculateQtySummary(periodTarget) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();

    const schemaMap = schemaGetMap();
    if (!schemaMap) throw new Error("Không thể tải cấu trúc SCHEMA từ hệ thống.");

    const locationTypeMap = RouteEngine.getLocationTypeMap(ss, schemaMap);
    const routeRulesMap   = RouteEngine.getRouteRulesMap(ss, schemaMap);

    const S_IQ    = "INVENTORY_QTY_SUMMARY";
    const S_TRANS = "TRANSACTION";
    const S_STOCK = "STOCKTAKE";
    const S_ITEM  = "ITEM_MASTER";
    const S_BAL   = "LOCATION_ITEM_BALANCE";

    const sIQ    = ss.getSheetByName(schemaGetSheetName(schemaMap, S_IQ));
    const sGD    = ss.getSheetByName(schemaGetSheetName(schemaMap, S_TRANS));
    const sKK    = ss.getSheetByName(schemaGetSheetName(schemaMap, S_STOCK));
    const sItem  = ss.getSheetByName(schemaGetSheetName(schemaMap, S_ITEM));
    const sBal   = ss.getSheetByName(schemaGetSheetName(schemaMap, S_BAL));

    if (!sIQ) throw new Error(`Không tìm thấy sheet ứng với schema_name: ${S_IQ}`);

    const dataIQ   = sIQ ? sIQ.getDataRange().getValues() : [];
    const dataGD   = sGD ? sGD.getDataRange().getValues() : [];
    const dataKK   = sKK ? sKK.getDataRange().getValues() : [];
    const dataItem = sItem ? sItem.getDataRange().getValues() : [];
    const dataBal  = sBal ? sBal.getDataRange().getValues() : [];

    const strPeriod = String(periodTarget).trim();
    const getIdx = (sName, cKey) => schemaGetColIndex(schemaMap, sName, cKey);

    const idxIQPeriod = getIdx(S_IQ, "period");
    const idxIQLoc    = getIdx(S_IQ, "location_code");
    const idxIQItem   = getIdx(S_IQ, "item_code");

    // 1. TẢI DANH MỤC ITEM_MASTER
    const mapItemMaster = {};
    if (dataItem.length > 1) {
      const idxItemCode = getIdx(S_ITEM, "item_code");
      const idxItemName = getIdx(S_ITEM, "item_name");
      const idxItemCat  = getIdx(S_ITEM, "category");
      const idxItemUnit = getIdx(S_ITEM, "base_unit");

      for (let i = 1; i < dataItem.length; i++) {
        const row = dataItem[i];
        const code = cleanCodeValue_(row[idxItemCode]);
        if (code) {
          mapItemMaster[code] = {
            name: idxItemName !== -1 ? String(row[idxItemName] || "").trim() : "",
            category: idxItemCat !== -1 ? String(row[idxItemCat] || "").trim() : "",
            unit: idxItemUnit !== -1 ? String(row[idxItemUnit] || "").trim() : ""
          };
        }
      }
    }

    const updateMasterFallback = (code, rawName, baseUnit, category) => {
      if (!mapItemMaster[code]) mapItemMaster[code] = { name: "", unit: "", category: "" };
      if (!mapItemMaster[code].name && rawName) mapItemMaster[code].name = rawName;
      if (!mapItemMaster[code].unit && baseUnit) mapItemMaster[code].unit = baseUnit;
      if (!mapItemMaster[code].category && category) mapItemMaster[code].category = category;
    };

    // 2. NẠP TỒN ĐẦU KỲ
    const mapOpeningQty = {};
    const setKeys = new Set(); 

    const [yStr, mStr] = strPeriod.split('-');
    let prevYear = parseInt(yStr, 10);
    let prevMonth = parseInt(mStr, 10) - 1;
    if (prevMonth === 0) {
      prevMonth = 12;
      prevYear -= 1;
    }
    const prevPeriod = `${prevYear}-${String(prevMonth).padStart(2, '0')}`;

    if (dataIQ.length > 1) {
      const idxCloseQty = getIdx(S_IQ, "closing_qty");
      for (let i = 1; i < dataIQ.length; i++) {
        const row = dataIQ[i];
        const p = String(row[idxIQPeriod] || "").trim().replace(/\.0$/, '');
        if (p === prevPeriod) {
          const loc = cleanCodeValue_(row[idxIQLoc]);
          const item = cleanCodeValue_(row[idxIQItem]);
          const qty = Number(row[idxCloseQty]) || 0;

          if (qty > 0 && loc && item) {
            const key = `${loc}:${item}`;
            mapOpeningQty[key] = qty;
            setKeys.add(key);
          }
        }
      }
    }

    if (setKeys.size === 0 && dataBal.length > 1) {
      const idxBalPeriod = getIdx(S_BAL, "period");
      const idxBalLoc    = getIdx(S_BAL, "location_code");
      const idxBalItem   = getIdx(S_BAL, "item_code");
      const idxBalQty    = getIdx(S_BAL, "opening_qty");

      for (let i = 1; i < dataBal.length; i++) {
        const row = dataBal[i];
        const p = String(row[idxBalPeriod] || "").trim().replace(/\.0$/, '');
        const loc = cleanCodeValue_(row[idxBalLoc]);
        const item = cleanCodeValue_(row[idxBalItem]);

        if (p === strPeriod && loc && item) {
          const key = `${loc}:${item}`;
          const qty = Number(row[idxBalQty]) || 0;

          if (qty > 0) {
            mapOpeningQty[key] = qty;
            setKeys.add(key);
          }
        }
      }
    }

    // 3. TẬP HỢP GIAO DỊCH TRONG KỲ (TRANSACTION)
    const mapInbound = {};
    const mapAdjust  = {};

    const idxTransPeriod   = getIdx(S_TRANS, "period");
    const idxTransItem     = getIdx(S_TRANS, "item_code");
    const idxTransToCode   = getIdx(S_TRANS, "to_code");
    const idxTransFromCode = getIdx(S_TRANS, "from_code");
    const idxTransQty      = getIdx(S_TRANS, "base_quantity");
    const idxTransRawName  = getIdx(S_TRANS, "raw_name");
    const idxTransBaseUnit = getIdx(S_TRANS, "base_unit");

    for (let i = 1; i < dataGD.length; i++) {
      const row = dataGD[i];
      const transPeriod = String(row[idxTransPeriod] || "").trim().replace(/\.0$/, '');
      if (transPeriod !== strPeriod) continue;

      const itemCode = cleanCodeValue_(row[idxTransItem]);
      const fromCode = cleanCodeValue_(row[idxTransFromCode]);
      const toCode   = cleanCodeValue_(row[idxTransToCode]);
      const qty      = Number(row[idxTransQty]) || 0;

      if (!itemCode || qty === 0) continue;

      updateMasterFallback(itemCode, String(row[idxTransRawName] || "").trim(), String(row[idxTransBaseUnit] || "").trim(), null);

      const fromType  = locationTypeMap.get(fromCode) || "UNKNOWN";
      const toType    = locationTypeMap.get(toCode) || "UNKNOWN";
      const routeRule = routeRulesMap.get(`${fromType}->${toType}`);

      if (!routeRule || !routeRule.isInventory) continue;

      if (toType === "PHYSICAL") {
        const keyTo = `${toCode}:${itemCode}`;
        setKeys.add(keyTo);

        if (routeRule.transactionType === "PURCHASE" || routeRule.transactionType === "INBOUND") {
          mapInbound[keyTo] = (mapInbound[keyTo] || 0) + qty;
        } else if (routeRule.transactionType === "TRANSFER") {
          mapAdjust[keyTo] = (mapAdjust[keyTo] || 0) + qty;
        }
      }

      if (fromType === "PHYSICAL") {
        const keyFrom = `${fromCode}:${itemCode}`;
        setKeys.add(keyFrom);
        if (routeRule.transactionType === "TRANSFER") {
          mapAdjust[keyFrom] = (mapAdjust[keyFrom] || 0) - qty;
        }
      }
    }

    // 4. TẬP HỢP KIỂM KÊ TỪ STOCKTAKE
    const mapClosing = {};
    const hasStocktakeRecord = new Set();

    const idxStockPeriod  = getIdx(S_STOCK, "period");
    const idxStockLoc     = getIdx(S_STOCK, "location");
    const idxStockItem    = getIdx(S_STOCK, "item_code");
    const idxStockBaseQty = getIdx(S_STOCK, "base_quantity");
    const idxStockCat     = getIdx(S_STOCK, "category");
    const idxStockRawName = getIdx(S_STOCK, "raw_name");
    const idxStockUnit    = getIdx(S_STOCK, "base_unit");

    for (let i = 1; i < dataKK.length; i++) {
      const row = dataKK[i];
      const pKK = String(row[idxStockPeriod] || "").trim().replace(/\.0$/, '');
      const locCode = cleanCodeValue_(row[idxStockLoc]);
      const itemCode = cleanCodeValue_(row[idxStockItem]);

      const baseQtyVal = row[idxStockBaseQty];
      const actualQty = (baseQtyVal !== "" && baseQtyVal !== null && !isNaN(Number(baseQtyVal))) 
                        ? Number(baseQtyVal) 
                        : 0;

      if (pKK === strPeriod && locCode && itemCode) {
        const key = `${locCode}:${itemCode}`;
        mapClosing[key] = (mapClosing[key] || 0) + actualQty;
        hasStocktakeRecord.add(key);
        setKeys.add(key);

        updateMasterFallback(
          itemCode,
          idxStockRawName !== -1 ? String(row[idxStockRawName] || "").trim() : "",
          idxStockUnit !== -1 ? String(row[idxStockUnit] || "").trim() : "",
          idxStockCat !== -1 ? String(row[idxStockCat] || "").trim() : ""
        );
      }
    }

    // 5. UPSERT VÀO BẢNG INVENTORY_QTY_SUMMARY
    let currentData = sIQ.getDataRange().getValues();
    const qtyCols = schemaMap[S_IQ].columns.sort((a, b) => a.colIndex - b.colIndex);
    const maxColIndex = Math.max(...qtyCols.map(c => c.colIndex));
    const headers = new Array(maxColIndex);
    qtyCols.forEach(c => {
      headers[c.colIndex - 1] = c.colHeader || c.colKey;
    });

    if (currentData.length === 0 || currentData[0][0] === "") {
      currentData = [headers];
    }

    const keyIndexMap = {};
    for (let i = 1; i < currentData.length; i++) {
      const p = String(currentData[i][idxIQPeriod] || "").trim().replace(/\.0$/, '');
      const loc = cleanCodeValue_(currentData[i][idxIQLoc]);
      const item = cleanCodeValue_(currentData[i][idxIQItem]);
      if (p && loc && item) keyIndexMap[`${p}_${loc}_${item}`] = i;
    }

    const nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

    Array.from(setKeys).sort().forEach(key => {
      const splitIdx = key.indexOf(":");
      const locCode = key.substring(0, splitIdx);
      const itemCode = key.substring(splitIdx + 1);

      const fullKey = `${strPeriod}_${locCode}_${itemCode}`;
      const master = mapItemMaster[itemCode] || {};

      const openingQty = mapOpeningQty[key] || 0;
      const inboundQty = mapInbound[key] || 0;
      const adjustQty  = mapAdjust[key] || 0;

      let closingQty = 0;
      let outboundQty = 0;

      if (hasStocktakeRecord.has(key)) {
        closingQty = mapClosing[key] || 0;
        outboundQty = (openingQty + inboundQty + adjustQty) - closingQty;
      } else {
        closingQty = 0;
        outboundQty = (openingQty + inboundQty + adjustQty) - closingQty;
      }

      if (outboundQty < 0) outboundQty = 0;
      if (closingQty < 0) closingQty = 0;

      const rowValues = new Array(maxColIndex).fill("");
      qtyCols.forEach(col => {
        const colIdx = col.colIndex - 1;
        const cKey = col.colKey;

        if (cKey === "period") rowValues[colIdx] = strPeriod;
        else if (cKey === "location_code") rowValues[colIdx] = locCode;
        else if (cKey === "item_code") rowValues[colIdx] = itemCode;
        else if (cKey === "item_name") rowValues[colIdx] = master.name || "N/A";
        else if (cKey === "base_unit") rowValues[colIdx] = master.unit || "N/A";
        else if (cKey === "opening_qty") rowValues[colIdx] = openingQty;
        else if (cKey === "inbound_qty") rowValues[colIdx] = inboundQty;
        else if (cKey === "adjust_qty") rowValues[colIdx] = adjustQty;
        else if (cKey === "closing_qty") rowValues[colIdx] = closingQty;
        else if (cKey === "outbound_qty") rowValues[colIdx] = outboundQty;
        else if (cKey === "category") rowValues[colIdx] = master.category || "Chưa phân nhóm";
        else if (cKey === "updated_at") rowValues[colIdx] = nowStr;
      });

      if (keyIndexMap[fullKey] !== undefined) {
        currentData[keyIndexMap[fullKey]] = rowValues;
      } else {
        currentData.push(rowValues);
      }
    });

    sIQ.clear({contentsOnly: true});
    sIQ.getRange(1, 1, currentData.length, maxColIndex).setValues(currentData);
    SpreadsheetApp.flush();
  }
  

  /**
   * 2. Tính toán và tổng hợp giá trị tồn kho (Value Summary) 
   * Áp dụng quy tắc Waterfall Price: Kỳ N -> Lùi kỳ lịch sử (N-1, N-2...) -> Fallback cost_price (Item Master)
   * @param {string} periodTarget - Kỳ cần tính (VD: "2026-08")
   * @param {boolean} isBatch - Cờ chỉ định có phải đang chạy hàng loạt hay không (true = ẩn popup tương tác)
   */
  static calculateValueSummary(periodTarget, isBatch = false) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const schemaMap = schemaGetMap();
    if (!schemaMap) throw new Error("Không thể tải cấu trúc SCHEMA từ hệ thống.");

    const S_QTY   = "INVENTORY_QTY_SUMMARY";
    const S_PRICE = "MONTHLY_PRICE_LIST";
    const S_VAL   = "INVENTORY_VALUE_SUMMARY";
    const S_ITEM  = "ITEM_MASTER";

    const sQty   = ss.getSheetByName(schemaGetSheetName(schemaMap, S_QTY));
    const sPrice = ss.getSheetByName(schemaGetSheetName(schemaMap, S_PRICE));
    const sVal   = ss.getSheetByName(schemaGetSheetName(schemaMap, S_VAL));
    const sItem  = ss.getSheetByName(schemaGetSheetName(schemaMap, S_ITEM));

    if (!sQty || !sPrice || !sVal) {
      throw new Error("Không tìm thấy đủ các sheet yêu cầu theo SCHEMA.");
    }

    const dataQty   = sQty.getDataRange().getValues();
    const targetClean = String(periodTarget).trim().replace(/\.0$/, '');
    const getIdx = (sName, cKey) => schemaGetColIndex(schemaMap, sName, cKey);

    const parsePeriod_ = (rawVal) => {
      if (!rawVal) return "";
      let str = String(rawVal).trim().replace(/\.0$/, '');
      if (rawVal instanceof Date) {
        const y = rawVal.getFullYear();
        const m = String(rawVal.getMonth() + 1).padStart(2, '0');
        return `${y}-${m}`;
      }
      if (/^\d{6}$/.test(str)) {
        return `${str.substring(0, 4)}-${str.substring(4, 6)}`;
      }
      return str;
    };

    // A. Tải cost_price từ ITEM_MASTER để làm fallback cuối cùng
    const mapCostPrice = {};
    if (sItem) {
      const dataItem = sItem.getDataRange().getValues();
      const idxItemCode = getIdx(S_ITEM, "item_code");
      const idxCostPrice = getIdx(S_ITEM, "cost_price");
      for (let i = 1; i < dataItem.length; i++) {
        const row = dataItem[i];
        const code = cleanCodeValue_(row[idxItemCode]);
        if (code && idxCostPrice !== -1) {
          mapCostPrice[code] = Number(row[idxCostPrice]) || 0;
        }
      }
    }

    // B. Tải toàn bộ bảng giá lịch sử từ MONTHLY_PRICE_LIST -> Map<item_code, Map<period, unit_price>>
    const priceHistoryMap = new Map();
    const dataPrice = sPrice.getDataRange().getValues();
    const idxPricePeriod = getIdx(S_PRICE, "period");
    const idxPriceItem   = getIdx(S_PRICE, "item_code");
    const idxPriceVal    = getIdx(S_PRICE, "unit_price");

    for (let i = 1; i < dataPrice.length; i++) {
      const rawP = dataPrice[i][idxPricePeriod];
      const p = parsePeriod_(rawP);
      const itemCode = cleanCodeValue_(dataPrice[i][idxPriceItem]);
      const price = Number(dataPrice[i][idxPriceVal]) || 0;

      if (itemCode && p) {
        if (!priceHistoryMap.has(itemCode)) {
          priceHistoryMap.set(itemCode, new Map());
        }
        priceHistoryMap.get(itemCode).set(p, price);
      }
    }

    // C. Hàm giải quyết đơn giá theo cơ chế Waterfall (Kỳ N -> Lùi dần về quá khứ -> Fallback Cost Price)
    const getResolvedPrice = (itemCode, currentPeriod) => {
      const itemPrices = priceHistoryMap.get(itemCode);
      if (itemPrices && itemPrices.size > 0) {
        let [year, month] = currentPeriod.includes('-') 
          ? currentPeriod.split('-').map(Number) 
          : [parseInt(currentPeriod.slice(0,4)), parseInt(currentPeriod.slice(4))];
        
        // Duyệt lùi tối đa 24 tháng về trước
        for (let i = 0; i < 24; i++) {
          const checkPeriod = `${year}-${String(month).padStart(2, '0')}`;
          if (itemPrices.has(checkPeriod)) {
            return itemPrices.get(checkPeriod);
          }
          month--;
          if (month < 1) {
            month = 12;
            year--;
          }
        }
      }

      // Fallback cost_price từ item_master nếu lịch sử không có
      if (mapCostPrice[itemCode] && mapCostPrice[itemCode] > 0) {
        return mapCostPrice[itemCode];
      }

      return 0;
    };

    // D. Chuẩn bị headers và dữ liệu hiện tại của INVENTORY_VALUE_SUMMARY (UPSERT)
    let currentValData = sVal.getDataRange().getValues();
    const valCols = schemaMap[S_VAL].columns.sort((a, b) => a.colIndex - b.colIndex);
    const maxColIndex = Math.max(...valCols.map(c => c.colIndex));
    const headers = new Array(maxColIndex);
    valCols.forEach(c => {
      headers[c.colIndex - 1] = c.colHeader || c.colKey;
    });

    if (currentValData.length === 0 || currentValData[0][0] === "") {
      currentValData = [headers];
    }

    const idxValPeriod = getIdx(S_VAL, "period");
    const idxValLoc    = getIdx(S_VAL, "location_code");
    const idxValItem   = getIdx(S_VAL, "item_code");

    const keyIndexMap = {};
    for (let i = 1; i < currentValData.length; i++) {
      const p = parsePeriod_(currentValData[i][idxValPeriod]);
      const loc = cleanCodeValue_(currentValData[i][idxValLoc]);
      const item = cleanCodeValue_(currentValData[i][idxValItem]);
      if (p && loc && item) {
        keyIndexMap[`${p}_${loc}_${item}`] = i;
      }
    }

    const nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

    const idxQtyPeriod = getIdx(S_QTY, "period");
    const idxQtyLoc    = getIdx(S_QTY, "location_code");
    const idxQtyItem   = getIdx(S_QTY, "item_code");

    let countMatched = 0;
    let countInserted = 0;
    let countUpdated = 0;

    for (let i = 1; i < dataQty.length; i++) {
      const rowQty = dataQty[i];
      const rawP = rowQty[idxQtyPeriod];
      const p = parsePeriod_(rawP);
      
      if (p !== targetClean) continue;

      const locationCode = cleanCodeValue_(rowQty[idxQtyLoc]);
      const itemCode     = cleanCodeValue_(rowQty[idxQtyItem]);
      if (!locationCode || !itemCode) continue;

      countMatched++;
      // Lấy đơn giá thông minh theo quy tắc Waterfall Pricing
      const unitPrice = getResolvedPrice(itemCode, targetClean);

      const rowValues = new Array(maxColIndex).fill("");
      valCols.forEach(col => {
        const colIdx = col.colIndex - 1;
        const cKey = col.colKey;

        if (cKey === "period") {
          rowValues[colIdx] = targetClean;
        } else if (cKey === "location_code") {
          rowValues[colIdx] = locationCode;
        } else if (cKey === "item_code") {
          rowValues[colIdx] = itemCode;
        } else if (cKey === "updated_at") {
          rowValues[colIdx] = nowStr;
        } else if (cKey.endsWith('_amt')) {
          const qtyKey = cKey.replace('_amt', '_qty');
          const qtyColIdx = getIdx(S_QTY, qtyKey);
          if (qtyColIdx !== -1) {
            const qtyVal = Number(rowQty[qtyColIdx]) || 0;
            rowValues[colIdx] = qtyVal * unitPrice;
          } else {
            rowValues[colIdx] = 0;
          }
        } else {
          const srcColIdx = getIdx(S_QTY, cKey);
          rowValues[colIdx] = srcColIdx !== -1 ? rowQty[srcColIdx] : "";
        }
      });

      const uniqueKey = `${targetClean}_${locationCode}_${itemCode}`;
      if (keyIndexMap[uniqueKey] !== undefined) {
        currentValData[keyIndexMap[uniqueKey]] = rowValues;
        countUpdated++;
      } else {
        currentValData.push(rowValues);
        keyIndexMap[uniqueKey] = currentValData.length - 1;
        countInserted++;
      }
    }

    sVal.getRange(1, 1, currentValData.length, maxColIndex).setValues(currentValData);
    SpreadsheetApp.flush();

    // Chỉ hiển thị hộp thoại pop-up khi chạy lẻ thủ công; bỏ qua hoàn toàn nếu gọi từ batch
    if (!isBatch) {
      SpreadsheetApp.getUi().alert(
        "✅ StockEngine Value (Upsert Thành Công)",
        `Đã xử lý kỳ: [${targetClean}] theo chuẩn Upsert (Waterfall Pricing)` +
        `\n- Khớp từ Qty: ${countMatched} dòng` +
        `\n- Thêm mới (Insert): ${countInserted} dòng` +
        `\n- Cập nhật (Update): ${countUpdated} dòng` +
        `\n- Tổng số dòng hiện tại trong bảng: ${currentValData.length - 1}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  }
}
