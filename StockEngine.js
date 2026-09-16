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
   * Kế thừa closing_amt tháng trước cho opening_amt, lấy đúng inbound_amt thực tế (0 nếu hàng tặng),
   * và dùng unitPrice hiện hành cho các cột khác.
   * Tính toán và cập nhật bảng giá trị tồn kho (INVENTORY_VALUE_SUMMARY)
   * Tuân thủ chuẩn: 
   * - opening_amt = opening_qty * Đơn giá WAVG chính thức của kỳ N-1
   * - inbound_amt = Lấy trực tiếp từ MONTHLY_AVG_PRICE
   * - outbound_amt, closing_amt... = Số lượng tương ứng * Đơn giá WAVG chính thức của kỳ N
   */
  static calculateValueSummary(periodTarget, isBatch = false) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const schemaMap = schemaGetMap();
    if (!schemaMap) throw new Error("Không thể tải cấu trúc SCHEMA từ hệ thống.");

    const S_QTY   = "INVENTORY_QTY_SUMMARY";
    const S_PRICE = "MONTHLY_PRICE_LIST";      // Chứa đơn giá bình quân gia quyền (WAVG) chính thức cuối kỳ
    const S_MAPrice = "MONTHLY_AVG_PRICE";     // Chứa tổng tiền nhập thực tế (inbound_amount) trong kỳ
    const S_VAL   = "INVENTORY_VALUE_SUMMARY";
    const S_ITEM  = "ITEM_MASTER";

    const sQty    = ss.getSheetByName(schemaGetSheetName(schemaMap, S_QTY));
    const sPrice  = ss.getSheetByName(schemaGetSheetName(schemaMap, S_PRICE));
    const sMAPrice = ss.getSheetByName(schemaGetSheetName(schemaMap, S_MAPrice));
    const sVal    = ss.getSheetByName(schemaGetSheetName(schemaMap, S_VAL));
    const sItem   = ss.getSheetByName(schemaGetSheetName(schemaMap, S_ITEM));

    if (!sQty || !sPrice || !sVal) {
      throw new Error("Không tìm thấy đủ các sheet yêu cầu theo SCHEMA (QTY, PRICE, VALUE).");
    }

    const dataQty    = sQty.getDataRange().getValues();
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

    const getPreviousPeriod = (currPeriod) => {
      let [year, month] = currPeriod.includes('-') 
        ? currPeriod.split('-').map(Number) 
        : [parseInt(currPeriod.slice(0,4)), parseInt(currPeriod.slice(4))];
      month--;
      if (month < 1) {
        month = 12;
        year--;
      }
      return `${year}-${String(month).padStart(2, '0')}`;
    };

    const prevPeriodClean = getPreviousPeriod(targetClean);

    // A. Tải cost_price dự phòng từ ITEM_MASTER
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

    // B. Tải bảng giá lịch sử từ MONTHLY_PRICE_LIST (WAVG chính thức của từng kỳ)
    const priceHistoryMap = new Map(); // Map<itemCode, Map<period, unit_price>>
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

    // C. Tải giá trị tiền nhập thực tế trong kỳ từ MONTHLY_AVG_PRICE (inbound_amount)
    const mapActualInboundAmt = {}; // Key: item_code -> Total Amount
    if (sMAPrice) {
      const dataMAPrice = sMAPrice.getDataRange().getValues();
      const idxMAPPeriod = getIdx(S_MAPrice, "period");
      const idxMAPItem   = getIdx(S_MAPrice, "item_code");
      const idxMAPAmt    = getIdx(S_MAPrice, "inbound_amount");

      for (let i = 1; i < dataMAPrice.length; i++) {
        const row = dataMAPrice[i];
        const p = parsePeriod_(row[idxMAPPeriod]);
        if (p !== targetClean) continue;

        const itemCode = cleanCodeValue_(row[idxMAPItem]);
        const amt = idxMAPAmt !== -1 ? (Number(row[idxMAPAmt]) || 0) : 0;
        if (itemCode) {
          mapActualInboundAmt[itemCode] = (mapActualInboundAmt[itemCode] || 0) + amt;
        }
      }
    }

    // D. Hàm giải quyết đơn giá theo cơ chế Waterfall (tra cứu ngược về quá khứ nếu kỳ hiện tại chưa có)
    const getResolvedPrice = (itemCode, currentPeriod) => {
      const itemPrices = priceHistoryMap.get(itemCode);
      if (itemPrices && itemPrices.size > 0) {
        let [year, month] = currentPeriod.includes('-') 
          ? currentPeriod.split('-').map(Number) 
          : [parseInt(currentPeriod.slice(0,4)), parseInt(currentPeriod.slice(4))];
        
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

      if (mapCostPrice[itemCode] && mapCostPrice[itemCode] > 0) {
        return mapCostPrice[itemCode];
      }

      return 0;
    };

    // Lấy đơn giá WAVG chính thức của kỳ N-1 phục vụ tính opening_amt (Hướng 1)
    const getPrevPeriodWavgPrice = (itemCode, prevPeriod) => {
      const itemPrices = priceHistoryMap.get(itemCode);
      if (itemPrices && itemPrices.has(prevPeriod)) {
        return itemPrices.get(prevPeriod);
      }
      // Nếu kỳ N-1 không có, fallback về hàm waterfall gần nhất
      return getResolvedPrice(itemCode, prevPeriod);
    };

    // E. Chuẩn bị headers và dữ liệu UPSERT cho INVENTORY_VALUE_SUMMARY
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

    // F. Vòng lặp chính xử lý từng dòng dữ liệu số lượng của kỳ đích
    for (let i = 1; i < dataQty.length; i++) {
      const rowQty = dataQty[i];
      const rawP = rowQty[idxQtyPeriod];
      const p = parsePeriod_(rawP);
      
      if (p !== targetClean) continue;

      const locationCode = cleanCodeValue_(rowQty[idxQtyLoc]);
      const itemCode     = cleanCodeValue_(rowQty[idxQtyItem]);
      if (!locationCode || !itemCode) continue;

      countMatched++;

      // Đơn giá WAVG hiện hành của kỳ N (dùng cho xuất, tồn cuối...)
      const unitPriceCurrent = getResolvedPrice(itemCode, targetClean);
      
      // Đơn giá WAVG chính thức của kỳ N-1 (dùng cho đầu kỳ theo Hướng 1)
      const unitPricePrev = getPrevPeriodWavgPrice(itemCode, prevPeriodClean);

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
        } else if (cKey === "opening_amt") {
          // HƯỚNG 1: Lấy số lượng đầu kỳ nhân với đơn giá WAVG chuẩn của kỳ N-1
          // Triệt tiêu hoàn toàn rác giá trị treo từ quá khứ nếu opening_qty = 0
          const qtyOpenIdx = getIdx(S_QTY, "opening_qty");
          const openQty = qtyOpenIdx !== -1 ? (Number(rowQty[qtyOpenIdx]) || 0) : 0;
          rowValues[colIdx] = openQty * unitPricePrev;
        } else if (cKey === "inbound_amt") {
          // Lấy trực tiếp giá trị tiền nhập thực tế từ MONTHLY_AVG_PRICE đã chuẩn hóa
          rowValues[colIdx] = mapActualInboundAmt[itemCode] !== undefined ? mapActualInboundAmt[itemCode] : 0;
        } else if (cKey.endsWith('_amt')) {
          // Các cột giá trị khác (closing_amt, outbound_amt...) = số lượng tương ứng * unitPriceCurrent
          const qtyKey = cKey.replace('_amt', '_qty');
          const qtyColIdx = getIdx(S_QTY, qtyKey);
          if (qtyColIdx !== -1) {
            const qtyVal = Number(rowQty[qtyColIdx]) || 0;
            rowValues[colIdx] = qtyVal * unitPriceCurrent;
          } else {
            rowValues[colIdx] = 0;
          }
        } else {
          // Kế thừa các cột thông tin mô tả khác từ sheet Qty
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

    if (!isBatch) {
      SpreadsheetApp.getUi().alert(
        "✅ StockEngine Value (Chuẩn Hướng 1)",
        `Đã xử lý kỳ: [${targetClean}] thành công!` +
        `\n- Khớp từ Qty: ${countMatched} dòng` +
        `\n- Thêm mới: ${countInserted} dòng | Cập nhật: ${countUpdated} dòng`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  }
  
  
  static generateOpeningBalance(targetPeriod, sourcePeriod) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const schemaMap = schemaGetMap();
    if (!schemaMap) throw new Error("Không thể tải cấu trúc SCHEMA từ hệ thống.");

    const S_SRC = "STOCKTAKE"; 
    const S_ITEM = "ITEM_MASTER";
    const S_BAL = "LOCATION_ITEM_BALANCE";

    const sSrcName = schemaGetSheetName(schemaMap, S_SRC);
    const sItemName = schemaGetSheetName(schemaMap, S_ITEM);
    const sBalName = schemaGetSheetName(schemaMap, S_BAL);

    const sSrc = sSrcName ? ss.getSheetByName(sSrcName) : null;
    const sItem = sItemName ? ss.getSheetByName(sItemName) : null;
    const sBal = sBalName ? ss.getSheetByName(sBalName) : null;

    if (!sSrc) throw new Error(`Không tìm thấy sheet ứng với schema "${S_SRC}"`);
    if (!sBal) throw new Error("Không tìm thấy sheet LOCATION_ITEM_BALANCE.");

    const dataSrc = sSrc.getDataRange().getValues();
    const dataItem = sItem ? sItem.getDataRange().getValues() : [];
    const dataBal = sBal.getDataRange().getValues();

    const getIdx = (sName, cKey) => schemaGetColIndex(schemaMap, sName, cKey);

    // 1. Lấy danh sách đơn giá gốc từ ITEM_MASTER
    const mapItemCost = {};
    if (dataItem.length > 1) {
      const idxCode = getIdx(S_ITEM, "item_code");
      const idxCost = getIdx(S_ITEM, "cost_price");
      for (let i = 1; i < dataItem.length; i++) {
        const code = cleanCodeValue_(dataItem[i][idxCode]);
        const cost = Number(dataItem[i][idxCost]) || 0;
        if (code) mapItemCost[code] = cost;
      }
    }

    // 2. Lấy chỉ số các cột từ Schema STOCKTAKE
    const idxSrcPeriod = getIdx(S_SRC, "period");
    const idxSrcLoc    = getIdx(S_SRC, "location");      
    const idxSrcItem   = getIdx(S_SRC, "item_code");     
    const idxSrcQty    = getIdx(S_SRC, "base_quantity"); 

    // Thay vì dùng chuỗi ghép dễ bị lỗi split, ta lưu thẳng cấu trúc object rõ ràng
    const aggregatedBalance = {}; // Key: `${locCode}__${itemCode}` -> { locCode, itemCode, totalQty }
    const targetPeriodClean = String(sourcePeriod).trim().replace(/[-\/]/g, '');

    // 3. Quét và CỘNG DỒN số lượng
    for (let i = 1; i < dataSrc.length; i++) {
      const row = dataSrc[i];
      const p = String(row[idxSrcPeriod] || "").trim().replace(/\.0$/, '').replace(/[-\/]/g, '');
      
      if (p === targetPeriodClean) {
        const locCode = cleanCodeValue_(row[idxSrcLoc]);   
        const itemCode = cleanCodeValue_(row[idxSrcItem]); 
        const qty = Number(row[idxSrcQty]) || 0;

        if (locCode && itemCode) {
          const key = `${locCode}__${itemCode}`;
          if (!aggregatedBalance[key]) {
            aggregatedBalance[key] = {
              locCode: locCode,
              itemCode: itemCode,
              totalQty: 0
            };
          }
          aggregatedBalance[key].totalQty += qty;
        }
      }
    }

    // 4. Chuẩn bị cấu trúc bảng LOCATION_ITEM_BALANCE để làm cơ chế UPSERT
    const balCols = schemaMap[S_BAL].columns.sort((a, b) => a.colIndex - b.colIndex);
    const maxColIndex = Math.max(...balCols.map(c => c.colIndex));
    const headers = new Array(maxColIndex);
    balCols.forEach(c => {
      headers[c.colIndex - 1] = c.colHeader || c.colKey;
    });

    let currentBalData = dataBal.length > 0 && dataBal[0][0] !== "" ? dataBal : [headers];
    
    const idxBalPeriod = getIdx(S_BAL, "period");
    const idxBalLoc    = getIdx(S_BAL, "location_code");
    const idxBalItem   = getIdx(S_BAL, "item_code");

    const existingRowMap = {};
    for (let i = 1; i < currentBalData.length; i++) {
      const row = currentBalData[i];
      const p = String(row[idxBalPeriod] || "").trim().replace(/\.0$/, '').replace(/[-\/]/g, '');
      const loc = cleanCodeValue_(row[idxBalLoc]);
      const item = cleanCodeValue_(row[idxBalItem]);
      if (p && loc && item) {
        existingRowMap[`${p}_${loc}_${item}`] = i; 
      }
    }

    const nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    const targetPeriodCleanFormat = String(targetPeriod).trim().replace(/[-\/]/g, '');

    // 5. Thực hiện Upsert an toàn tuyệt đối không cắt chuỗi
    Object.keys(aggregatedBalance).forEach(key => {
      const itemData = aggregatedBalance[key];
      const locCode = itemData.locCode;
      const itemCode = itemData.itemCode;
      const totalQty = itemData.totalQty;

      const unitCost = mapItemCost[itemCode] || 0;
      const openingVal = totalQty * unitCost;

      const rowValues = new Array(maxColIndex).fill("");
      balCols.forEach(col => {
        const colIdx = col.colIndex - 1;
        switch (col.colKey) {
          case "period":        rowValues[colIdx] = targetPeriod; break; 
          case "location_code": rowValues[colIdx] = locCode; break;
          case "item_code":     rowValues[colIdx] = itemCode; break;
          case "opening_qty":   rowValues[colIdx] = totalQty; break;
          case "unit_cost":     rowValues[colIdx] = unitCost; break;
          case "opening_value": rowValues[colIdx] = openingVal; break;
          case "updated_at":    rowValues[colIdx] = nowStr; break;
          default:              rowValues[colIdx] = ""; break;
        }
      });

      const uniqueKey = `${targetPeriodCleanFormat}_${locCode}_${itemCode}`;

      if (existingRowMap[uniqueKey] !== undefined) {
        const rowIndex = existingRowMap[uniqueKey];
        currentBalData[rowIndex] = rowValues;
      } else {
        currentBalData.push(rowValues);
        existingRowMap[uniqueKey] = currentBalData.length - 1;
      }
    });

    sBal.getRange(1, 1, currentBalData.length, maxColIndex).setValues(currentBalData);
    SpreadsheetApp.flush();
    SpreadsheetApp.getActiveSpreadsheet().toast(`Kết chuyển và đồng bộ số dư thành công cho kỳ ${targetPeriod}!`, "Hoàn tất", 5);
  }
  
}
