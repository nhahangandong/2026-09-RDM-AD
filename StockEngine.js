/**
 * Module StockEngine: Tính toán Tổng hợp Tồn kho (Số lượng) chuẩn Schema-driven.
 * Sử dụng bảng độc lập LOCATION_ITEM_BALANCE và lưu giá trị trực tiếp dạng số (không dùng công thức).
 */
class StockEngine {

  /**
   * Tính toán và Upsert dữ liệu tổng hợp tồn kho theo Kỳ (YYYY-MM)
   * @param {string|number} periodTarget - Kỳ tính toán (Ví dụ: '2026-07', '2026-08')
   */
  static calculateQtySummary(periodTarget) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();

    // Load bản đồ Schema
    const schemaMap = schemaGetMap();
    if (!schemaMap) throw new Error("Không thể tải cấu trúc SCHEMA từ hệ thống.");

    // Tra cứu Route qua RouteEngine
    const locationTypeMap = RouteEngine.getLocationTypeMap(ss, schemaMap);
    const routeRulesMap   = RouteEngine.getRouteRulesMap(ss, schemaMap);

    // Định nghĩa các Schema Key
    const S_IQ    = "INVENTORY_QTY_SUMMARY";
    const S_TRANS = "TRANSACTION";
    const S_STOCK = "STOCKTAKE";
    const S_ITEM  = "ITEM_MASTER";
    const S_BAL   = "LOCATION_ITEM_BALANCE";

    // Lấy Sheets theo Schema Name
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

    // Khai báo đầy đủ các index cột cho INVENTORY_QTY_SUMMARY (S_IQ)
    const idxIQPeriod = getIdx(S_IQ, "period");
    const idxIQLoc    = getIdx(S_IQ, "location_code");
    const idxIQItem   = getIdx(S_IQ, "item_code");

    // -------------------------------------------------------------
    // 1. TẢI DANH MỤC ITEM_MASTER
    // -------------------------------------------------------------
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

    // -------------------------------------------------------------
    // 2. NẠP TỒN ĐẦU KỲ TỪ BẢNG ĐỘC LẬP LOCATION_ITEM_BALANCE
    // -------------------------------------------------------------
    const mapOpeningQty = {};
    const mapOpeningCost = {};
    const setKeys = new Set(); 

    if (dataBal.length > 1) {
      const idxBalPeriod = getIdx(S_BAL, "period");
      const idxBalLoc    = getIdx(S_BAL, "location_code");
      const idxBalItem   = getIdx(S_BAL, "item_code");
      const idxBalQty    = getIdx(S_BAL, "opening_qty");
      const idxBalCost   = getIdx(S_BAL, "unit_cost");

      for (let i = 1; i < dataBal.length; i++) {
        const row = dataBal[i];
        const p = String(row[idxBalPeriod] || "").trim().replace(/\.0$/, '');
        const loc = cleanCodeValue_(row[idxBalLoc]);
        const item = cleanCodeValue_(row[idxBalItem]);

        if (p === strPeriod && loc && item) {
          const key = `${loc}:${item}`;
          const qty = Number(row[idxBalQty]) || 0;
          const cost = Number(row[idxBalCost]) || 0;

          if (qty > 0) {
            mapOpeningQty[key] = qty;
            mapOpeningCost[key] = cost;
            setKeys.add(key);
          }
        }
      }
    }

    // -------------------------------------------------------------
    // 3. TẬP HỢP CÁC GIAO DỊCH TRONG KỲ (TRANSACTION)
    // -------------------------------------------------------------
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

      if (fromType === "PHYSICAL" && routeRule.transactionType === "TRANSFER") {
        const keyFrom = `${fromCode}:${itemCode}`;
        setKeys.add(keyFrom);
        mapAdjust[keyFrom] = (mapAdjust[keyFrom] || 0) - qty;
      }
    }

    // -------------------------------------------------------------
    // 4. TẬP HỢP KIỂM KÊ TỪ STOCKTAKE
    // -------------------------------------------------------------
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

    // -------------------------------------------------------------
    // 5. UPSERT VÀO BẢNG INVENTORY_QTY_SUMMARY & CHUYỂN KỲ KẾ TIẾP
    // -------------------------------------------------------------
    let currentData = sIQ.getDataRange().getValues();
    const cols = schemaMap[S_IQ].columns;
    const headers = cols.map(c => c.colHeader || c.colKey);

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
    
    // Gọi hàm helper toàn cục getNextPeriod_ từ Utils.gs
    const nextPeriod = getNextPeriod_(strPeriod);
    const nextBalanceRowsToAdd = [];

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
      if (hasStocktakeRecord.has(key)) {
        closingQty = mapClosing[key] || 0;
      } else {
        closingQty = openingQty + inboundQty + adjustQty;
      }

      let outboundQty = (openingQty + inboundQty + adjustQty) - closingQty;
      if (outboundQty < 0) outboundQty = 0;

      if (closingQty > 0) {
        const unitCost = mapOpeningCost[key] || 0;
        const totalValue = closingQty * unitCost; // Lưu trực tiếp giá trị số để tối ưu hiệu suất, tránh lỗi công thức

        nextBalanceRowsToAdd.push([
          nextPeriod, 
          locCode, 
          itemCode, 
          closingQty, 
          unitCost, 
          totalValue, 
          nowStr
        ]);
      }

      const rowValues = new Array(cols.length);
      rowValues[getIdx(S_IQ, "period")]        = strPeriod;
      rowValues[getIdx(S_IQ, "location_code")]  = locCode;
      rowValues[getIdx(S_IQ, "item_code")]      = itemCode;
      rowValues[getIdx(S_IQ, "item_name")]      = master.name || "N/A";
      rowValues[getIdx(S_IQ, "base_unit")]      = master.unit || "N/A";
      rowValues[getIdx(S_IQ, "opening_qty")]   = openingQty;
      rowValues[getIdx(S_IQ, "inbound_qty")]   = inboundQty;
      rowValues[getIdx(S_IQ, "adjust_qty")]    = adjustQty;
      rowValues[getIdx(S_IQ, "closing_qty")]   = closingQty;
      rowValues[getIdx(S_IQ, "outbound_qty")]  = outboundQty;
      rowValues[getIdx(S_IQ, "category")]      = master.category || "Chưa phân nhóm";
      rowValues[getIdx(S_IQ, "updated_at")]    = nowStr;

      if (keyIndexMap[fullKey] !== undefined) {
        currentData[keyIndexMap[fullKey]] = rowValues;
      } else {
        currentData.push(rowValues);
      }
    });

    // Ghi dữ liệu vào Sheet INVENTORY_QTY_SUMMARY
    sIQ.getRange(1, 1, currentData.length, cols.length).setValues(currentData);

    // Gọi hàm helper toàn cục updateNextPeriodBalance_ từ Utils.gs
    updateNextPeriodBalance_(ss, schemaMap, nextPeriod, nextBalanceRowsToAdd);

    SpreadsheetApp.flush();
  }
}



/**
* Hàm tính giá trị tồn kho dựa trên _qty và month_price_list
*/

class StockEngineValue {
  static calculate(periodTarget) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const schemaMap = schemaGetMap();
    if (!schemaMap) throw new Error("Không thể tải cấu trúc SCHEMA từ hệ thống.");

    const S_QTY   = "INVENTORY_QTY_SUMMARY";
    const S_PRICE = "MONTHLY_PRICE_LIST";
    const S_VAL   = "INVENTORY_VALUE_SUMMARY";

    const sQtyName   = schemaGetSheetName(schemaMap, S_QTY);
    const sPriceName = schemaGetSheetName(schemaMap, S_PRICE);
    const sValName   = schemaGetSheetName(schemaMap, S_VAL);

    const sQty   = sQtyName ? ss.getSheetByName(sQtyName) : null;
    const sPrice = sPriceName ? ss.getSheetByName(sPriceName) : null;
    const sVal   = sValName ? ss.getSheetByName(sValName) : null;

    if (!sQty || !sPrice || !sVal) {
      throw new Error("Không tìm thấy đủ các sheet yêu cầu (INVENTORY_QTY_SUMMARY, MONTHLY_PRICE_LIST, INVENTORY_VALUE_SUMMARY) theo SCHEMA.");
    }

    const dataQty   = sQty.getDataRange().getValues();
    const dataPrice = sPrice.getDataRange().getValues();
    const dataVal   = sVal.getDataRange().getValues();

    const strPeriod = String(periodTarget).trim();
    const targetClean = strPeriod.replace(/[-\/]/g, '');
    const getIdx = (sName, cKey) => schemaGetColIndex(schemaMap, sName, cKey);

    // 1. Tải đơn giá từ MONTHLY_PRICE_LIST của kỳ tương ứng
    const mapUnitPrice = {};
    const idxPricePeriod = getIdx(S_PRICE, "period");
    const idxPriceItem   = getIdx(S_PRICE, "item_code");
    const idxPriceVal    = getIdx(S_PRICE, "unit_price");

    for (let i = 1; i < dataPrice.length; i++) {
      const p = String(dataPrice[i][idxPricePeriod] || "").trim().replace(/\.0$/, '').replace(/[-\/]/g, '');
      if (p === targetClean) {
        const itemCode = cleanCodeValue_(dataPrice[i][idxPriceItem]);
        const price = Number(dataPrice[i][idxPriceVal]) || 0;
        if (itemCode) {
          mapUnitPrice[itemCode] = price;
        }
      }
    }

    // 2. Chuẩn bị cấu trúc headers vật lý cho INVENTORY_VALUE_SUMMARY
    const valCols = schemaMap[S_VAL].columns.sort((a, b) => a.colIndex - b.colIndex);
    const maxColIndex = Math.max(...valCols.map(c => c.colIndex));
    const headers = new Array(maxColIndex);
    valCols.forEach(c => {
      headers[c.colIndex - 1] = c.colHeader;
    });

    let currentValData = dataVal;
    if (currentValData.length === 0 || currentValData[0][0] === "") {
      currentValData = [headers];
    } else {
      currentValData[0] = headers;
    }

    const keyIndexMap = {};
    const idxValPeriod = getIdx(S_VAL, "period");
    const idxValLoc    = getIdx(S_VAL, "location_code");
    const idxValItem   = getIdx(S_VAL, "item_code");

    for (let i = 1; i < currentValData.length; i++) {
      const p = String(currentValData[i][idxValPeriod] || "").trim().replace(/\.0$/, '').replace(/[-\/]/g, '');
      const loc = cleanCodeValue_(currentValData[i][idxValLoc]);
      const item = cleanCodeValue_(currentValData[i][idxValItem]);
      if (p && loc && item) {
        keyIndexMap[`${p}_${loc}_${item}`] = i;
      }
    }

    const nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

    // 3. Đọc dữ liệu từ INVENTORY_QTY_SUMMARY và ánh xạ tính giá trị theo quy tắc đối xứng _qty -> _value
    const idxQtyPeriod = getIdx(S_QTY, "period");
    const idxQtyLoc    = getIdx(S_QTY, "location_code");
    const idxQtyItem   = getIdx(S_QTY, "item_code");

    for (let i = 1; i < dataQty.length; i++) {
      const rowQty = dataQty[i];
      const p = String(rowQty[idxQtyPeriod] || "").trim().replace(/\.0$/, '').replace(/[-\/]/g, '');
      
      if (p !== targetClean) continue;

      const locationCode = cleanCodeValue_(rowQty[idxQtyLoc]);
      const itemCode     = cleanCodeValue_(rowQty[idxQtyItem]);
      if (!locationCode || !itemCode) continue;

      const unitPrice = mapUnitPrice[itemCode] || 0;
      const fullKey = `${targetClean}_${locationCode}_${itemCode}`;

      const rowValues = new Array(maxColIndex);
      valCols.forEach(col => {
        const colIdx = col.colIndex - 1;
        const cKey = col.colKey;

        if (cKey === "period") {
          rowValues[colIdx] = strPeriod;
        } else if (cKey === "location_code") {
          rowValues[colIdx] = locationCode;
        } else if (cKey === "item_code") {
          rowValues[colIdx] = itemCode;
        } else if (cKey === "unit_price") {
          rowValues[colIdx] = unitPrice;
        } else if (cKey === "updated_at") {
          rowValues[colIdx] = nowStr;
        } else if (cKey.endsWith('_value')) {
          // Tự động tìm cột _qty tương ứng và nhân với đơn giá vốn
          const qtyKey = cKey.replace('_value', '_qty');
          const qtyColIdx = getIdx(S_QTY, qtyKey);
          if (qtyColIdx !== -1) {
            const qtyVal = Number(rowQty[qtyColIdx]) || 0;
            rowValues[colIdx] = qtyVal * unitPrice;
          } else {
            rowValues[colIdx] = 0;
          }
        } else {
          // Các cột thông tin text/danh mục chung (tên hàng, đvt, nhóm hàng,...) lấy trực tiếp từ QTY
          const srcColIdx = getIdx(S_QTY, cKey);
          rowValues[colIdx] = srcColIdx !== -1 ? rowQty[srcColIdx] : "";
        }
      });

      if (keyIndexMap[fullKey] !== undefined) {
        currentValData[keyIndexMap[fullKey]] = rowValues;
      } else {
        currentValData.push(rowValues);
        keyIndexMap[fullKey] = currentValData.length - 1;
      }
    }

    // 4. Ghi kết quả ra sheet INVENTORY_VALUE_SUMMARY
    sVal.getRange(1, 1, currentValData.length, maxColIndex).setValues(currentValData);
    SpreadsheetApp.flush();
  }
}
