/**
 * Module AvgPriceEngine: Tính toán đơn giá bình quân gia quyền và quản lý danh mục giá.
 */
class AvgPriceEngine {

  /**
   * Hàm 1: monthly_avg_price - Tính giá bình quân giao dịch tức thời cho kỳ hiện tại
   * Ưu tiên lấy raw_name và total_amount chuẩn từ TRANSACTION.
   */
static monthly_avg_price(periodTarget) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const schemaMap = schemaGetMap();
    if (!schemaMap) throw new Error("Không thể tải cấu trúc SCHEMA từ hệ thống.");

    const locationTypeMap = RouteEngine.getLocationTypeMap(ss, schemaMap);
    const routeRulesMap   = RouteEngine.getRouteRulesMap(ss, schemaMap);

    const S_TRANS = "TRANSACTION";
    const S_PRICE = "MONTHLY_AVG_PRICE";

    const sTrans = ss.getSheetByName(schemaGetSheetName(schemaMap, S_TRANS));
    const sPrice = ss.getSheetByName(schemaGetSheetName(schemaMap, S_PRICE));

    if (!sPrice) throw new Error(`Không tìm thấy sheet ứng với schema_name: ${S_PRICE}`);

    const dataTrans = sTrans ? sTrans.getDataRange().getValues() : [];
    const dataPrice = sPrice ? sPrice.getDataRange().getValues() : [];

    const strPeriod = String(periodTarget).trim();
    const getIdx = (sName, cKey) => schemaGetColIndex(schemaMap, sName, cKey);

    const mapInboundQty = {};
    const mapInboundAmt = {};
    const itemDynamicInfo = {};

    const idxTransPeriod   = getIdx(S_TRANS, "period");
    const idxTransItem     = getIdx(S_TRANS, "item_code");
    const idxTransFromCode = getIdx(S_TRANS, "from_code");
    const idxTransToCode   = getIdx(S_TRANS, "to_code");
    const idxTransQty      = getIdx(S_TRANS, "base_quantity");
    const idxTransAmount   = getIdx(S_TRANS, "total_amount");
    const idxTransRawName  = getIdx(S_TRANS, "raw_name");
    const idxTransBaseUnit = getIdx(S_TRANS, "base_unit");
    const idxTransCategory = getIdx(S_TRANS, "category"); // Lấy nhóm hàng nếu đã fill sẵn trong transaction

    for (let i = 1; i < dataTrans.length; i++) {
      const row = dataTrans[i];
      const transPeriod = String(row[idxTransPeriod] || "").trim().replace(/\.0$/, '');
      if (transPeriod !== strPeriod) continue;

      const itemCode = cleanCodeValue_(row[idxTransItem]);
      const fromCode = cleanCodeValue_(row[idxTransFromCode]);
      const toCode   = cleanCodeValue_(row[idxTransToCode]);
      const qty      = Number(row[idxTransQty]) || 0;
      const amount   = idxTransAmount !== -1 ? (Number(row[idxTransAmount]) || 0) : 0;

      if (!itemCode || qty <= 0) continue;

      const fromType  = locationTypeMap.get(fromCode) || "UNKNOWN";
      const toType    = locationTypeMap.get(toCode) || "UNKNOWN";
      const routeRule = routeRulesMap.get(`${fromType}->${toType}`);

      if (!routeRule || !routeRule.isInventory) continue;

      if (toType === "PHYSICAL" && (routeRule.transactionType === "PURCHASE" || routeRule.transactionType === "INBOUND" || routeRule.transactionType === "TRANSFER")) {
        mapInboundQty[itemCode] = (mapInboundQty[itemCode] || 0) + qty;
        mapInboundAmt[itemCode] = (mapInboundAmt[itemCode] || 0) + amount;

        const rName = idxTransRawName !== -1 ? String(row[idxTransRawName] || "").trim() : "";
        const bUnit = idxTransBaseUnit !== -1 ? String(row[idxTransBaseUnit] || "").trim() : "";
        const cCat  = idxTransCategory !== -1 ? String(row[idxTransCategory] || "").trim() : "";

        // Ưu tiên lưu thông tin trực tiếp từ transaction
        if (!itemDynamicInfo[itemCode]) {
          itemDynamicInfo[itemCode] = { name: rName, unit: bUnit, category: cCat };
        } else {
          if (!itemDynamicInfo[itemCode].name && rName) itemDynamicInfo[itemCode].name = rName;
          if (!itemDynamicInfo[itemCode].unit && bUnit) itemDynamicInfo[itemCode].unit = bUnit;
          if (!itemDynamicInfo[itemCode].category && cCat) itemDynamicInfo[itemCode].category = cCat;
        }
      }
    }

    const cols = schemaMap[S_PRICE].columns;
    const headers = cols.map(c => c.colHeader || c.colKey);

    let currentData = dataPrice;
    if (currentData.length === 0 || currentData[0][0] === "") {
      currentData = [headers];
    }

    const idxPricePeriod = getIdx(S_PRICE, "period");
    const idxPriceItem   = getIdx(S_PRICE, "item_code");

    const keyIndexMap = {};
    for (let i = 1; i < currentData.length; i++) {
      const p = String(currentData[i][idxPricePeriod] || "").trim().replace(/\.0$/, '');
      const item = cleanCodeValue_(currentData[i][idxPriceItem]);
      if (p && item) keyIndexMap[`${p}_${item}`] = i;
    }

    const nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

    Object.keys(mapInboundQty).forEach(itemCode => {
      const totalQty = mapInboundQty[itemCode];
      const totalAmt = mapInboundAmt[itemCode] || 0;
      const unitPrice = totalQty > 0 ? (totalAmt / totalQty) : 0;

      const dynamic = itemDynamicInfo[itemCode] || {};
      
      const itemName = dynamic.name || "N/A";
      const baseUnit = dynamic.unit || "N/A";
      const category = dynamic.category || "Chưa phân nhóm";

      const fullKey = `${strPeriod}_${itemCode}`;
      const rowValues = new Array(cols.length);

      rowValues[getIdx(S_PRICE, "period")]        = strPeriod;
      rowValues[getIdx(S_PRICE, "item_code")]     = itemCode;
      rowValues[getIdx(S_PRICE, "raw_name")]     = itemName; // Khớp đúng với col_key trong schema của MONTHLY_AVG_PRICE
      rowValues[getIdx(S_PRICE, "base_unit")]     = baseUnit;
      rowValues[getIdx(S_PRICE, "inbound_qty")]   = totalQty;
      rowValues[getIdx(S_PRICE, "inbound_amount")]  = totalAmt;
      rowValues[getIdx(S_PRICE, "unit_price")]    = unitPrice;
      rowValues[getIdx(S_PRICE, "price_source")]  = "PURCHASE_AVG";
      rowValues[getIdx(S_PRICE, "category")]      = category;
      rowValues[getIdx(S_PRICE, "updated_at")]    = nowStr;

      if (keyIndexMap[fullKey] !== undefined) {
        currentData[keyIndexMap[fullKey]] = rowValues;
      } else {
        currentData.push(rowValues);
      }
    });

    sPrice.getRange(1, 1, currentData.length, cols.length).setValues(currentData);
    SpreadsheetApp.flush();
  }


  /**
   * Hàm 2: monthly_price_list - Tính giá bình quân chốt tồn kho cuối tháng
   * Giá trị được ghi đồng thời vào 2 sheet kỳ N vào monthly_price_list, kỳ N + 1 vào Balance_Opening.
   */
  static monthly_price_list(periodTarget) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const schemaMap = schemaGetMap();
    if (!schemaMap) throw new Error("Không thể tải cấu trúc SCHEMA từ hệ thống.");

    const locationTypeMap = RouteEngine.getLocationTypeMap(ss, schemaMap);
    const routeRulesMap   = RouteEngine.getRouteRulesMap(ss, schemaMap);

    const S_TRANS = "TRANSACTION";
    const S_ITEM  = "ITEM_MASTER";
    const S_PRICE = "MONTHLY_PRICE_LIST";
    const S_BAL   = "LOCATION_ITEM_BALANCE";

    const sTransName = schemaGetSheetName(schemaMap, S_TRANS);
    const sItemName  = schemaGetSheetName(schemaMap, S_ITEM);
    const sPriceName = schemaGetSheetName(schemaMap, S_PRICE);
    const sBalName   = schemaGetSheetName(schemaMap, S_BAL);

    const sTrans  = sTransName ? ss.getSheetByName(sTransName) : null;
    const sItem   = sItemName ? ss.getSheetByName(sItemName) : null;
    const sPrice  = sPriceName ? ss.getSheetByName(sPriceName) : null;
    const sBal    = sBalName ? ss.getSheetByName(sBalName) : null;

    if (!sPrice || !sBal) throw new Error("Không tìm thấy sheet MONTHLY_PRICE_LIST hoặc BALANCE_OPENING theo cấu trúc SCHEMA.");

    const dataTrans  = sTrans ? sTrans.getDataRange().getValues() : [];
    const dataItem   = sItem ? sItem.getDataRange().getValues() : [];
    const dataPrice  = sPrice ? sPrice.getDataRange().getValues() : [];
    const dataBal    = sBal ? sBal.getDataRange().getValues() : [];

    const strPeriod = String(periodTarget).trim();
    const getIdx = (sName, cKey) => schemaGetColIndex(schemaMap, sName, cKey);

    // 1. Tải danh mục ITEM_MASTER (chỉ lọc source_group === 'INVENTORY')
    const mapItemMaster = {};
    if (dataItem.length > 1) {
      const idxItemCode   = getIdx(S_ITEM, "item_code");
      const idxItemName   = getIdx(S_ITEM, "item_name");
      const idxItemCat    = getIdx(S_ITEM, "category");
      const idxItemUnit   = getIdx(S_ITEM, "base_unit");
      const idxItemCost   = getIdx(S_ITEM, "cost_price");
      const idxSourceGroup= getIdx(S_ITEM, "source_group");

      for (let i = 1; i < dataItem.length; i++) {
        const row = dataItem[i];
        const code = cleanCodeValue_(row[idxItemCode]);
        if (!code) continue;

        const sourceGroup = idxSourceGroup !== -1 ? String(row[idxSourceGroup] || "").trim().toUpperCase() : "";
        if (sourceGroup === "INVENTORY") {
          mapItemMaster[code] = {
            name: idxItemName !== -1 ? String(row[idxItemName] || "").trim() : "",
            category: idxItemCat !== -1 ? String(row[idxItemCat] || "").trim() : "",
            unit: idxItemUnit !== -1 ? String(row[idxItemUnit] || "").trim() : "",
            costPrice: idxItemCost !== -1 ? (Number(row[idxItemCost]) || 0) : 0
          };
        }
      }
    }

    // 2. Tải lịch sử đơn giá tháng trước
    const rawPrevPeriod = getPreviousPeriod_(strPeriod); 
    const prevPeriodClean = rawPrevPeriod.replace(/[-\/]/g, ''); 
    const mapHistoricalPrice = {};
    
    const idxPricePeriod = getIdx(S_PRICE, "period");
    const idxPriceItem   = getIdx(S_PRICE, "item_code");
    const idxPriceVal    = getIdx(S_PRICE, "unit_price");

    for (let i = 1; i < dataPrice.length; i++) {
      const p = String(dataPrice[i][idxPricePeriod] || "").trim().replace(/\.0$/, '').replace(/[-\/]/g, '');
      if (p === prevPeriodClean || p === rawPrevPeriod) {
        const item = cleanCodeValue_(dataPrice[i][idxPriceItem]);
        const price = Number(dataPrice[i][idxPriceVal]) || 0;
        if (item && price > 0) {
          mapHistoricalPrice[item] = price;
        }
      }
    }

    // 3. Quét TRANSACTION lấy lượng và giá trị nhập trong kỳ
    const mapInboundQty = {};
    const mapInboundAmt = {};
    const itemDynamicInfo = {};

    const idxTransPeriod   = getIdx(S_TRANS, "period");
    const idxTransItem     = getIdx(S_TRANS, "item_code");
    const idxTransFromCode = getIdx(S_TRANS, "from_code");
    const idxTransToCode   = getIdx(S_TRANS, "to_code");
    const idxTransQty      = getIdx(S_TRANS, "base_quantity");
    const idxTransAmount   = getIdx(S_TRANS, "total_amount");
    const idxTransBaseUnit = getIdx(S_TRANS, "base_unit");
    const idxTransCategory = getIdx(S_TRANS, "category");

    for (let i = 1; i < dataTrans.length; i++) {
      const row = dataTrans[i];
      const transPeriod = String(row[idxTransPeriod] || "").trim().replace(/\.0$/, '').replace(/[-\/]/g, '');
      if (transPeriod !== strPeriod.replace(/[-\/]/g, '')) continue;

      const itemCode = cleanCodeValue_(row[idxTransItem]);
      const fromCode = cleanCodeValue_(row[idxTransFromCode]);
      const toCode   = cleanCodeValue_(row[idxTransToCode]);
      const qty      = Number(row[idxTransQty]) || 0;
      const amount   = idxTransAmount !== -1 ? (Number(row[idxTransAmount]) || 0) : 0;

      if (!itemCode) continue;

      const fromType  = locationTypeMap.get(fromCode) || "UNKNOWN";
      const toType    = locationTypeMap.get(toCode) || "UNKNOWN";
      const routeRule = routeRulesMap.get(`${fromType}->${toType}`);

      if (!routeRule || !routeRule.isInventory) continue;

      if (toType === "PHYSICAL" && (routeRule.transactionType === "PURCHASE" || routeRule.transactionType === "INBOUND" || routeRule.transactionType === "TRANSFER")) {
        mapInboundQty[itemCode] = (mapInboundQty[itemCode] || 0) + qty;
        mapInboundAmt[itemCode] = (mapInboundAmt[itemCode] || 0) + amount;

        const bUnit = idxTransBaseUnit !== -1 ? String(row[idxTransBaseUnit] || "").trim() : "";
        const cCat  = idxTransCategory !== -1 ? String(row[idxTransCategory] || "").trim() : "";

        if (!itemDynamicInfo[itemCode]) {
          itemDynamicInfo[itemCode] = { unit: bUnit, category: cCat };
        } else {
          if (!itemDynamicInfo[itemCode].unit && bUnit) itemDynamicInfo[itemCode].unit = bUnit;
          if (!itemDynamicInfo[itemCode].category && cCat) itemDynamicInfo[itemCode].category = cCat;
        }
      }
    }

    // 4. Chuẩn bị cấu trúc headers vật lý cho MONTHLY_PRICE_LIST
    const priceCols = schemaMap[S_PRICE].columns.sort((a, b) => a.colIndex - b.colIndex);
    const maxColIndex = Math.max(...priceCols.map(c => c.colIndex));
    const headers = new Array(maxColIndex);
    priceCols.forEach(c => {
      headers[c.colIndex - 1] = c.colHeader;
    });

    let currentPriceData = dataPrice;
    if (currentPriceData.length === 0 || currentPriceData[0][0] === "") {
      currentPriceData = [headers];
    } else {
      currentPriceData[0] = headers;
    }

    const keyIndexMap = {};
    for (let i = 1; i < currentPriceData.length; i++) {
      const p = String(currentPriceData[i][idxPricePeriod] || "").trim().replace(/\.0$/, '').replace(/[-\/]/g, '');
      const item = cleanCodeValue_(currentPriceData[i][idxPriceItem]);
      if (p && item) {
        keyIndexMap[`${p}_${item}`] = i;
      }
    }

    const nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    const mapUnitPrice = {};

    // 5. Gom tất cả các mã cần xử lý từ nhiều nguồn (Transaction, Lịch sử, Master, Balance Kỳ N và Kỳ N+1)
    const activeItemCodes = new Set(Object.keys(mapInboundQty));
    Object.keys(mapHistoricalPrice).forEach(k => activeItemCodes.add(k));
    Object.keys(mapItemMaster).forEach(k => {
      if (mapItemMaster[k].costPrice > 0) activeItemCodes.add(k);
    });

    // Vét thêm từ BALANCE_OPENING của cả kỳ hiện tại (N) và kỳ kế tiếp (N+1) để tránh bỏ sót mã bị trễ/lọt sổ
    const targetPeriodClean = strPeriod.replace(/[-\/]/g, '');
    const nextPeriodClean = getNextPeriod_(strPeriod).replace(/[-\/]/g, '');
    
    const idxBalPeriod = getIdx(S_BAL, "period");
    const idxBalItem   = getIdx(S_BAL, "item_code");

    for (let i = 1; i < dataBal.length; i++) {
      const p = String(dataBal[i][idxBalPeriod] || "").trim().replace(/\.0$/, '').replace(/[-\/]/g, '');
      if (p === targetPeriodClean || p === nextPeriodClean) {
        const itemCode = cleanCodeValue_(dataBal[i][idxBalItem]);
        if (itemCode) activeItemCodes.add(itemCode);
      }
    }

    activeItemCodes.forEach(itemCode => {
      if (!mapItemMaster[itemCode]) return; 

      const totalQty = mapInboundQty[itemCode] || 0;
      const totalAmt = mapInboundAmt[itemCode] || 0;
      
      let unitPrice = 0;
      let priceSource = "";

      const master = mapItemMaster[itemCode] || {};
      const dynamic = itemDynamicInfo[itemCode] || {};

      if (totalQty > 0) {
        unitPrice = totalAmt / totalQty;
        priceSource = "PURCHASE_AVG";
      } else if (mapHistoricalPrice[itemCode] !== undefined && mapHistoricalPrice[itemCode] > 0) {
        unitPrice = mapHistoricalPrice[itemCode];
        priceSource = "HISTORICAL_LAST";
      } else if (master.costPrice && master.costPrice > 0) {
        unitPrice = master.costPrice;
        priceSource = "FIXED_MASTER";
      } else {
        unitPrice = 0;
        priceSource = "FIXED_MASTER";
      }

      mapUnitPrice[itemCode] = unitPrice;

      const itemName = master.name || "N/A";
      const baseUnit = master.unit || dynamic.unit || "N/A";
      const category = master.category || dynamic.category || "Chưa phân nhóm";

      const fullKey = `${targetPeriodClean}_${itemCode}`;
      const rowValues = new Array(maxColIndex);

      priceCols.forEach(col => {
        const colIdx = col.colIndex - 1;
        switch (col.colKey) {
          case "period":        rowValues[colIdx] = strPeriod; break;
          case "item_code":     rowValues[colIdx] = itemCode; break;
          case "item_name":     rowValues[colIdx] = itemName; break;
          case "base_unit":     rowValues[colIdx] = baseUnit; break;
          case "inbound_qty":   rowValues[colIdx] = totalQty; break;
          case "inbound_amount":rowValues[colIdx] = totalAmt; break;
          case "unit_price":    rowValues[colIdx] = unitPrice; break;
          case "price_source":  rowValues[colIdx] = priceSource; break;
          case "category":      rowValues[colIdx] = category; break;
          case "updated_at":    rowValues[colIdx] = nowStr; break;
          default:              rowValues[colIdx] = ""; break;
        }
      });

      if (keyIndexMap[fullKey] !== undefined) {
        currentPriceData[keyIndexMap[fullKey]] = rowValues;
      } else {
        currentPriceData.push(rowValues);
        keyIndexMap[fullKey] = currentPriceData.length - 1;
      }
    });

    // 6. Ghi dữ liệu ra sheet MONTHLY_PRICE_LIST
    sPrice.getRange(1, 1, currentPriceData.length, maxColIndex).setValues(currentPriceData);

    // 7. Cập nhật ngược lại BALANCE_OPENING độc quyền cho KỲ KẾ TIẾP (nextPeriod)
    const idxBalQty    = getIdx(S_BAL, "opening_qty");
    const idxBalCost   = getIdx(S_BAL, "unit_cost");
    const idxBalVal    = getIdx(S_BAL, "opening_value");
    const idxBalDate   = getIdx(S_BAL, "updated_at");

    let isBalUpdated = false;
    for (let i = 1; i < dataBal.length; i++) {
      const p = String(dataBal[i][idxBalPeriod] || "").trim().replace(/\.0$/, '').replace(/[-\/]/g, '');
      
      // Chỉ ghi nhận vào số dư đầu kỳ của KỲ KẾ TIẾP (ví dụ chạy tháng 7 -> cập nhật đầu kỳ tháng 8)
      if (p === nextPeriodClean) {
        const itemCode = cleanCodeValue_(dataBal[i][idxBalItem]);
        const qty = Number(dataBal[i][idxBalQty]) || 0;
        
        if (mapUnitPrice[itemCode] !== undefined) {
          const unitCost = mapUnitPrice[itemCode];
          const totalVal = qty * unitCost;

          if (idxBalCost !== -1) dataBal[i][idxBalCost] = unitCost;
          if (idxBalVal !== -1)  dataBal[i][idxBalVal]  = totalVal;
          if (idxBalDate !== -1) dataBal[i][idxBalDate] = nowStr;
          
          isBalUpdated = true;
        }
      }
    }

    if (isBalUpdated) {
      sBal.getDataRange().setValues(dataBal);
    }

    SpreadsheetApp.flush();
  }
}
