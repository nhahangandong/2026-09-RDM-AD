/**
 * Module AvgPriceEngine: Tính toán đơn giá bình quân gia quyền và quản lý danh mục giá.
 * Tuân thủ quy ước: Upsert dữ liệu, tuyệt đối không xóa bảng (clear), không ghi đè LOCATION_ITEM_BALANCE.
 */
class AvgPriceEngine {

  /**
   * Hàm 1: monthly_avg_price - Tính giá bình quân giao dịch tức thời cho kỳ hiện tại
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
    const idxTransCategory = getIdx(S_TRANS, "category");

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

        if (!itemDynamicInfo[itemCode]) {
          itemDynamicInfo[itemCode] = { name: rName, unit: bUnit, category: cCat };
        } else {
          if (!itemDynamicInfo[itemCode].name && rName) itemDynamicInfo[itemCode].name = rName;
          if (!itemDynamicInfo[itemCode].unit && bUnit) itemDynamicInfo[itemCode].unit = bUnit;
          if (!itemDynamicInfo[itemCode].category && cCat) itemDynamicInfo[itemCode].category = cCat;
        }
      }
    }

    const priceCols = schemaMap[S_PRICE].columns.sort((a, b) => a.colIndex - b.colIndex);
    const maxColIndex = Math.max(...priceCols.map(c => c.colIndex));
    const headers = new Array(maxColIndex);
    priceCols.forEach(c => {
      headers[c.colIndex - 1] = c.colHeader || c.colKey;
    });

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
      const rowValues = new Array(maxColIndex).fill("");

      priceCols.forEach(col => {
        const colIdx = col.colIndex - 1;
        switch (col.colKey) {
          case "period":        rowValues[colIdx] = strPeriod; break;
          case "item_code":     rowValues[colIdx] = itemCode; break;
          case "raw_name":      rowValues[colIdx] = itemName; break;
          case "base_unit":     rowValues[colIdx] = baseUnit; break;
          case "inbound_qty":   rowValues[colIdx] = totalQty; break;
          case "inbound_amount":rowValues[colIdx] = totalAmt; break;
          case "unit_price":    rowValues[colIdx] = unitPrice; break;
          case "price_source":  rowValues[colIdx] = "PURCHASE_AVG"; break;
          case "category":      rowValues[colIdx] = category; break;
          case "updated_at":    rowValues[colIdx] = nowStr; break;
          default:              rowValues[colIdx] = ""; break;
        }
      });

      if (keyIndexMap[fullKey] !== undefined) {
        currentData[keyIndexMap[fullKey]] = rowValues;
      } else {
        currentData.push(rowValues);
        keyIndexMap[fullKey] = currentData.length - 1;
      }
    });

    sPrice.getRange(1, 1, currentData.length, maxColIndex).setValues(currentData);
    SpreadsheetApp.flush();
  }


  /**
   * Hàm 2: monthly_price_list - Tính giá bình quân chốt tồn kho cuối tháng
   * Đã sửa lỗi: Nạp đầy đủ cả closing_qty lẫn closing_amt từ tháng trước để tính đúng openQty và openVal.
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
    const S_QTY   = "INVENTORY_QTY_SUMMARY";
    const S_VAL   = "INVENTORY_VALUE_SUMMARY";
    const S_BAL   = "LOCATION_ITEM_BALANCE";

    const sTransName = schemaGetSheetName(schemaMap, S_TRANS);
    const sItemName  = schemaGetSheetName(schemaMap, S_ITEM);
    const sPriceName = schemaGetSheetName(schemaMap, S_PRICE);
    const sQtyName   = schemaGetSheetName(schemaMap, S_QTY);
    const sValName   = schemaGetSheetName(schemaMap, S_VAL);
    const sBalName   = schemaGetSheetName(schemaMap, S_BAL);

    const sTrans  = sTransName ? ss.getSheetByName(sTransName) : null;
    const sItem   = sItemName ? ss.getSheetByName(sItemName) : null;
    const sPrice  = sPriceName ? ss.getSheetByName(sPriceName) : null;
    const sQty    = sQtyName ? ss.getSheetByName(sQtyName) : null;
    const sVal    = sValName ? ss.getSheetByName(sValName) : null;
    const sBal    = sBalName ? ss.getSheetByName(sBalName) : null;

    if (!sPrice) throw new Error("Không tìm thấy sheet MONTHLY_PRICE_LIST theo cấu trúc SCHEMA.");

    const dataTrans  = sTrans ? sTrans.getDataRange().getValues() : [];
    const dataItem   = sItem ? sItem.getDataRange().getValues() : [];
    const dataPrice  = sPrice ? sPrice.getDataRange().getValues() : [];
    const dataQty    = sQty ? sQty.getDataRange().getValues() : [];
    const dataVal    = sVal ? sVal.getDataRange().getValues() : [];
    const dataBal    = sBal ? sBal.getDataRange().getValues() : [];

    const strPeriod = String(periodTarget).trim();
    const targetPeriodClean = strPeriod.replace(/[-\/]/g, '');
    const getIdx = (sName, cKey) => schemaGetColIndex(schemaMap, sName, cKey);

    // 1. Tải danh mục ITEM_MASTER
    const mapItemMaster = {};
    if (dataItem.length > 1) {
      const idxItemCode    = getIdx(S_ITEM, "item_code");
      const idxItemName    = getIdx(S_ITEM, "item_name");
      const idxItemCat     = getIdx(S_ITEM, "category");
      const idxItemUnit    = getIdx(S_ITEM, "base_unit");
      const idxItemCost    = getIdx(S_ITEM, "cost_price");
      const idxSourceGroup = getIdx(S_ITEM, "source_group");

      for (let i = 1; i < dataItem.length; i++) {
        const row = dataItem[i];
        const code = cleanCodeValue_(row[idxItemCode]);
        if (!code) continue;

        const sourceGroup = idxSourceGroup !== -1 ? String(row[idxSourceGroup] || "").trim().toUpperCase() : "";
        if (sourceGroup === "INVENTORY" || sourceGroup === "") {
          mapItemMaster[code] = {
            name: idxItemName !== -1 ? String(row[idxItemName] || "").trim() : "",
            category: idxItemCat !== -1 ? String(row[idxItemCat] || "").trim() : "",
            unit: idxItemUnit !== -1 ? String(row[idxItemUnit] || "").trim() : "",
            costPrice: idxItemCost !== -1 ? (Number(row[idxItemCost]) || 0) : 0
          };
        }
      }
    }

    // 2. LẤY SỐ DƯ ĐẦU KỲ (Cả Lượng lẫn Tiền) GỘP TẤT CẢ CÁC KHO từ tháng trước (N-1)
    const mapOpeningQtyByItem = {};
    const mapOpeningValByItem = {};

    const rawPrevPeriod = getPreviousPeriod_(strPeriod); 
    const prevPeriodClean = rawPrevPeriod.replace(/[-\/]/g, '');

    // 2.1. Lấy lượng đầu kỳ gộp từ INVENTORY_QTY_SUMMARY của tháng trước
    if (dataQty.length > 1) {
      const idxQPeriod = getIdx(S_QTY, "period");
      const idxQItem   = getIdx(S_QTY, "item_code");
      const idxQClose  = getIdx(S_QTY, "closing_qty"); // Tồn cuối tháng trước chính là tồn đầu tháng này

      for (let i = 1; i < dataQty.length; i++) {
        const row = dataQty[i];
        const p = String(row[idxQPeriod] || "").trim().replace(/\.0$/, '').replace(/[-\/]/g, '');
        if (p === prevPeriodClean || p === rawPrevPeriod) {
          const item = cleanCodeValue_(row[idxQItem]);
          const qty = idxQClose !== -1 ? (Number(row[idxQClose]) || 0) : 0;
          if (item) {
            mapOpeningQtyByItem[item] = (mapOpeningQtyByItem[item] || 0) + qty;
          }
        }
      }
    }

    // 2.2. Lấy tiền đầu kỳ gộp từ INVENTORY_VALUE_SUMMARY của tháng trước
    if (dataVal.length > 1) {
      const idxValPeriod   = getIdx(S_VAL, "period");
      const idxValItem     = getIdx(S_VAL, "item_code");
      const idxValCloseAmt = getIdx(S_VAL, "closing_amt");

      for (let i = 1; i < dataVal.length; i++) {
        const row = dataVal[i];
        const p = String(row[idxValPeriod] || "").trim().replace(/\.0$/, '').replace(/[-\/]/g, '');
        if (p === prevPeriodClean || p === rawPrevPeriod) {
          const item = cleanCodeValue_(row[idxValItem]);
          const amt = idxValCloseAmt !== -1 ? (Number(row[idxValCloseAmt]) || 0) : 0;
          if (item) {
            mapOpeningValByItem[item] = (mapOpeningValByItem[item] || 0) + amt;
          }
        }
      }
    }

    // Dự phòng quét LOCATION_ITEM_BALANCE nếu chưa có dữ liệu qua Summary
    if (Object.keys(mapOpeningQtyByItem).length === 0 && dataBal.length > 1) {
      const idxBalPeriod = getIdx(S_BAL, "period");
      const idxBalItem   = getIdx(S_BAL, "item_code");
      const idxBalQty    = getIdx(S_BAL, "opening_qty");
      const idxBalVal    = getIdx(S_BAL, "opening_value");

      for (let i = 1; i < dataBal.length; i++) {
        const row = dataBal[i];
        const p = String(row[idxBalPeriod] || "").trim().replace(/\.0$/, '').replace(/[-\/]/g, '');
        if (p === targetPeriodClean || p === prevPeriodClean || p === "") {
          const item = cleanCodeValue_(row[idxBalItem]);
          const qty = Number(row[idxBalQty]) || 0;
          const val = Number(row[idxBalVal]) || (qty * (Number(row[getIdx(S_BAL, "unit_cost")]) || 0));

          if (item) {
            mapOpeningQtyByItem[item] = (mapOpeningQtyByItem[item] || 0) + qty;
            mapOpeningValByItem[item] = (mapOpeningValByItem[item] || 0) + val;
          }
        }
      }
    }

    // 3. Quét TRANSACTION lấy lượng và giá trị nhập thực tế trong kỳ (Hỗ trợ hàng tặng giá trị bằng 0)
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
      if (transPeriod !== targetPeriodClean) continue;

      const itemCode = cleanCodeValue_(row[idxTransItem]);
      const fromCode = cleanCodeValue_(row[idxTransFromCode]);
      const toCode   = cleanCodeValue_(row[idxTransToCode]);
      const qty      = Number(row[idxTransQty]) || 0;
      const amount   = idxTransAmount !== -1 ? (Number(row[idxTransAmount]) || 0) : 0; // Nhận diện đúng 0 nếu hàng tặng

      if (!itemCode) continue;

      const fromType  = locationTypeMap.get(fromCode) || "UNKNOWN";
      const toType    = locationTypeMap.get(toCode) || "UNKNOWN";
      const routeRule = routeRulesMap.get(`${fromType}->${toType}`);

      if (!routeRule || !routeRule.isInventory) continue;

      if (toType === "PHYSICAL" && (routeRule.transactionType === "PURCHASE" || routeRule.transactionType === "INBOUND" || routeRule.transactionType === "TRANSFER")) {
        mapInboundQty[itemCode] = (mapInboundQty[itemCode] || 0) + qty;
        mapInboundAmt[itemCode] = (mapInboundAmt[itemCode] || 0) + amount; // Cộng dồn tiền nhập (có thể bằng 0)

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

    // 4. VÉT BỔ SUNG MÃ HÀNG
    const activeItemCodes = new Set(Object.keys(mapInboundQty));
    Object.keys(mapOpeningValByItem).forEach(k => activeItemCodes.add(k));
    Object.keys(mapItemMaster).forEach(k => {
      if (mapItemMaster[k].costPrice > 0) activeItemCodes.add(k);
    });

    if (dataQty.length > 1) {
      const idxQtyPeriod = getIdx(S_QTY, "period");
      const idxQtyItem   = getIdx(S_QTY, "item_code");

      for (let i = 1; i < dataQty.length; i++) {
        const row = dataQty[i];
        const qtyPeriod = String(row[idxQtyPeriod] || "").trim().replace(/\.0$/, '').replace(/[-\/]/g, '');
        if (qtyPeriod === targetPeriodClean) {
          const itemCode = cleanCodeValue_(row[idxQtyItem]);
          if (itemCode) activeItemCodes.add(itemCode);
        }
      }
    }

    // 5. Chuẩn bị headers và Upsert vào MONTHLY_PRICE_LIST
    const priceCols = schemaMap[S_PRICE].columns.sort((a, b) => a.colIndex - b.colIndex);
    const maxColIndex = Math.max(...priceCols.map(c => c.colIndex));
    const headers = new Array(maxColIndex);
    priceCols.forEach(c => {
      headers[c.colIndex - 1] = c.colHeader || c.colKey;
    });

    let currentPriceData = dataPrice;
    if (currentPriceData.length === 0 || currentPriceData[0][0] === "") {
      currentPriceData = [headers];
    }

    const keyIndexMap = {};
    const idxPricePeriod = getIdx(S_PRICE, "period");
    const idxPriceItem   = getIdx(S_PRICE, "item_code");

    for (let i = 1; i < currentPriceData.length; i++) {
      const p = String(currentPriceData[i][idxPricePeriod] || "").trim().replace(/\.0$/, '').replace(/[-\/]/g, '');
      const item = cleanCodeValue_(currentPriceData[i][idxPriceItem]);
      if (p && item) keyIndexMap[`${p}_${item}`] = i;
    }

    const nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

    // 6. TÍNH TOÁN BÌNH QUÂN GIA QUYỀN CHUẨN XÁC (Cộng đủ Lượng & Tiền đầu kỳ)
    activeItemCodes.forEach(itemCode => {
      if (!mapItemMaster[itemCode]) {
        mapItemMaster[itemCode] = { name: itemCode, category: "Chưa phân nhóm", unit: "N/A", costPrice: 0 };
      }

      const openQty = mapOpeningQtyByItem[itemCode] || 0; // Lượng tồn đầu kỳ
      const openVal = mapOpeningValByItem[itemCode] || 0; // Tiền tồn đầu kỳ

      const inboundQty = mapInboundQty[itemCode] || 0;
      const inboundAmt = mapInboundAmt[itemCode] || 0; // 0 nếu là hàng tặng

      // MẪU SỐ: Tổng lượng = Lượng đầu kỳ + Lượng nhập trong kỳ
      const totalQty = openQty + inboundQty;
      
      // TỬ SỐ: Tổng tiền = Tiền đầu kỳ + Tiền nhập trong kỳ
      const totalAmt = openVal + inboundAmt; 

      let unitPrice = 0;
      let priceSource = "";

      if (totalQty > 0) {
        // Đơn giá bình quân = Tổng tiền / Tổng lượng (pha loãng chính xác)
        unitPrice = totalAmt / totalQty; 
        priceSource = (openQty > 0 || inboundQty > 0) ? "WEIGHTED_AVG_WITH_OPENING" : "PURCHASE_AVG";
      } else if (mapItemMaster[itemCode].costPrice > 0) {
        unitPrice = mapItemMaster[itemCode].costPrice;
        priceSource = "FIXED_MASTER";
      } else {
        unitPrice = 0;
        priceSource = "ZERO_PRICE";
      }

      const master = mapItemMaster[itemCode];
      const dynamic = itemDynamicInfo[itemCode] || {};
      const itemName = master.name || itemCode;
      const baseUnit = master.unit || dynamic.unit || "N/A";
      const category = master.category || dynamic.category || "Chưa phân nhóm";

      const fullKey = `${targetPeriodClean}_${itemCode}`;
      const rowValues = new Array(maxColIndex).fill("");

      priceCols.forEach(col => {
        const colIdx = col.colIndex - 1;
        switch (col.colKey) {
          case "period":        rowValues[colIdx] = strPeriod; break;
          case "item_code":     rowValues[colIdx] = itemCode; break;
          case "item_name":     rowValues[colIdx] = itemName; break;
          case "base_unit":     rowValues[colIdx] = baseUnit; break;
          case "inbound_qty":   rowValues[colIdx] = inboundQty; break;
          case "inbound_amount":rowValues[colIdx] = inboundAmt; break; // Ghi nhận đúng 0 VNĐ tiền nhập
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

    sPrice.getRange(1, 1, currentPriceData.length, maxColIndex).setValues(currentPriceData);
    SpreadsheetApp.flush();
  }
}
