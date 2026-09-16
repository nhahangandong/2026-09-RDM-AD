/**
 * ============================================================================
 * MODULE: TheoreticalEngine.gs
 * Chuyên trách tính toán tiêu hao lý thuyết (Theoretical Usage) theo chuẩn UPSERT
 * ============================================================================
 */

class TheoreticalEngine {

  /**
   * Tính toán và tổng hợp tiêu hao lý thuyết cho một kỳ chỉ định theo chuẩn UPSERT
   * @param {string} targetPeriod - Kỳ tính toán (VD: "2026-08")
   * @param {boolean} isBatch - True nếu chạy hàng loạt (bỏ popup), False nếu chạy lẻ
   */
  static calculateUsage(targetPeriod, isBatch = false) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    const schemaMap = schemaGetMap();
    if (!schemaMap) {
      if (!isBatch) ui.alert("⚠️ Lỗi Schema", "Không thể đọc cấu trúc SCHEMA từ hệ thống!", ui.ButtonSet.OK);
      return;
    }
    
    // Lấy tên sheet từ schema
    const salesSheetName = schemaGetSheetName(schemaMap, "SALES") || "Bán hàng";
    const itemMasterSheetName = schemaGetSheetName(schemaMap, "ITEM_MASTER") || "ITEM_MASTER";
    const menuSheetName = schemaGetSheetName(schemaMap, "MENU") || "Thực đơn";
    const bomSheetName = schemaGetSheetName(schemaMap, "RECIPE_BOM") || "Định lượng";
    const priceSheetName = schemaGetSheetName(schemaMap, "MONTHLY_PRICE_LIST") || "MONTHLY_PRICE_LIST";
    const targetSheetName = schemaGetSheetName(schemaMap, "THEORETICAL_USAGE") || "Tiêu hao lý thuyết";
    
    const salesSheet = ss.getSheetByName(salesSheetName);
    const itemMasterSheet = ss.getSheetByName(itemMasterSheetName);
    const menuSheet = ss.getSheetByName(menuSheetName);
    const bomSheet = ss.getSheetByName(bomSheetName);
    const priceSheet = ss.getSheetByName(priceSheetName);
    let targetSheet = ss.getSheetByName(targetSheetName);
    
    if (!salesSheet || !itemMasterSheet || !bomSheet) {
      const errMsg = "Thiếu một trong các sheet cốt lõi (SALES, ITEM_MASTER, RECIPE_BOM)!";
      if (!isBatch) ui.alert("⚠️ Lỗi", errMsg, ui.ButtonSet.OK);
      throw new Error(errMsg);
    }
    
    if (!targetSheet) {
      targetSheet = ss.insertSheet(targetSheetName);
      const usageCols = schemaMap["THEORETICAL_USAGE"].columns.sort((a, b) => a.colIndex - b.colIndex);
      const headers = usageCols.map(c => c.colHeader || c.colKey);
      targetSheet.appendRow(headers);
    }

    const targetClean = String(targetPeriod).trim().replace(/\.0$/, '');

    // 1. Đọc cấu hình động từ SYSTEM_CONFIG cho module THEORETICAL
    const configMap = ConfigLoader.getConfigMap("THEORETICAL");
    const excludedItemTypes = configMap.get("EXCLUDED_ITEM_TYPES") || [];
    const excludedSet = new Set(excludedItemTypes.map(t => String(t).trim().toUpperCase()));

    // 2. Load Item Master
    const itemMasterData = itemMasterSheet.getDataRange().getValues();
    const idxImCode = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_code");
    const idxImName = schemaGetColIndex(schemaMap, "ITEM_MASTER", "item_name");
    const idxImUnit = schemaGetColIndex(schemaMap, "ITEM_MASTER", "base_unit");
    const idxImCat  = schemaGetColIndex(schemaMap, "ITEM_MASTER", "category");
    const idxImHasBom = schemaGetColIndex(schemaMap, "ITEM_MASTER", "has_bom");
    const idxImCostPrice = schemaGetColIndex(schemaMap, "ITEM_MASTER", "cost_price");

    const itemMasterMap = new Map();
    for (let i = 1; i < itemMasterData.length; i++) {
      const row = itemMasterData[i];
      const code = row[idxImCode] ? String(row[idxImCode]).trim() : "";
      if (code) {
        const hasBomVal = idxImHasBom !== -1 ? row[idxImHasBom] : false;
        const isTrue = (hasBomVal === true || String(hasBomVal).toUpperCase() === "TRUE" || String(hasBomVal) === "1" || hasBomVal === "CHECKED");
        const costPriceVal = idxImCostPrice !== -1 ? (Number(row[idxImCostPrice]) || 0) : 0;
        
        itemMasterMap.set(code, {
          name: idxImName !== -1 ? row[idxImName] : "",
          unit: idxImUnit !== -1 ? row[idxImUnit] : "",
          category: idxImCat !== -1 ? String(row[idxImCat]).trim() : "Chưa phân nhóm",
          hasBom: isTrue,
          costPrice: costPriceVal
        });
      }
    }

    // 3. Load Menu Map
    const menuMap = new Map();
    if (menuSheet && schemaMap["MENU"]) {
      const menuData = menuSheet.getDataRange().getValues();
      const idxMenuCode = schemaGetColIndex(schemaMap, "MENU", "menu_code");
      const idxMenuLinkedItem = schemaGetColIndex(schemaMap, "MENU", "item_code");
      const idxMenuItemType = schemaGetColIndex(schemaMap, "MENU", "item_type");

      for (let i = 1; i < menuData.length; i++) {
        const row = menuData[i];
        const mCode = row[idxMenuCode] ? String(row[idxMenuCode]).trim() : "";
        const lCode = idxMenuLinkedItem !== -1 && row[idxMenuLinkedItem] ? String(row[idxMenuLinkedItem]).trim() : "";
        const iType = idxMenuItemType !== -1 && row[idxMenuItemType] ? String(row[idxMenuItemType]).trim().toUpperCase() : "";

        if (mCode) {
          const isExcluded = excludedSet.has(iType);
          menuMap.set(mCode, {
            linkedItem: lCode,
            itemType: iType,
            isExcluded: isExcluded
          });
        }
      }
    }

    // 4. Load BOM Tree
    const bomData = bomSheet.getDataRange().getValues();
    const idxBomParent = schemaGetColIndex(schemaMap, "RECIPE_BOM", "parent_code");
    const idxBomChild = schemaGetColIndex(schemaMap, "RECIPE_BOM", "child_item_code");
    const idxBomQty = schemaGetColIndex(schemaMap, "RECIPE_BOM", "quantity");
    const idxBomWastage = schemaGetColIndex(schemaMap, "RECIPE_BOM", "wastage_rate");

    const bomTree = new Map();
    for (let i = 1; i < bomData.length; i++) {
      const row = bomData[i];
      const pCode = row[idxBomParent] ? String(row[idxBomParent]).trim() : "";
      const cCode = row[idxBomChild] ? String(row[idxBomChild]).trim() : "";
      const qty = Number(row[idxBomQty]) || 0;
      const wastage = idxBomWastage !== -1 ? (Number(row[idxBomWastage]) || 0) : 0;

      if (pCode && cCode && qty > 0) {
        if (!bomTree.has(pCode)) bomTree.set(pCode, []);
        bomTree.get(pCode).push({ childCode: cCode, normQty: qty, wastage: wastage });
      }
    }

    // 5. Đệ quy bóc tách BOM
    const explodeBom = (currentCode, accumulatedQty, resultRawMap) => {
      const imInfo = itemMasterMap.get(currentCode);
      const requiresBom = imInfo ? imInfo.hasBom : false;

      if (requiresBom) {
        if (bomTree.has(currentCode)) {
          const children = bomTree.get(currentCode);
          children.forEach(child => {
            const effectiveQty = accumulatedQty * child.normQty * (1 + child.wastage);
            explodeBom(child.childCode, effectiveQty, resultRawMap);
          });
        }
      } else {
        resultRawMap.set(currentCode, (resultRawMap.get(currentCode) || 0) + accumulatedQty);
      }
    };

    // 6. Đọc SALES và quy đổi tiêu hao
    const salesData = salesSheet.getDataRange().getValues();
    const idxSalesPeriod = schemaGetColIndex(schemaMap, "SALES", "period");
    const idxSalesItemCode = schemaGetColIndex(schemaMap, "SALES", "item_code");
    const idxSalesQty = schemaGetColIndex(schemaMap, "SALES", "quantity");
    const idxSalesDept = schemaGetColIndex(schemaMap, "SALES", "department");

    const rawUsageMap = new Map();

    for (let i = 1; i < salesData.length; i++) {
      const row = salesData[i];
      const period = idxSalesPeriod !== -1 ? String(row[idxSalesPeriod]).trim().replace(/\.0$/, '') : "";
      
      if (targetClean && period !== targetClean) continue;

      let soldItemCode = idxSalesItemCode !== -1 ? String(row[idxSalesItemCode]).trim() : "";
      const soldQty = Number(row[idxSalesQty]) || 0;
      const dept = (idxSalesDept !== -1 && row[idxSalesDept]) ? String(row[idxSalesDept]).trim() : "";

      if (!soldItemCode || soldQty <= 0) continue;

      if (menuMap.has(soldItemCode)) {
        const menuInfo = menuMap.get(soldItemCode);
        if (menuInfo.isExcluded) continue;
        if (menuInfo.linkedItem) soldItemCode = menuInfo.linkedItem;
      }

      const currentSalesRawMap = new Map();
      explodeBom(soldItemCode, soldQty, currentSalesRawMap);

      currentSalesRawMap.forEach((qtyVal, rawCode) => {
        const key = `${dept}|${rawCode}`;
        rawUsageMap.set(key, (rawUsageMap.get(key) || 0) + qtyVal);
      });
    }

    // 7. Load Price History
    const priceHistoryMap = new Map();
    if (priceSheet && schemaMap["MONTHLY_PRICE_LIST"]) {
      const priceData = priceSheet.getDataRange().getValues();
      const idxPricePeriod = schemaGetColIndex(schemaMap, "MONTHLY_PRICE_LIST", "period");
      const idxPriceItem = schemaGetColIndex(schemaMap, "MONTHLY_PRICE_LIST", "item_code");
      const idxPriceVal = schemaGetColIndex(schemaMap, "MONTHLY_PRICE_LIST", "unit_price");

      for (let i = 1; i < priceData.length; i++) {
        const row = priceData[i];
        const pPeriod = idxPricePeriod !== -1 ? String(row[idxPricePeriod]).trim().replace(/\.0$/, '') : "";
        const itemCode = idxPriceItem !== -1 ? String(row[idxPriceItem]).trim() : "";
        const unitPrice = idxPriceVal !== -1 ? (Number(row[idxPriceVal]) || 0) : 0;

        if (itemCode && pPeriod) {
          if (!priceHistoryMap.has(itemCode)) {
            priceHistoryMap.set(itemCode, new Map());
          }
          priceHistoryMap.get(itemCode).set(pPeriod, unitPrice);
        }
      }
    }

    const getResolvedPrice = (itemCode, currentPeriod) => {
      const itemPrices = priceHistoryMap.get(itemCode);
      if (itemPrices && itemPrices.size > 0) {
        let [year, month] = currentPeriod.includes('-') ? currentPeriod.split('-').map(Number) : [parseInt(currentPeriod.slice(0,4)), parseInt(currentPeriod.slice(4))];
        
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

      const imInfo = itemMasterMap.get(itemCode);
      if (imInfo && imInfo.costPrice > 0) return imInfo.costPrice;
      return 0;
    };

    // 8. XỬ LÝ THEO CHUẨN UPSERT THUẦN TÚY
    const usageCols = schemaMap["THEORETICAL_USAGE"].columns.sort((a, b) => a.colIndex - b.colIndex);
    const maxColCount = Math.max(...usageCols.map(c => c.colIndex));
    
    let currentData = targetSheet.getDataRange().getValues();
    if (currentData.length === 0 || currentData[0][0] === "") {
      const headers = usageCols.map(c => c.colHeader || c.colKey);
      currentData = [headers];
    }

    const idxUPeriod = schemaGetColIndex(schemaMap, "THEORETICAL_USAGE", "period");
    const idxUItem = schemaGetColIndex(schemaMap, "THEORETICAL_USAGE", "item_code");
    const idxUDept = schemaGetColIndex(schemaMap, "THEORETICAL_USAGE", "department_code");

    // Lập bản đồ (Map) định vị dòng dữ liệu cũ theo khóa [period + department + item_code] để UPSERT
    const keyIndexMap = {};
    for (let i = 1; i < currentData.length; i++) {
      const p = String(currentData[i][idxUPeriod] || "").trim().replace(/\.0$/, '');
      const dept = idxUDept !== -1 ? String(currentData[i][idxUDept] || "").trim() : "";
      const item = String(currentData[i][idxUItem] || "").trim();
      if (p && item) {
        keyIndexMap[`${p}_${dept}_${item}`] = i;
      }
    }

    let countUpdated = 0;
    let countInserted = 0;

    rawUsageMap.forEach((theoQty, keyStr) => {
      const parts = keyStr.split("|");
      const dept = parts[0];
      const itemCode = parts[1];
      const fullKey = `${targetClean}_${dept}_${itemCode}`;

      const imInfo = itemMasterMap.get(itemCode) || { name: itemCode, unit: "", category: "Chưa phân nhóm" };
      const unitPrice = getResolvedPrice(itemCode, targetClean);
      const theoVal = theoQty * unitPrice;

      const rowArr = new Array(maxColCount).fill("");
      usageCols.forEach(col => {
        const colIdx = col.colIndex - 1;
        const cKey = col.colKey;

        if (cKey === "period") rowArr[colIdx] = targetClean;
        else if (cKey === "department_code") rowArr[colIdx] = dept;
        else if (cKey === "item_code") rowArr[colIdx] = itemCode;
        else if (cKey === "item_name") rowArr[colIdx] = imInfo.name;
        else if (cKey === "base_unit") rowArr[colIdx] = imInfo.unit;
        else if (cKey === "category") rowArr[colIdx] = imInfo.category;
        else if (cKey === "theo_qy" || cKey === "theoretical_qty") rowArr[colIdx] = Math.round(theoQty * 1000) / 1000;
        else if (cKey === "unit_price") rowArr[colIdx] = Math.round(unitPrice * 100) / 100;
        else if (cKey === "theo_val" || cKey === "theoretical_amount") rowArr[colIdx] = Math.round(theoVal * 100) / 100;
      });

      if (keyIndexMap[fullKey] !== undefined) {
        // Cập nhật (Update) dòng hiện có
        currentData[keyIndexMap[fullKey]] = rowArr;
        countUpdated++;
      } else {
        // Thêm mới (Insert)
        currentData.push(rowArr);
        keyIndexMap[fullKey] = currentData.length - 1;
        countInserted++;
      }
    });

    // Ghi lại toàn bộ dữ liệu xuống sheet một lần duy nhất tối ưu hiệu năng
    targetSheet.clearContents();
    targetSheet.getRange(1, 1, currentData.length, maxColCount).setValues(currentData);
    SpreadsheetApp.flush();

    const summaryMsg = `Kỳ [${targetClean}]: Cập nhật ${countUpdated}, Thêm mới ${countInserted} dòng.`;
    if (isBatch) {
      ss.toast(summaryMsg, "TheoreticalEngine", 3);
    } else {
      ui.alert("✅ Upsert Thành Công", summaryMsg, ui.ButtonSet.OK);
    }
  }
}


