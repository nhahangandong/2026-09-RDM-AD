/**
 * ============================================================================
 * FILE: MenuUI.gs
 * MỤC ĐÍCH: Khai báo các hàm Entry Point phục vụ gán sự kiện gọi từ Menu UI.
 * Quy tắc đặt tên chuẩn: run + [Hành động] + [Thực thể] (CamelCase)
 * ============================================================================
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🛠️ QUẢN TRỊ HỆ THỐNG')
    // 1. TÁC VỤ VẬN HÀNH CHÍNH
    .addItem('📌 Gán mã SKU vào các sheet giao dịch', 'runAssignItemCodeToTransactions')
    .addItem('📝 Tự động phân loại, phân nhóm (CLASSIFICATION)', 'runAutoSuggestClassification')
    .addItem('🔄 Điền tất cả cột tính toán (Giao dịch & Kiểm kê)', 'runFillCalculatedColumnsAllSheets')
    .addSeparator()

    // 2. SUBMENU MAPPING
    .addSubMenu(ui.createMenu('🍁 1. Chuẩn hóa Mapping (Raw Data)')
      .addItem('1.1. Quét dữ liệu thô vào MAPPING', 'runSyncMappingFromTransactions')
      .addItem('1.2. Gợi ý item_name tự động (AUTO_MAP)', 'runAutoSuggestMappingNames')
      .addItem('1.3. Sinh mã SKU tự động cho MAPPING', 'runAutoGenerateSKUForMapping')
    )

    // 3. SUBMENU MASTER DATA
    .addSubMenu(ui.createMenu('🏷️ Danh mục & Master Data')
      .addItem('1. Đồng bộ ITEM_MASTER từ Giao dịch', 'runSyncItemMasterFromInventoryAndStocktake')
      .addItem('2. Đồng bộ Sales -> Menu', 'runSyncSalesToMenu')
      .addItem('3. Đồng bộ Menu -> Item Master', 'runSyncItemMasterFromMenu')
    )
    .addSeparator()

    // 4. SUBMENU QUẢN LÝ KHO & GIÁ (Đã lược bỏ batch đơn giá/giá trị không an toàn)
    .addSubMenu(ui.createMenu('📦 Quản lý Kho & Giá')
      .addItem('📥 Tạo/Kết chuyển số dư đầu kỳ', 'runOpeningBalanceDialog')
      .addSeparator()
      // A. Số lượng Kho (Cho phép Batch vì đứng đầu quy trình)
      .addItem('▶️ [Đơn lẻ] Tính Tồn kho Số lượng (1 kỳ)', 'runStockEngineQty')
      .addItem('🚀 [Batch] Tính Tồn kho Số lượng (Nhiều kỳ)', 'runBatchCalculateQtySummary')
      .addSeparator()
      // B. Đơn giá tháng (Chạy Đơn lẻ từng kỳ để kiểm soát chặt chẽ)
      .addItem('🔍 [Kiểm tra] Đơn giá Tức thì / Check NCC (Monthly Avg Price)', 'runMonthlyAvgPrice')
      .addItem('🔍 [Batch] Đơn giá mua bình quân tháng', 'runBatchMonthlyAvgPrice')
      .addItem('▶️ [Đơn lẻ] Chốt Đơn giá tháng chính thức (1 kỳ)', 'runMonthlyPriceList')
      .addSeparator()
      // C. Giá trị tồn kho (Chạy Đơn lẻ từng kỳ)
      .addItem('▶️ [Đơn lẻ] Tính Giá trị tồn kho (1 kỳ)', 'runStockEngineValue')
      .addSeparator()
      // D. Chuẩn nhất: Chạy toàn trình tuần tự cho 1 kỳ (Qty -> Price -> Value)
      .addItem('⚡ [Toàn trình 1 Kỳ] Chạy Full (Qty -> Price -> Value)', 'runBatchPipelineSinglePeriod')
    )
    .addSeparator()

    // 5. SUBMENU ĐỊNH LƯỢNG MÓN ĂN & ĐỒ UỐNG
    .addSubMenu(ui.createMenu('🍳 Quản lý BOM')
      .addItem(' Tạo bảng định lượng', 'runInitRecipeBom')
      .addItem(' Điền tên món & nguyên liệu', 'runPopulateRecipeBomNames')
      .addItem(' Làm sạch dữ liệu BOM (xóa trùng lặp)', 'runCleanRecipeBomDuplicates')
      .addSeparator()
      .addItem(' Tính nhanh chi phí cost', 'runCalculateMenuCost')
      .addItem(' Lưu dữ liệu chi phí cost theo kỳ', 'runSaveMenuCostSnapshot')
    )
    .addSeparator()

    // 6. SUBMENU PHÂN TÍCH GIÁ VÀ ĐỐI CHIẾU
    .addSubMenu(ui.createMenu('📈 Phân tích & Đối chiếu')
      .addItem('🔄 Tính tiêu hao lý thuyết', 'runCalculateTheoreticalUsage')
    )
    .addSeparator()
    
    // 7. SUBMENU QUẢN LÝ CHI PHÍ (EXPENSE)
    .addSubMenu(ui.createMenu('💰 Quản lý Chi Phí (EXPENSE)')
      .addItem('🔄 Tổng hợp EXPENSE từ Transaction', 'runExpenseSyncFromTransaction')
    )
    .addSeparator()
    
    // 8. SUBMENU KHỞI TẠO & HỆ THỐNG
    .addSubMenu(ui.createMenu('⚙️ Thiết lập hệ thống')
      .addItem('🌱 Khởi tạo / Cập nhật Sheet từ SCHEMA', 'setupSheetsFromSchema')
    )

    .addToUi();
}


// ============================================================================
// ENTRY POINTS: CÁC HÀM GỌI CHỨC NĂNG
// ============================================================================

function runSyncItemMasterFromInventoryAndStocktake() { itemMasterSyncFromInventoryAndStocktake(); }
function runSyncItemMasterFromMenu() { itemMasterSyncFromMenu(); }
function runAutoSuggestClassification() { itemMasterAutoSuggestClassification(); }
function runAssignItemCodeToTransactions() { transactionAssignItemCodeToAllSheets(); }
function runSyncSalesToMenu() { productSyncSalesToMenu(); }
function runFillCalculatedColumnsAllSheets() {
  try {
    transactionFillCalculatedColumnsAllSheets();
    SpreadsheetApp.getActiveSpreadsheet().toast("✅ Đã tự động tính toán và điền dữ liệu thành công!", "Thành công");
  } catch (err) {
    SpreadsheetApp.getUi().alert("❌ Lỗi thực thi: " + err.message);
  }
}

/**
 * 1. QUẢN LÝ SỐ LƯỢNG KHO (Đơn lẻ & Batch nhiều kỳ)
 */
function runStockEngineQty() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('🔄 Tính Tồn Kho Số Lượng (Đơn Lẻ)', 'Nhập Kỳ báo cáo (Định dạng YYYY-MM, VD: 2026-03):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const periodTarget = response.getResponseText().trim();

  if (!/^\d{4}-\d{2}$/.test(periodTarget)) {
    ui.alert('⚠️ Lỗi Định Dạng', 'Vui lòng nhập đúng định dạng YYYY-MM.', ui.ButtonSet.OK);
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Đang xử lý tồn kho số lượng kỳ ${periodTarget}...`, 'Hệ thống');
    SpreadsheetApp.flush();
    StockEngine.calculateQtySummary(periodTarget);
    SpreadsheetApp.flush();
    ui.alert('✅ Thành Công', `Đã hoàn tất tính số lượng kho kỳ ${periodTarget}.`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Lỗi Xử Lý', error.message, ui.ButtonSet.OK);
  }
}

function runBatchCalculateQtySummary() {
  const ui = SpreadsheetApp.getUi();
  const startRes = ui.prompt('🔄 Batch Số Lượng Kho', 'Nhập KỲ BẮT ĐẦU (YYYY-MM):', ui.ButtonSet.OK_CANCEL);
  if (startRes.getSelectedButton() !== ui.Button.OK) return;
  const startPeriod = startRes.getResponseText().trim();

  const endRes = ui.prompt('🔄 Batch Số Lượng Kho', 'Nhập KỲ KẾT THÚC (YYYY-MM):', ui.ButtonSet.OK_CANCEL);
  if (endRes.getSelectedButton() !== ui.Button.OK) return;
  const endPeriod = endRes.getResponseText().trim();

  const periods = generatePeriodRange_(startPeriod, endPeriod);
  if (periods.length === 0) {
    ui.alert('⚠️ Lỗi Khoảng Thời Gian', 'Khoảng thời gian không hợp lệ!', ui.ButtonSet.OK);
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Đang chạy batch số lượng ${periods.length} kỳ...`, 'Hệ thống');
    for (let i = 0; i < periods.length; i++) {
      StockEngine.calculateQtySummary(periods[i]);
    }
    SpreadsheetApp.flush();
    ui.alert('✅ Hoàn Tất', `Đã chạy xong số lượng kho từ ${startPeriod} đến ${endPeriod}.`, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('❌ Lỗi Batch', err.message, ui.ButtonSet.OK);
  }
}


/**
 * 2. QUẢN LÝ ĐƠN GIÁ (Chỉ chạy Đơn lẻ cho từng kỳ để bảo đảm tính chính xác phụ thuộc)
 */
function runMonthlyAvgPrice() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('🔍 Kiểm tra Đơn giá Tức thì / Check NCC', 'Nhập kỳ tính toán (Định dạng YYYY-MM):', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const period = response.getResponseText().trim();
    if (!period) return;

    const periodRegex = /^\d{4}-(?:0[1-9]|1[0-2])$/;
    if (!periodRegex.test(period)) {
      ui.alert('Lỗi: Định dạng kỳ không hợp lệ! Vui lòng nhập theo chuẩn YYYY-MM.');
      return;
    }

    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Đang chạy kiểm tra đơn giá tức thì kỳ ${period}...`, 'Hệ thống');
      AvgPriceEngine.monthly_avg_price(period);
      ui.alert(`✅ Đã cập nhật bảng MONTHLY_AVG_PRICE cho kỳ ${period}.`);
    } catch (err) {
      ui.alert(`❌ Lỗi: ${err.message}`);
    }
  }
}


/**
 * Hàm giao diện Entry Point chạy hàng loạt Đơn giá bình quân tháng cho nhiều kỳ
 */
function runBatchMonthlyAvgPrice() {
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    console.warn("Hàm này cần được chạy trực tiếp từ giao diện Google Sheets.");
    return;
  }
  
  const response = ui.prompt(
    '🔄 Chạy Hàng Loạt Đơn Giá Bình Quân Tháng', 
    'Nhập khoảng kỳ cần chạy theo định dạng [Kỳ Bắt Đầu] đến [Kỳ Kết Thúc]\n(Ví dụ: 2026-03 đến 2026-08):', 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const inputVal = response.getResponseText().trim();
  if (!inputVal) return;

  const periodRegex = /^\d{4}-(?:0[1-9]|1[0-2])$/;
  let periodsToRun = [];

  // Tìm tất cả các định dạng YYYY-MM xuất hiện trong chuỗi người dùng nhập
  const matches = inputVal.match(/\d{4}-\d{2}/g);

  if (matches && matches.length >= 2) {
    // Nếu tìm thấy từ 2 mốc thời gian trở lên (Ví dụ: 2026-03 và 2026-08) -> Hiểu là chạy một khoảng
    let [yStart, mStart] = matches[0].split('-').map(Number);
    let [yEnd, mEnd] = matches[1].split('-').map(Number);
    
    let currentY = yStart;
    let currentM = mStart;
    
    while (currentY < yEnd || (currentY === yEnd && currentM <= mEnd)) {
      periodsToRun.push(`${currentY}-${String(currentM).padStart(2, '0')}`);
      currentM++;
      if (currentM > 12) {
        currentM = 1;
        currentY++;
      }
    }
  } else if (matches && matches.length === 1) {
    // Nếu chỉ nhập 1 kỳ duy nhất đúng chuẩn
    if (periodRegex.test(matches[0])) {
      periodsToRun.push(matches[0]);
    }
  }

  // Fallback kiểm tra thô nếu người dùng gõ trực tiếp 1 kỳ đơn lẻ
  if (periodsToRun.length === 0 && periodRegex.test(inputVal)) {
    periodsToRun.push(inputVal);
  }

  if (periodsToRun.length === 0) {
    ui.alert('❌ Lỗi: Không nhận diện được định dạng kỳ hợp lệ! Vui lòng nhập theo chuẩn YYYY-MM (Ví dụ: 2026-03 đến 2026-06).');
    return;
  }

  let successCount = 0;
  let errorLog = [];

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    for (let i = 0; i < periodsToRun.length; i++) {
      const period = periodsToRun[i];
      ss.toast(`Đang xử lý kỳ ${period} (${i + 1}/${periodsToRun.length})...`, 'Hệ thống Batch', 5);
      
      try {
        AvgPriceEngine.monthly_avg_price(period);
        successCount++;
      } catch (err) {
        errorLog.push(`Kỳ ${period}: ${err.message}`);
      }
    }
    
    SpreadsheetApp.flush();
    
    let msg = `✅ Đã chạy hoàn tất hàng loạt!\n- Thành công: ${successCount}/${periodsToRun.length} kỳ.`;
    if (errorLog.length > 0) {
      msg += `\n- Có lỗi phát sinh:\n` + errorLog.join('\n');
    }
    ui.alert('Kết quả Chạy Batch MONTHLY_AVG_PRICE', msg, ui.ButtonSet.OK);

} catch (err) {
    ui.alert(`❌ Lỗi hệ thống Batch: ${err.message}`);
  }
}



function runMonthlyPriceList() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('💰 Chốt Giá Tháng Chính Thức (Đơn Lẻ)', 'Nhập kỳ tính toán (Định dạng YYYY-MM):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const period = response.getResponseText().trim();

  if (!/^\d{4}-\d{2}$/.test(period)) {
    ui.alert('Lỗi: Định dạng kỳ không hợp lệ! Dùng chuẩn YYYY-MM.');
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Đang tính đơn giá tháng cho kỳ ${period}...`, 'Hệ thống');
    AvgPriceEngine.monthly_price_list(period);
    ui.alert(`✅ Đã hoàn tất chốt giá tháng và cập nhật Balance cho kỳ ${period}.`);
  } catch (err) {
    ui.alert(`❌ Lỗi: ${err.message}`);
  }
}


/**
 * 3. QUẢN LÝ GIÁ TRỊ TỒN KHO (Chạy Đơn lẻ cho từng kỳ)
 */
function runStockEngineValue() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('📦 Tính Giá Trị Tồn Kho (Đơn Lẻ)', 'Nhập kỳ tính toán (Định dạng YYYY-MM):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const period = response.getResponseText().trim();

  if (!/^\d{4}-\d{2}$/.test(period)) {
    ui.alert('Lỗi: Định dạng kỳ không hợp lệ! Dùng chuẩn YYYY-MM.');
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Đang tính giá trị tồn kho kỳ ${period}...`, 'Hệ thống');
    StockEngine.calculateValueSummary(period);
    ui.alert(`✅ Đã cập nhật bảng INVENTORY_VALUE_SUMMARY cho kỳ ${period}.`);
  } catch (err) {
    ui.alert(`❌ Lỗi: ${err.message}`);
  }
}


/**
 * 4. CHẠY TOÀN TRÌNH PIPELINE (CHO 1 KỲ: QTY -> PRICE -> VALUE)
 */
function runBatchPipelineSinglePeriod() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('⚡ Chạy Pipeline Full (1 Kỳ)', 'Nhập kỳ tính toán (Định dạng YYYY-MM):', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const period = response.getResponseText().trim();
    if (!/^\d{4}-(?:0[1-9]|1[0-2])$/.test(period)) {
      ui.alert('Lỗi: Định dạng kỳ không hợp lệ! Vui lòng dùng YYYY-MM.');
      return;
    }

    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Đang chạy chuỗi toàn trình cho kỳ ${period}...`, 'Hệ thống');
      StockEngine.calculateQtySummary(period);
      AvgPriceEngine.monthly_price_list(period);
      StockEngine.calculateValueSummary(period);

      ui.alert(`✅ Thành công! Đã chạy chuẩn chuỗi (Qty -> Price -> Value) cho kỳ ${period}.`);
    } catch (err) {
      ui.alert(`❌ Đã xảy ra lỗi: ${err.message}`);
    }
  }
}


// ============================================================================
// ENTRY POINTS CÁC MODUL KHÁC
// ============================================================================

function runInitRecipeBom() { recipeBomInitFromItemMaster(); }
function runPopulateRecipeBomNames() { recipeBomPopulateNames(); }
function runCleanRecipeBomDuplicates() { recipeBomCleanDuplicates(); }

function runCalculateMenuCost() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("📊 Tính Chi Phí Món Theo Kỳ", "Nhập kỳ tra cứu (VD: 2026-03):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    const period = response.getResponseText().trim();
    if (period) menuCalculateCost(period);
    else ui.alert("⚠️ Cảnh báo", "Kỳ không được để trống!", ui.ButtonSet.OK);
  }
}

function runSaveMenuCostSnapshot() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("📸 Chốt Snapshot Cost", "Nhập kỳ báo cáo (VD: 2026-03):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    const period = response.getResponseText().trim();
    if (period) menuSaveCostSnapshot(period);
    else ui.alert("⚠️ Cảnh báo", "Kỳ không được để trống!", ui.ButtonSet.OK);
  }
}

function runExpenseSyncFromTransaction() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("💰 Đồng Bộ Chi Phí (EXPENSE)", "Nhập kỳ hạch toán (YYYY-MM):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    const period = response.getResponseText().trim();
    if (!/^\d{4}-\d{2}$/.test(period)) {
      ui.alert("⚠️ Lỗi Định Dạng", "Dùng chuẩn YYYY-MM.", ui.ButtonSet.OK);
      return;
    }
    try {
      expenseSyncFromTransaction(period);
      ui.alert("✅ Thành công", `Đã đồng bộ EXPENSE kỳ ${period}`, ui.ButtonSet.OK);
    } catch (error) {
      ui.alert("❌ Lỗi", error.message, ui.ButtonSet.OK);
    }
  }
}

function runCalculateTheoreticalUsage() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("📊 Tính Tiêu Hao Lý Thuyết", "Nhập kỳ báo cáo (YYYY-MM):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    const period = response.getResponseText().trim();
    if (!/^\d{4}-\d{2}$/.test(period)) {
      ui.alert("⚠️ Lỗi Định Dạng", "Dùng chuẩn YYYY-MM.", ui.ButtonSet.OK);
      return;
    }
    try {
      TheoreticalEngine.calculateUsage(period);
    } catch (err) {
      ui.alert("❌ Lỗi", err.message, err.message, ui.ButtonSet.OK);
    }
  }
}

function runOpeningBalanceDialog() {
  const ui = SpreadsheetApp.getUi();
  const targetResponse = ui.prompt('Bước 1/2: Kỳ đích', 'Nhập kỳ nhận số dư (YYYY-MM):', ui.ButtonSet.OK_CANCEL);
  if (targetResponse.getSelectedButton() !== ui.Button.OK) return;
  const targetPeriod = targetResponse.getResponseText().trim();

  const sourceResponse = ui.prompt('Bước 2/2: Kỳ nguồn', 'Nhập kỳ nguồn lấy số liệu (YYYY-MM):', ui.ButtonSet.OK_CANCEL);
  if (sourceResponse.getSelectedButton() !== ui.Button.OK) return;
  const sourcePeriod = sourceResponse.getResponseText().trim();

  try {
    StockEngine.generateOpeningBalance(targetPeriod, sourcePeriod);
    ui.alert("✅ Thành công", `Đã kết chuyển số dư từ ${sourcePeriod} sang ${targetPeriod}`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Lỗi', error.message, ui.ButtonSet.OK);
  }
}
