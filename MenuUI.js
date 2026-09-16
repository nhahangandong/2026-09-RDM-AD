/**
 * ============================================================================
 * FILE: MenuUI.gs
 * MỤC ĐÍCH: Khai báo các hàm Entry Point phục vụ gán sự kiện gọi từ Menu UI.
 * Quy tắc đặt tên: actionEntityDescription (CamelCase)
 * ============================================================================
 */


 /**
 * Simple Trigger tự động tạo Menu khi mở file Google Sheets.
 * Thiết kế phân cấp Submenu theo quy trình chuẩn quản trị.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🛠️ QUẢN TRỊ HỆ THỐNG')
    // 1. TÁC VỤ VẬN HÀNH CHÍNH (Được gọi nhiều nhất)
    .addItem('📌 Gán mã SKU vào các sheet giao dịch', 'runAssignItemCodeToTransactions')
    .addItem('📝 Tự động phân loại, phân nhóm (CLASSIFICATION)', 'runAutoSuggestClassification')
    .addItem('🔄 Điền tất cả cột tính toán (Giao dịch & Kiểm kê)', 'runFillCalculatedColumnsAllSheets')
    .addSeparator()

    // 2. SUBMENU MAPPING (Chuẩn hóa tên thô)
    .addSubMenu(ui.createMenu('🍁 1. Chuẩn hóa Mapping (Raw Data)')
      .addItem('1.1. Quét dữ liệu thô vào MAPPING', 'runSyncMappingFromTransactions')
      .addItem('1.2. Gợi ý item_name tự động (AUTO_MAP)', 'runAutoSuggestMappingNames')
      .addItem('1.3. Sinh mã SKU tự động cho MAPPING', 'runAutoGenerateSKUForMapping')
    )

    // 3. SUBMENU MASTER DATA (Đã bổ sung Engine Phân loại SKU)
    .addSubMenu(ui.createMenu('🏷️ Danh mục & Master Data')
      .addItem('1. Đồng bộ ITEM_MASTER từ Giao dịch', 'runSyncItemMasterFromInventoryAndStocktake')
      .addItem('2. Đồng bộ Sales -> Menu', 'runSyncSalesToMenu')
      .addItem('3. Đồng bộ Menu -> Item Master', 'runSyncItemMasterFromMenu')
    )
    .addSeparator()

    // 4. SUBMENU TINH TỒN KHO
    .addSubMenu(ui.createMenu('📦 Quản lý kho')
      .addItem('▶️ Chạy tính Tồn kho Số lượng (Cho một kỳ)', 'runStockEngineQty')
      .addItem('🚀 Chạy tính Tồn kho Số lượng (Cho nhiều kỳ)', 'runBatchCalculateQtySummary')
      .addSeparator()
      .addItem('🚀 Chạy Batch Tổng hợp Kỳ (Qty + Price)', 'uiRunBatchPipeline')
      .addSeparator()
      .addItem('🔍 Kiểm tra Đơn giá Tức thì (Monthly Avg Price)', 'uiRunMonthlyAvgPrice')
      .addItem('📝 Chốt Giá Tháng Chính Thức (Monthly Price List)', 'uiRunMonthlyPriceList')
      .addItem('🚀 Chạy Batch Đơn Giá Tháng', 'runBatchMonthlyPriceList')
      .addSeparator()
      .addItem('▶️ Tính giá trị tồn kho (một kỳ)', 'runStockEngineValue')
      .addItem('🚀 Tính giá trị tồn kho (nhiều kỳ)', 'runBatchStockEngineValueRange')
    )
    .addSeparator()

    // 5. SUBMENU ĐỊNH LƯỢNG MÓN ĂN & ĐỒ UỐNG
    .addSubMenu(ui.createMenu('🍳 Quản lý BOM')
      .addItem(' Tạo bảng định lượng', 'runInitRecipeBom')
      .addItem(' Điền tên món & nguyên liệu', 'runPopulateRecipeBomNames')
      .addItem(' Làm sạch dữ liệu BOM (xóa trùng lặp)', 'runCleanRecipeBomDuplicates')
      .addSeparator()
      .addItem(' Tính nhanh chi phí cost', 'runCalculateMenuCost')
      .addItem(' Lữu dữ liệu chi phí cost theo kỳ', 'runSaveMenuCostSnapshot')
    )
    .addSeparator()

    // 6. SUBMENU PHÂN TÍCH GIÁ VÀ ĐỐI CHIẾU
    .addSubMenu(ui.createMenu('📈 Phân tích & Đối chiếu')
      .addItem('🔄 Tính tiêu hao lý thuyết', 'runCalculateTheoreticalUsage')
    )
    .addSeparator()

    
    // 7. SUBMENU QUẢN LÝ CHI PHÍ (EXPENSE) - MỚI THÊM
    .addSubMenu(ui.createMenu('💰 Quản lý Chi Phí (EXPENSE)')
      .addItem('🔄 Tổng hợp EXPENSE từ Transaction', 'runExpenseSyncFromTransaction')
    )
    .addSeparator()
    
    // 8. SUBMENU KHỞI TẠO & HỆ THỐNG
    .addSubMenu(ui.createMenu('⚙️ Thiết lập hệ thống')
      .addItem('🌱  Khởi tạo / Cập nhật Sheet từ SCHEMA', 'setupSheetsFromSchema')
    )

    .addToUi();
}



/**
 * Entry Point: Đồng bộ SKU từ Transaction và Stocktake vào Item Master.
 */
function runSyncItemMasterFromInventoryAndStocktake() {
  itemMasterSyncFromInventoryAndStocktake();
}

/**
 * Entry Point: Đồng bộ Mã món từ Menu sang Item Master.
 */
function runSyncItemMasterFromMenu() {
  itemMasterSyncFromMenu();
}

/**
 * Entry Point gợi ý phân loại tự động từ Menu UI / Button
 * Cú pháp: actionEntityDescription (Tiền tố 'run')
 */
function runAutoSuggestClassification() {
  itemMasterAutoSuggestClassification();
}

/**
 * Entry Point: Tra cứu và điền mã SKU (item_code) từ sheet MAPPING quay trở lại các sheet giao dịch.
 */
function runAssignItemCodeToTransactions() {
  transactionAssignItemCodeToAllSheets();
}

/**
 * Entry Point: đồng bộ item_name & item_code từ Sales to Menu
 */
function runSyncSalesToMenu() {
  productSyncSalesToMenu();
}


/**
 * File: MainController.gs
 * Điều phối luồng chạy toàn bộ quy trình chốt kho tháng
 */
function runMonthlyInventoryProcess() {
  const ui = SpreadsheetApp.getUi();
  const promptKy = ui.prompt("Chốt kho tháng", "Nhập kỳ YYYYMM (Ví dụ: 202608):", ui.ButtonSet.OK_CANCEL);
  if (promptKy.getSelectedButton() !== ui.Button.OK) return;
  
  const period = promptKy.getResponseText().trim();
  
  try {
    // Bước 1: Tính tồn kho số lượng
    Logger.log(`[1/3] Đang tính INVENTORY_QTY_SUMMARY cho kỳ ${period}...`);
    StockEngine.calculateQtySummary(period);

    // Bước 2: Tính giá bình quân tháng
    Logger.log(`[2/3] Đang tính MONTHLY_AVG_PRICE cho kỳ ${period}...`);
    StockEngine.calculateMonthlyAvgPrice(period);

    // Bước 3: Tính tồn kho giá trị
    Logger.log(`[3/3] Đang tính INVENTORY_VALUE_SUMMARY cho kỳ ${period}...`);
    StockEngine.calculateValueSummary(period);

    ui.alert(`✅ Đã chốt thành công toàn bộ số liệu kho & giá trị kỳ ${period}!`);
  } catch (error) {
    Logger.log(`❌ Lỗi quy trình chốt kho: ${error.stack}`);
    ui.alert(`❌ LỖI QUY TRÌNH: ${error.message}`);
  }
}


function runStockEngineQty() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. Hiển thị hộp thoại yêu cầu nhập kỳ báo cáo
  const response = ui.prompt(
    '🔄 Tính Toán Tồn Kho Số Lượng',
    'Vui lòng nhập Kỳ báo cáo cần tính toán (Định dạng: YYYY-MM, VD: 2026-03):',
    ui.ButtonSet.OK_CANCEL
  );

  // 2. Xử lý khi người dùng nhấn OK
  if (response.getSelectedButton() === ui.Button.OK) {
    const periodTarget = response.getResponseText().trim();

    // Validate định dạng YYYY-MM
    if (!/^\d{4}-\d{2}$/.test(periodTarget)) {
      ui.alert(
        '⚠️ Lỗi Định Dạng', 
        'Kỳ báo cáo không hợp lệ. Vui lòng nhập đúng định dạng YYYY-MM (VD: 2026-03).', 
        ui.ButtonSet.OK
      );
      return;
    }

    // 3. Thực thi Engine
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Đang xử lý dữ liệu tồn kho kỳ ${periodTarget}...`, 'Hệ thống');

      // Ép hệ thống đồng bộ các thay đổi trên Sheet trước khi đọc
      SpreadsheetApp.flush();

      // Gọi hàm Core của StockEngine
      StockEngine.calculateQtySummary(periodTarget);

      // Đẩy dữ liệu mới tính toán xuống Sheet ngay lập tức
      SpreadsheetApp.flush();

      ui.alert(
        '✅ Thành Công', 
        `Đã hoàn tất tính toán và cập nhật số liệu INVENTORY_QTY_SUMMARY cho kỳ ${periodTarget}.`, 
        ui.ButtonSet.OK
      );
    } catch (error) {
      ui.alert(
        '❌ Lỗi Xử Lý', 
        `Đã xảy ra lỗi trong quá trình tính toán:\n\n${error.message}\n\nVui lòng kiểm tra lại cấu trúc SCHEMA hoặc dữ liệu đầu vào.`, 
        ui.ButtonSet.OK
      );
    }
  }
}


/**
 * Controller bắt sự kiện từ Menu UI
 */
function runFillCalculatedColumnsAllSheets() {
  try {
    transactionFillCalculatedColumnsAllSheets();
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "✅ Đã tự động tính toán và điền dữ liệu cho Giao dịch & Kiểm kê!",
      "Thành công"
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert("❌ Lỗi thực thi: " + err.message);
  }
}

// Khai báo trong Menu custom của Apps Script:
// .addItem('🔄 Điền tất cả cột tính toán (Giao dịch & Kiểm kê)', 'runFillCalculatedColumnsAllSheets')


/**
 * Kịch bản chạy hàng loạt (Batch) cho nhiều tháng liên tục
 * Nhập kỳ bắt đầu và kỳ kết thúc, tự động kiểm tra định dạng và chạy tuần tự.
 */
function runBatchCalculateQtySummary() {
  const ui = SpreadsheetApp.getUi();

  // 1. Nhập kỳ bắt đầu
  const startResponse = ui.prompt(
    '🔄 Chạy Tồn Kho Hàng Loạt', 
    'Nhập KỲ BẮT ĐẦU (Định dạng YYYY-MM, VD: 2026-01):', 
    ui.ButtonSet.OK_CANCEL
  );
  if (startResponse.getSelectedButton() !== ui.Button.OK) return;
  const startPeriod = startResponse.getResponseText().trim();

  if (!/^\d{4}-\d{2}$/.test(startPeriod)) {
    ui.alert('⚠️ Lỗi Định Dạng', 'Kỳ bắt đầu không đúng định dạng YYYY-MM.', ui.ButtonSet.OK);
    return;
  }

  // 2. Nhập kỳ kết thúc
  const endResponse = ui.prompt(
    '🔄 Chạy Tồn Kho Hàng Loạt', 
    'Nhập KỲ KẾT THÚC (Định dạng YYYY-MM, VD: 2026-06):', 
    ui.ButtonSet.OK_CANCEL
  );
  if (endResponse.getSelectedButton() !== ui.Button.OK) return;
  const endPeriod = endResponse.getResponseText().trim();

  if (!/^\d{4}-\d{2}$/.test(endPeriod)) {
    ui.alert('⚠️ Lỗi Định Dạng', 'Kỳ kết thúc không đúng định dạng YYYY-MM.', ui.ButtonSet.OK);
    return;
  }

  // 3. Tạo danh sách các kỳ theo trình tự thời gian
  const periods = generatePeriodRange_(startPeriod, endPeriod);
  if (periods.length === 0) {
    ui.alert('⚠️ Lỗi Khoảng Thời Gian', 'Kỳ bắt đầu lớn hơn kỳ kết thúc hoặc khoảng thời gian không hợp lệ!', ui.ButtonSet.OK);
    return;
  }

  // 4. Thực thi tuần tự từng kỳ (Bắt buộc chạy đúng chiều thời gian để bảo toàn tồn đầu kỳ)
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Đang xử lý tổng số ${periods.length} kỳ từ ${startPeriod} đến ${endPeriod}...`, 'Hệ thống');
    SpreadsheetApp.flush();
    
    for (let i = 0; i < periods.length; i++) {
      const targetPeriod = periods[i];
      StockEngine.calculateQtySummary(targetPeriod);
    }

    SpreadsheetApp.flush();
    ui.alert(
      '✅ Hoàn Tất Thành Công', 
      `Đã tính toán xong dữ liệu tồn kho cho toàn bộ các kỳ:\n${periods.join(', ')}`, 
      ui.ButtonSet.OK
    );

  } catch (err) {
    ui.alert(
      '❌ Lỗi Thực Thi Hàng Loạt', 
      `Đã xảy ra lỗi tại một kỳ trong chuỗi xử lý:\n\n${err.message}`, 
      ui.ButtonSet.OK
    );
  }
}



/**
 * Giao diện nhập kỳ và chạy toàn bộ chuỗi Pipeline tự động
 */
function uiRunBatchPipeline() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Chạy Batch Tổng hợp Kho', 'Nhập kỳ tính toán (Định dạng YYYY-MM):', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const period = response.getResponseText().trim();
    if (!period) {
      ui.alert('Vui lòng nhập kỳ hợp lệ!');
      return;
    }

    // Quy tắc xác thực định dạng kỳ (YYYY-MM)
    const periodRegex = /^\d{4}-(?:0[1-9]|1[0-2])$/;
    if (!periodRegex.test(period)) {
      ui.alert('Lỗi: Định dạng kỳ không hợp lệ! Vui lòng nhập theo chuẩn YYYY-MM (Ví dụ: 2026-08).');
      return;
    }

    try {
      StockEngine.calculateQtySummary(period);
      AvgPriceEngine.monthly_avg_price(period);
      AvgPriceEngine.monthly_price_list(period);

      ui.alert(`Thành công! Đã hoàn tất xử lý toàn bộ dữ liệu cho kỳ ${period}.`);
    } catch (err) {
      ui.alert(`Đã xảy ra lỗi: ${err.message}`);
    }
  }
}


/**
 * Giao diện chạy độc lập hàm kiểm tra đơn giá tức thì
 */
function uiRunMonthlyAvgPrice() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Kiểm tra Đơn giá Tức thì', 'Nhập kỳ tính toán (Định dạng YYYY-MM):', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const period = response.getResponseText().trim();
    if (!period) return;

    // Quy tắc xác thực định dạng kỳ (YYYY-MM)
    const periodRegex = /^\d{4}-(?:0[1-9]|1[0-2])$/;
    if (!periodRegex.test(period)) {
      ui.alert('Lỗi: Định dạng kỳ không hợp lệ! Vui lòng nhập theo chuẩn YYYY-MM (Ví dụ: 2026-08).');
      return;
    }

    try {
      AvgPriceEngine.monthly_avg_price(period);
      ui.alert(`Đã cập nhật bảng MONTHLY_AVG_PRICE cho kỳ ${period}.`);
    } catch (err) {
      ui.alert(`Lỗi: ${err.message}`);
    }
  }
}

/**
 * Giao diện chạy độc lập hàm chốt giá tháng chính thức
 */
function uiRunMonthlyPriceList() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Chốt Giá Tháng Chính Thức', 'Nhập kỳ tính toán (Định dạng YYYY-MM):', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const period = response.getResponseText().trim();
    if (!period) return;

    // Quy tắc xác thực định dạng kỳ (YYYY-MM)
    const periodRegex = /^\d{4}-(?:0[1-9]|1[0-2])$/;
    if (!periodRegex.test(period)) {
      ui.alert('Lỗi: Định dạng kỳ không hợp lệ! Vui lòng nhập theo chuẩn YYYY-MM (Ví dụ: 2026-08).');
      return;
    }

    try {
      AvgPriceEngine.monthly_price_list(period);
      ui.alert(`Đã hoàn tất chốt giá tháng và cập nhật Balance cho kỳ ${period}.`);
    } catch (err) {
      ui.alert(`Lỗi: ${err.message}`);
    }
  }
}



/**
 * Hàm UI / Entry Point: Chạy Batch nhiều kỳ liên tiếp cho Đơn giá tháng
 * Theo quy ước: actionEntityDescription (runBatchMonthlyPriceList)
 */
/**
 * Chạy Batch Đơn Giá Tháng hàng loạt từ kỳ Bắt đầu đến Kết thúc (Dùng Toast thông báo)
 */
function runBatchMonthlyPriceList() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Nhập kỳ BẮT ĐẦU (Giữ lại prompt tương tác duy nhất lúc khởi tạo)
  const startRes = ui.prompt(
    "Chạy Batch Đơn Giá Tháng",
    "Nhập kỳ BẮT ĐẦU (Định dạng YYYY-MM, VD: 2026-03):",
    ui.ButtonSet.OK_CANCEL
  );
  if (startRes.getSelectedButton() !== ui.Button.OK) return;
  const startPeriod = startRes.getResponseText().trim();

  // 2. Nhập kỳ KẾT THÚC (Giữ lại prompt tương tác duy nhất lúc khởi tạo)
  const endRes = ui.prompt(
    "Chạy Batch Đơn Giá Tháng",
    "Nhập kỳ KẾT THÚC (Định dạng YYYY-MM, VD: 2026-08):",
    ui.ButtonSet.OK_CANCEL
  );
  if (endRes.getSelectedButton() !== ui.Button.OK) return;
  const endPeriod = endRes.getResponseText().trim();

  if (!startPeriod || !endPeriod || startPeriod > endPeriod) {
    ui.alert("Lỗi", "Kỳ nhập không hợp lệ hoặc kỳ bắt đầu lớn hơn kỳ kết thúc.", ui.ButtonSet.OK);
    return;
  }

  let currentPeriod = startPeriod;
  let successCount = 0;
  let errorLog = [];

  // Hàm helper hiển thị toast ở góc phải dưới màn hình Google Sheets
  const progressToast = (msg, title = "Batch Đơn Giá Tháng") => {
    ss.toast(msg, title, 6);
  };

  progressToast(`🚀 Bắt đầu chạy batch từ ${startPeriod} đến ${endPeriod}...`, "Khởi động");

  try {
    // 3. Vòng lặp chạy ngầm xuyên suốt các kỳ không bị ngắt quãng bởi hộp thoại
    while (currentPeriod <= endPeriod) {
      progressToast(`⏳ Đang xử lý tính đơn giá cho kỳ: [ ${currentPeriod} ]...`);
      
      try {
        // Gọi hàm xử lý cốt lõi tính đơn giá bình quân của kỳ hiện tại
        AvgPriceEngine.monthly_price_list(currentPeriod);
        successCount++;
      } catch (err) {
        errorLog.push(`Kỳ ${currentPeriod}: ${err.message}`);
      }

      // Chuyển sang kỳ tiếp theo thông qua hàm phụ trợ getNextPeriod_ có sẵn trong dự án
      currentPeriod = getNextPeriod_(currentPeriod);
      if (!currentPeriod) break;
    }

    // 4. Tổng kết toàn bộ quá trình bằng 1 thông báo duy nhất khi kết thúc
    let summaryMsg = `Đã hoàn thành chạy batch từ ${startPeriod} đến ${endPeriod}.\n- Thành công: ${successCount} kỳ.`;
    if (errorLog.length > 0) {
      summaryMsg += `\n- Có lỗi xảy ra ở các kỳ:\n${errorLog.join("\n")}`;
      ui.alert("⚠️ Hoàn tất có cảnh báo", summaryMsg, ui.ButtonSet.OK);
    } else {
      progressToast(`✨ Đã chạy batch thành công toàn bộ ${successCount} kỳ!`, "Hoàn tất");
      ui.alert("✅ Thành công", summaryMsg, ui.ButtonSet.OK);
    }

  } catch (e) {
    ui.alert("❌ Lỗi Hệ thống", `Quá trình chạy batch bị gián đoạn: ${e.message}`, ui.ButtonSet.OK);
  }
}


/**
 * Kịch bản chạy hàng loạt (Batch) cho giá trị tồn kho nhiều tháng liên tục
 * Nhập kỳ bắt đầu và kỳ kết thúc, tự động kiểm tra định dạng, chạy tuần tự và báo cáo qua Toast/Alert cuối kỳ.
 */
function runBatchStockEngineValueRange() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Nhập kỳ bắt đầu
  const startResponse = ui.prompt(
    '🔄 Chạy Giá Trị Tồn Kho Hàng Loạt', 
    'Nhập KỲ BẮT ĐẦU (Định dạng YYYY-MM, VD: 2026-01):', 
    ui.ButtonSet.OK_CANCEL
  );
  if (startResponse.getSelectedButton() !== ui.Button.OK) return;
  const startPeriod = startResponse.getResponseText().trim();

  if (!/^\d{4}-\d{2}$/.test(startPeriod)) {
    ui.alert('⚠️ Lỗi Định Dạng', 'Kỳ bắt đầu không đúng định dạng YYYY-MM.', ui.ButtonSet.OK);
    return;
  }

  // 2. Nhập kỳ kết thúc
  const endResponse = ui.prompt(
    '🔄 Chạy Giá Trị Tồn Kho Hàng Loạt', 
    'Nhập KỲ KẾT THÚC (Định dạng YYYY-MM, VD: 2026-12):', 
    ui.ButtonSet.OK_CANCEL
  );
  if (endResponse.getSelectedButton() !== ui.Button.OK) return;
  const endPeriod = endResponse.getResponseText().trim();

  if (!/^\d{4}-\d{2}$/.test(endPeriod)) {
    ui.alert('⚠️ Lỗi Định Dạng', 'Kỳ kết thúc không đúng định dạng YYYY-MM.', ui.ButtonSet.OK);
    return;
  }

  // 3. Tạo danh sách các kỳ theo trình tự thời gian
  const periods = generatePeriodRange_(startPeriod, endPeriod);
  if (periods.length === 0) {
    ui.alert('⚠️ Lỗi Khoảng Thời Gian', 'Kỳ bắt đầu lớn hơn kỳ kết thúc hoặc khoảng thời gian không hợp lệ!', ui.ButtonSet.OK);
    return;
  }

  // Hàm helper hiển thị toast ở góc phải dưới màn hình Google Sheets
  const progressToast = (msg, title = "Batch Giá Trị Tồn Kho") => {
    ss.toast(msg, title, 6);
  };

  progressToast(`🚀 Bắt đầu chạy batch tồn kho từ ${startPeriod} đến ${endPeriod} (${periods.length} kỳ)...`, "Khởi động");
  SpreadsheetApp.flush();

  let successCount = 0;
  let errorLog = [];

  // 4. Thực thi tuần tự từng kỳ với cơ chế bẫy lỗi độc lập cho từng tháng
  try {
    for (let i = 0; i < periods.length; i++) {
      const targetPeriod = periods[i];
      progressToast(`⏳ Đang xử lý kỳ [ ${targetPeriod} ] (${i + 1}/${periods.length})...`);
      
      try {
        StockEngine.calculateValueSummary(targetPeriod, true);
        successCount++;
      } catch (err) {
        errorLog.push(`Kỳ ${targetPeriod}: ${err.message}`);
      }
    }

    SpreadsheetApp.flush();

    // 5. Báo cáo tổng kết cuối cùng
    let summaryMsg = `Đã hoàn thành chạy batch tồn kho từ ${startPeriod} đến ${endPeriod}.\n- Thành công: ${successCount}/${periods.length} kỳ.`;
    
    if (errorLog.length > 0) {
      summaryMsg += `\n\n- Phát sinh lỗi tại các kỳ:\n${errorLog.join("\n")}`;
      ui.alert('⚠️ Hoàn Tất Có Cảnh Báo', summaryMsg, ui.ButtonSet.OK);
    } else {
      progressToast(`✨ Đã chạy batch thành công toàn bộ ${successCount} kỳ tồn kho!`, "Hoàn tất");
      ui.alert('✅ Hoàn Tất Thành Công', summaryMsg, ui.ButtonSet.OK);
    }

  } catch (e) {
    ui.alert(
      '❌ Lỗi Hệ Thống', 
      `Quá trình chạy batch bị gián đoạn toàn cục: ${e.message}`, 
      ui.ButtonSet.OK
    );
  }
}

/**
 * Giao diện chạy độc lập tính giá trị tồn kho
 */
function runStockEngineValue() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Kiểm tra Đơn giá Tức thì', 'Nhập kỳ tính toán (Định dạng YYYY-MM):', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const period = response.getResponseText().trim();
    if (!period) return;

    // Quy tắc xác thực định dạng kỳ (YYYY-MM)
    const periodRegex = /^\d{4}-(?:0[1-9]|1[0-2])$/;
    if (!periodRegex.test(period)) {
      ui.alert('Lỗi: Định dạng kỳ không hợp lệ! Vui lòng nhập theo chuẩn YYYY-MM (Ví dụ: 2026-08).');
      return;
    }

    try {
      StockEngine.calculateValueSummary(period);
      ui.alert(`Đã cập nhật bảng INVENTORY_VALUE_SUMMARY cho kỳ ${period}.`);
    } catch (err) {
      ui.alert(`Lỗi: ${err.message}`);
    }
  }
}



/** Entry Point gọi từ Menu UI để khởi tạo định mức từ ITEM_MASTER */
function runInitRecipeBom() {
  recipeBomInitFromItemMaster();
}

/** Entry Point gọi từ Menu UI để điền tên tự động cho RECIPE_BOM */
function runPopulateRecipeBomNames() {
  recipeBomPopulateNames();
}

/** Entry Point gọi từ Menu UI để dọn dẹp dòng trùng lặp trong RECIPE_BOM */
function runCleanRecipeBomDuplicates() {
  recipeBomCleanDuplicates();
}




/** Entry Point gọi từ Menu UI để tính toán lại chi phí món ăn theo kỳ */
function runCalculateMenuCost() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "📊 Tính Chi Phí Món Theo Kỳ", 
    "Vui lòng nhập kỳ tra cứu đơn giá (VD: 2026-03):", 
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const period = response.getResponseText().trim();
    if (period) {
      menuCalculateCost(period);
    } else {
      ui.alert("⚠️ Cảnh báo", "Kỳ tra cứu không được để trống!", ui.ButtonSet.OK);
    }
  }
}




/** Entry Point gọi từ Menu UI để chạy chốt snapshot cost món theo kỳ */
function runSaveMenuCostSnapshot() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "📸 Chốt Snapshot Cost Món Theo Kỳ", 
    "Vui lòng nhập kỳ báo cáo (VD: 2026-03):", 
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const period = response.getResponseText().trim();
    if (period) {
      menuSaveCostSnapshot(period);
    } else {
      ui.alert("⚠️ Cảnh báo", "Kỳ báo cáo không được để trống!", ui.ButtonSet.OK);
    }
  }
}



/**
 * Entry Point gọi từ Menu UI để đồng bộ và tổng hợp chi phí (EXPENSE) từ Transaction theo kỳ.
 */
function runExpenseSyncFromTransaction() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "💰 Đồng Bộ Chi Phí (EXPENSE)", 
    "Vui lòng nhập kỳ hạch toán cần tổng hợp (Định dạng YYYY-MM, VD: 2026-08):", 
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const period = response.getResponseText().trim();
    
    // Validate định dạng YYYY-MM
    if (!/^\d{4}-\d{2}$/.test(period)) {
      ui.alert(
        "⚠️ Lỗi Định Dạng", 
        "Kỳ hạch toán không hợp lệ. Vui lòng nhập đúng định dạng YYYY-MM (VD: 2026-08).", 
        ui.ButtonSet.OK
      );
      return;
    }

    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Đang tổng hợp chi phí EXPENSE cho kỳ ${period}...`, 'Hệ thống');
      SpreadsheetApp.flush();

      // Gọi hàm Core đã viết trong module EXPENSE
      expenseSyncFromTransaction(period);

    } catch (error) {
      ui.alert(
        "❌ Lỗi Thực Thi", 
        `Đã xảy ra lỗi trong quá trình đồng bộ EXPENSE:\n\n${error.message}`, 
        ui.ButtonSet.OK
      );
    }
  }
}



/**
 * Entry Point gọi từ Menu UI để chạy tính tiêu hao lý thuyết thông qua TheoreticalEngine
 */
function runCalculateTheoreticalUsage() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "📊 Tính Tiêu Hao Lý Thuyết", 
    "Vui lòng nhập kỳ báo cáo (Định dạng YYYY-MM, VD: 2026-08):", 
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const period = response.getResponseText().trim();
    if (!/^\d{4}-\d{2}$/.test(period)) {
      ui.alert("⚠️ Lỗi Định Dạng", "Kỳ không hợp lệ. Vui lòng nhập theo chuẩn YYYY-MM.", ui.ButtonSet.OK);
      return;
    }
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Đang tính tiêu hao lý thuyết kỳ ${period}...`, 'Hệ thống');
      TheoreticalEngine.calculateUsage(period);
    } catch (err) {
      ui.alert("❌ Lỗi thực thi", err.message, ui.ButtonSet.OK);
    }
  }
}
