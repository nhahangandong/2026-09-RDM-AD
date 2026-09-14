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
      .addItem('▶️ Chạy tính Tồn kho Số lượng (Kiểm thử)', 'runStockEngineQtyUI')
      .addItem('▶️ Chạy tính Tồn kho Số lượng (Kiểm thử)', 'runBatchCalculateQtySummary')
      .addSeparator()
      .addItem('🚀 Chạy Batch Tổng hợp Kỳ (Qty + Price)', 'uiRunBatchPipeline')
      .addSeparator()
      .addItem('🔍 Kiểm tra Đơn giá Tức thì (Monthly Avg Price)', 'uiRunMonthlyAvgPrice')
      .addItem('📝 Chốt Giá Tháng Chính Thức (Monthly Price List)', 'uiRunMonthlyPriceList')
      .addItem('Chạy Batch Đơn Giá Tháng', 'runBatchMonthlyPriceList')
    )
    .addSeparator()

    // 5. SUBMENU KHỞI TẠO & HỆ THỐNG
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


function runStockEngineQtyUI() {
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
function runBatchMonthlyPriceList() {
  const ui = SpreadsheetApp.getUi();
  
  const startRes = ui.prompt(
    "Chạy Batch Đơn Giá Tháng",
    "Nhập kỳ BẮT ĐẦU (Định dạng YYYY-MM, VD: 2026-03):",
    ui.ButtonSet.OK_CANCEL
  );
  if (startRes.getSelectedButton() !== ui.Button.OK) return;
  const startPeriod = startRes.getResponseText().trim();

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

  const progressToast = (msg) => SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Batch Đơn Giá", 5);

  try {
    while (currentPeriod <= endPeriod) {
      progressToast(`Đang xử lý kỳ: ${currentPeriod}...`);
      
      try {
        // Gọi hàm xử lý cốt lõi (ví dụ từ AvgPriceEngine hoặc service tương ứng)
        AvgPriceEngine.monthly_price_list(currentPeriod);
        successCount++;
      } catch (err) {
        errorLog.push(`Kỳ ${currentPeriod}: ${err.message}`);
      }

      currentPeriod = getNextPeriod_(currentPeriod);
      if (!currentPeriod) break;
    }

    let summaryMsg = `Đã hoàn thành chạy batch từ ${startPeriod} đến ${endPeriod}.\n- Thành công: ${successCount} kỳ.`;
    if (errorLog.length > 0) {
      summaryMsg += `\n- Có lỗi xảy ra:\n${errorLog.join("\n")}`;
    }
    ui.alert("Kết quả chạy Batch", summaryMsg, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert("Lỗi Hệ thống", `Quá trình chạy batch bị gián đoạn: ${e.message}`, ui.ButtonSet.OK);
  }
}
