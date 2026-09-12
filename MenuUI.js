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
    .addSeparator()

    // 2. SUBMENU MAPPING (Chuẩn hóa tên thô)
    .addSubMenu(ui.createMenu('🍁 1. Chuẩn hóa Mapping (Raw Data)')
      .addItem('1.1. Quét dữ liệu thô vào MAPPING', 'runSyncMappingFromTransactions')
      .addItem('1.2. Gợi ý item_name tự động (AUTO_MAP)', 'runAutoSuggestMappingNames')
      .addItem('1.3. Sinh mã SKU tự động cho MAPPING', 'runAutoGenerateSKUForMapping')
    )

    // 3. SUBMENU MASTER DATA (Đã bổ sung Engine Phân loại SKU)
    .addSubMenu(ui.createMenu('🏷️ 2. Danh mục & Master Data')
      .addItem('1. Đồng bộ ITEM_MASTER từ Giao dịch', 'runSyncItemMasterFromInventoryAndStocktake')
      .addItem('2. Đồng bộ Sales -> Menu', 'runSyncSalesToMenu')
      .addItem('3. Đồng bộ Menu -> Item Master', 'runSyncItemMasterFromMenu')
      .addSeparator()
      .addItem('4. Tự động phân loại 1k SKU (CLASSIFICATION)', 'runAutoSuggestItemMasterClassification')
    )
    .addSeparator()

    // 4. SUBMENU KHỞI TẠO & HỆ THỐNG
    .addSubMenu(ui.createMenu('⚙️ 3. Thiết lập hệ thống')
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
 * Entry Point: Gợi ý/Tự động phân loại danh mục cho Item Master.
 */
function runAutoSuggestItemMasterClassification() {
  itemMasterAutoSuggestClassification();
}

/**
 * Entry Point: Tra cứu và điền mã SKU (item_code) từ sheet MAPPING quay trở lại các sheet giao dịch.
 */
function runAssignItemCodeToTransactions() {
  transactionAssignItemCodeToAllSheets();
}
