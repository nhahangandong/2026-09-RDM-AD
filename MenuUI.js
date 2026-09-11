/**
 * Simple Trigger tự động tạo Menu khi mở file Google Sheets.
 * Thiết kế phân cấp Submenu theo chuẩn quản trị.
 */
    
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  // 1. Tạo Menu Cấp 1 trên thanh công cụ và nhúng Submenu vào
  ui.createMenu('🛠️ QUẢN TRỊ HỆ THỐNG')  
    .addSubMenu(ui.createMenu('⚙️ Hệ thống')
      .addItem('🚀 Khởi tạo / Cập nhật Sheet từ SCHEMA', 'setupSheetsFromSchema')
      .addSeparator()
      .addItem('🔄 1. Quét dữ liệu thô vào MAPPING', 'runSyncMappingFromTransactions')
      .addItem('⚡ 2. Sinh mã SKU tự động cho MAPPING', 'runAutoGenerateSKUForMapping')
      .addItem('📌 3. Gán mã SKU vào các sheet giao dịch', 'runAssignItemCodeToTransactions')
    )

    .addSubMenu(ui.createMenu('📦 Danh mục & Master Data')
      .addItem('1. Đồng bộ ITEM_MASTER từ Giao dịch', 'runSyncItemMasterFromInventoryAndStocktake')
      .addItem("2. Đồng bộ Sales -> Menu", "runSyncSalesToMenu")
      .addItem("3. Đồng bộ Menu -> Item Master", "runSyncItemMasterFromMenu")
    )

    .addToUi();
}
