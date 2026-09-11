function setupSheetsFromSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  const schemaSheet = ss.getSheetByName("SCHEMA");
  const ui = SpreadsheetApp.getUi();

  if (!schemaSheet) {
    ui.alert("⚠️ Thông báo", "Không tìm thấy sheet 'SCHEMA'!", ui.ButtonSet.OK);
    return;
  }

  const schemaData = schemaSheet.getDataRange().getValues();
  if (schemaData.length <= 1) {
    ui.alert("⚠️ Thông báo", "Sheet 'SCHEMA' đang trống!", ui.ButtonSet.OK);
    return;
  }

  // 1. Gom nhóm danh sách cột theo SCHEMA_NAME từ bảng SCHEMA
  const mapSchema = {};
  for (let i = 1; i < schemaData.length; i++) {
    const row = schemaData[i];
    const schemaName = row[0] ? row[0].toString().trim() : "";
    const sheetName  = row[1] ? row[1].toString().trim() : "";
    const colIndex   = parseInt(row[2], 10);
    const colHeader  = row[4] ? row[4].toString().trim() : "";

    // Bỏ qua dòng trống, dòng thuộc về chính SCHEMA hoặc dữ liệu cột thiếu
    if (schemaName && schemaName.toUpperCase() !== "SCHEMA" && sheetName && !isNaN(colIndex) && colHeader) {
      if (!mapSchema[schemaName]) {
        mapSchema[schemaName] = {
          sheetName: sheetName,
          columns: []
        };
      }
      mapSchema[schemaName].columns.push({
        index: colIndex,
        header: colHeader
      });
    }
  }

  // 2. Lấy danh sách các Table ĐÃ TỒN TẠI để KHÔNG XÓA NHỮNG BẢNG ĐANG CÓ DỮ LIỆU
  const existingTableSheetIds = new Set();
  try {
    const spreadsheetInfo = Sheets.Spreadsheets.get(spreadsheetId, { includeGridData: false });
    spreadsheetInfo.sheets.forEach(s => {
      if (s.tables && s.tables.length > 0) {
        existingTableSheetIds.add(s.properties.sheetId);
      }
    });
  } catch (err) {
    Logger.log("Chưa thể lấy danh sách Tables: " + err.message);
  }

  const addTableRequests = [];
  let createdSheetsCount = 0;
  let processedTablesCount = 0;

  // 3. Lặp qua các SCHEMA_NAME đã gom nhóm
  Object.keys(mapSchema).forEach(schemaName => {
    const schemaObj = mapSchema[schemaName];
    const sheetName = schemaObj.sheetName;

    let targetSheet = ss.getSheetByName(sheetName);

    // Nếu chưa có Sheet vật lý thì tạo mới theo sheet_name
    if (!targetSheet) {
      targetSheet = ss.insertSheet(sheetName);
      createdSheetsCount++;
    }

    const sheetId = targetSheet.getSheetId();

    // NẾU SHEET NÀY ĐÃ CÓ TABLE RỒI -> BỎ QUA KHÔNG CAN THIỆP ĐỂ BẢO VỆ DỮ LIỆU
    if (existingTableSheetIds.has(sheetId)) {
      return;
    }

    // Sắp xếp các cột theo col_index
    const columns = schemaObj.columns.sort((a, b) => a.index - b.index);
    const headers = columns.map(col => col.header);

    let lastRow = targetSheet.getLastRow();

    // Nếu Sheet hoàn toàn trống (mới tạo) -> Ghi Header vào dòng 1
    if (lastRow === 0) {
      targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      lastRow = 1;
    }

    // Phạm vi dòng của Table: Từ dòng 1 tới dòng dữ liệu cuối cùng (tối thiểu là dòng 2)
    const endRowIndex = Math.max(lastRow, 2);

    // ĐẶT TÊN BẢNG THEO SCHEMA_NAME THAY VÌ SHEET_NAME VẬT LÝ
    const tableName = `tbl_${schemaName.toLowerCase().trim().replace(/[^a-z0-9_]/g, "_")}`;

    // Tạo Request addTable mà KHÔNG XÓA DỮ LIỆU CŨ
    addTableRequests.push({
      addTable: {
        table: {
          name: tableName,
          range: {
            sheetId: sheetId,
            startRowIndex: 0,
            endRowIndex: endRowIndex,
            startColumnIndex: 0,
            endColumnIndex: headers.length
          }
        }
      }
    });

    processedTablesCount++;
  });

  // 4. Thực thi Batch Update tạo Table mới cho các Sheet chưa có Table
  if (addTableRequests.length > 0) {
    try {
      Sheets.Spreadsheets.batchUpdate({ requests: addTableRequests }, spreadsheetId);
    } catch (e) {
      Logger.log("Lỗi thực thi API Table: " + e.message);
      ui.alert("⚠️ Lỗi API Table", e.message, ui.ButtonSet.OK);
      return;
    }
  }

  ui.alert(
    "✅ Thành công",
    `Đã hoàn tất đồng bộ Table từ SCHEMA theo schema_name:\n- Số sheet mới tạo: ${createdSheetsCount}\n- Số Bảng (Table) mới khởi tạo: ${processedTablesCount}`,
    ui.ButtonSet.OK
  );
}
