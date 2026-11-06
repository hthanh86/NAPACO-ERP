/**
 * NAPACO ENTERPRISE OMS - DATA REPOSITORY LAYER
 * File: Repository.gs
 * Target: Google Apps Script Data Access Tier
 */

var Repository = {

  /**
   * Đọc toàn bộ dữ liệu bảng tính thành danh sách Object dựa trên Tên Cột Định Nghĩa
   */
  getTableData: function(spreadsheetId, sheetName, columnConfigObj) {
    AppLogger.info('REPOSITORY', 'Bắt đầu đọc dữ liệu bảng', { sheetName: sheetName });
    
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      AppLogger.error('REPOSITORY', 'Không tìm thấy Sheet: ' + sheetName);
      throw new Error('Sheet ' + sheetName + ' không tồn tại trong hệ thống!');
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    if (values.length <= 1) {
      AppLogger.warn('REPOSITORY', 'Bảng rỗng hoặc chỉ có dòng tiêu đề', { sheetName: sheetName });
      return [];
    }

    var headers = values[0];
    var colIndexMap = {};

    // Ánh xạ vị trí cột thực tế trên Sheet với Key trong COLUMNS_CONFIG
    for (var key in columnConfigObj) {
      var targetColName = columnConfigObj[key];
      var idx = headers.indexOf(targetColName);
      if (idx !== -1) {
        colIndexMap[key] = idx;
      } else {
        AppLogger.warn('REPOSITORY', 'Cột cấu hình không tìm thấy trên Sheet: ' + targetColName, { sheet: sheetName, key: key });
      }
    }

    var resultList = [];
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      var record = {};
      var hasData = false;

      for (var fieldKey in colIndexMap) {
        var cIdx = colIndexMap[fieldKey];
        var cellVal = row[cIdx];
        record[fieldKey] = cellVal;
        if (cellVal !== '' && cellVal !== null && cellVal !== undefined) {
          hasData = true;
        }
      }

      if (hasData) {
        resultList.push(record);
      }
    }

    AppLogger.info('REPOSITORY', 'Đã tải thành công ' + resultList.length + ' bản ghi', { sheetName: sheetName });
    AppLogger.table('REPOSITORY', sheetName, resultList);
    
    return resultList;
  },

  /**
   * Ghi theo lô (Batch Write) đảm bảo đúng thứ tự cột của Bảng
   */
  batchInsert: function(spreadsheetId, sheetName, columnConfigObj, recordsArray) {
    if (!recordsArray || recordsArray.length === 0) return 0;
    
    AppLogger.info('REPOSITORY', 'Chuẩn bị Batch Insert vào ' + sheetName, { count: recordsArray.length });
    
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName(sheetName);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Tạo ma trận ghi dữ liệu chuẩn theo thứ tự Header vật lý trên Sheet
    var matrixToWrite = [];
    for (var i = 0; i < recordsArray.length; i++) {
      var item = recordsArray[i];
      var rowData = new Array(headers.length).fill('');

      for (var fieldKey in columnConfigObj) {
        var colName = columnConfigObj[fieldKey];
        var colIdx = headers.indexOf(colName);
        if (colIdx !== -1 && item[fieldKey] !== undefined) {
          rowData[colIdx] = item[fieldKey];
        }
      }
      matrixToWrite.push(rowData);
    }

    // Kiểm tra dữ liệu qua Logger trước khi ghi vật lý
    AppLogger.table('REPOSITORY_WRITE_PREVIEW', sheetName, matrixToWrite);

    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, matrixToWrite.length, headers.length).setValues(matrixToWrite);
    SpreadsheetApp.flush();

    AppLogger.info('REPOSITORY', 'Batch Insert thành công vào ' + sheetName, { rowCount: matrixToWrite.length });
    return matrixToWrite.length;
  }
};