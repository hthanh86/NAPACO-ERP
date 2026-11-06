/**
 * NAPACO ENTERPRISE OMS - LOGGING UTILITY
 * File: LoggerUtils.gs
 * Target: Google Apps Script Utils Layer
 */

var AppLogger = {
  
  /**
   * Kiểm tra trạng thái Debug
   * @private
   */
  isDebugEnabled_: function() {
    return typeof CONFIG !== 'undefined' && Boolean(CONFIG.DEBUG_MODE);
  },

  /**
   * Log thông tin tiến trình nghiệp vụ
   */
  info: function(tag, message, data) {
    if (!this.isDebugEnabled_()) return;
    var logMsg = '[INFO] [' + tag + '] ' + message;
    if (data !== undefined) {
      logMsg += ' | DATA: ' + (typeof data === 'object' ? JSON.stringify(data) : data);
    }
    console.log(logMsg);
  },

  /**
   * Log cảnh báo
   */
  warn: function(tag, message, data) {
    if (!this.isDebugEnabled_()) return;
    var logMsg = '[WARN] [' + tag + '] ' + message;
    if (data !== undefined) {
      logMsg += ' | DATA: ' + (typeof data === 'object' ? JSON.stringify(data) : data);
    }
    console.warn(logMsg);
  },

  /**
   * Log lỗi nghiệp vụ hoặc ngoại lệ hệ thống
   */
  error: function(tag, message, err) {
    // Lỗi nghiêm trọng luôn được ghi nhận vào Console của GCP bất kể cờ DEBUG
    var errMsg = '[ERROR] [' + tag + '] ' + message;
    if (err) {
      errMsg += ' | DETAILS: ' + (err.stack ? err.stack : err.toString());
    }
    console.error(errMsg);
  },

  /**
   * Log dạng bảng đối chiếu cấu trúc dữ liệu trước khi ghi
   */
  table: function(tag, label, dataArray) {
    if (!this.isDebugEnabled_()) return;
    console.log('====== [DATA CHECK TABLE: ' + tag + ' - ' + label + '] ======');
    if (Array.isArray(dataArray) && dataArray.length > 0) {
      console.log('Tổng số bản ghi: ' + dataArray.length);
      console.log(JSON.stringify(dataArray.slice(0, 5), null, 2)); // In tối đa 5 phần tử mẫu
    } else {
      console.log('Mảng dữ liệu rỗng!');
    }
    console.log('===========================================================');
  }
};