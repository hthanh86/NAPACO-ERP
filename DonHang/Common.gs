/**
 * FILE: Common.gs
 * Bộ hàm tiện ích ánh xạ cột động đa từ khóa, lắp ráp dòng theo Schema & Sinh mã tự động
 */

/**
 * Chuẩn hóa khóa so sánh tiêu đề cột
 * - Xử lý triệt để chữ hoa/thường, cả 'Đ' và 'đ'
 * - Khử dấu tiếng Việt chuẩn NFD
 * - Ép loại bỏ toàn bộ khoảng trắng, ký tự đặc biệt, giữ lại [a-z0-9]
 */
function normalizeKey(str) {
  if (!str) return '';
  let s = String(str).trim().toLowerCase();
  s = s.replace(/đ/g, 'd').replace(/Đ/g, 'd');
  s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  return s.replace(/[^a-z0-9]/g, '');
}

/**
 * THÊM MỚI: Trích xuất mảng từ khóa cho Form (Sheet TAO DON) từ Schema
 */
function getFormHeaders(tableName, key) {
  if (!TABLE_MAPPING_SCHEMA[tableName] || !TABLE_MAPPING_SCHEMA[tableName].fields) return [];
  const field = TABLE_MAPPING_SCHEMA[tableName].fields.find(f => f.key === key);
  return (field && field.formHeaders) ? field.formHeaders : [];
}

/**
 * THÊM MỚI: Trích xuất mảng từ khóa cho Database Đích từ Schema
 */
function getTargetHeaders(tableName, key) {
  if (!TABLE_MAPPING_SCHEMA[tableName] || !TABLE_MAPPING_SCHEMA[tableName].fields) return [];
  const field = TABLE_MAPPING_SCHEMA[tableName].fields.find(f => f.key === key);
  return (field && field.targetHeaders) ? field.targetHeaders : [];
}

/**
 * HÀM QUÉT TIÊU ĐỀ ĐỘNG
 */
function getDynamicSheetInfo(sheet, preferredHeaderRow = 7, anchorText = '') {
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const scanRows = Math.min(Math.max(lastRow, 1), 15);
  const matrix = sheet.getRange(1, 1, scanRows, lastCol).getValues();
  let hRow = preferredHeaderRow - 1;
  if (anchorText) {
    const target = normalizeKey(anchorText);
    for (let r = 0; r < matrix.length; r++) {
      const rowNorm = matrix[r].map(c => normalizeKey(c));
      if (rowNorm.includes(target)) {
        hRow = r;
        break;
      }
    }
  }
  const headerRowData = matrix[hRow] || [];
  const colMap = {};
  for (let c = 0; c < headerRowData.length; c++) {
    const key = normalizeKey(headerRowData[c]);
    if (key) colMap[key] = c + 1; // 1-based index
  }
  return {
    headerRow: hRow + 1,
    startDataRow: hRow + 2,
    lastRow: lastRow,
    lastCol: lastCol,
    colMap: colMap
  };
}

function getColIndex(sheetInfo, colNames) {
  if (!Array.isArray(colNames)) colNames = [colNames];
  for (let name of colNames) {
    const key = normalizeKey(name);
    if (sheetInfo.colMap[key]) {
      return sheetInfo.colMap[key] - 1; // 0-based index
    }
  }
  return -1;
}

function generateAutoPO(customerName, dateObj) {
  if (!customerName) return '';
  const yy = String(dateObj.getFullYear()).slice(2);
  const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
  const dd = String(dateObj.getDate()).padStart(2, '0');
  const dateStr = yy + mm + dd;
  let cleanName = String(customerName).normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  cleanName = cleanName.replace(/đ/g, 'd').replace(/Đ/g, 'D');
  cleanName = cleanName.toUpperCase().replace(/[^A-Z0-9]/g, '-').replace(/-+/g, '-').replace(/^-|-$/g, '');
  return `PO-${dateStr}-${cleanName}`;
}

/**
 * LẮP RÁP DÒNG THEO SCHEMA CONFIG (Single Source of Truth)
 */
function buildRowBySchema(schemaConfig, sheetInfo, dataObj) {
  const row = new Array(sheetInfo.lastCol).fill('');
  for (let f = 0; f < schemaConfig.fields.length; f++) {
    const field = schemaConfig.fields[f];
    const colIdx = getColIndex(sheetInfo, field.targetHeaders);
    if (colIdx >= 0) {
      let val = dataObj[field.key];
      if (val === undefined || val === null || val === '') {
        val = field.defaultValue !== undefined ? field.defaultValue : '';
      }
      row[colIdx] = val;
    }
  }
  return row;
}

/**
 * SINH MÃ SẢN PHẨM / KHÁCH HÀNG TỰ TĂNG: PRExxxx (SP0001, KH0001...)
 */
function getNextFixedSeqId(sheet, prefix, colIdx) {
  const lastR = sheet.getLastRow();
  if (lastR < 2 || colIdx < 0) return `${prefix}0001`;

  const vals = sheet.getRange(2, colIdx + 1, lastR - 1, 1).getValues();
  const regex = new RegExp(`^${prefix}(\\d+)$`, 'i');
  let maxSeq = 0;

  for (let i = 0; i < vals.length; i++) {
    const val = String(vals[i][0] || '').trim();
    const match = val.match(regex);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxSeq) maxSeq = num;
    }
  }

  return `${prefix}${String(maxSeq + 1).padStart(4, '0')}`;
}

function getNextFixedSeqId(sheet, prefix, colIdx) {
  const lastR = sheet.getLastRow();
  if (lastR < 2 || colIdx < 0) return 0;

  const vals = sheet.getRange(2, colIdx + 1, lastR - 1, 1).getValues();
  const regex = new RegExp(`^${prefix}(\\d+)$`, 'i');
  let maxSeq = 0;

  for (let i = 0; i < vals.length; i++) {
    const val = String(vals[i][0] || '').trim();
    const match = val.match(regex);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxSeq) maxSeq = num;
    }
  }
  return maxSeq;
}