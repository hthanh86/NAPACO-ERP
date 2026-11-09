const FilmColumns_ = {
  po: ['PO_NVL_ID', 'Ngay_Dat', 'Ma_NCC', 'Nguoi_Dat', 'Ghi_Chu', 'Trang_Thai'],
  line: ['DN_ID', 'Loai_Nhua', 'VL_ID', 'So_Luong', 'Don_Gia', 'Ghi_Chu', 'Trang_Thai'],
  lineNew: ['PO_NVL_ID', 'Ngay_Giao_Du_Kien'],
  material: ['VL_ID', 'Cấp Độ', 'Màu Sắc', 'Tĩnh Điện', 'Độ Dày', 'Khổ Nhựa'],
  supplier: ['Ma_NCC', 'Ten_NCC', 'Loai_Cung_Cap'],
  receipt: ['N_NVL_ID', 'DN_ID', 'So_KG_Nhan']
};

function getFilmOrders() {
  const db = filmDatabase_(true);
  const orders = filmOrders_(db);
  for (const order of orders) {
    for (const line of order.lines) {
      line.hasAssignmentColumn = db.details.positions.has(CONFIG.COLUMNS.ORDER_DETAILS.FILM_LINE_ID);
      line.canAssign = line.hasAssignmentColumn &&
        order.Trang_Thai !== 'Hủy' && line.Trang_Thai !== 'Hủy';
    }
  }
  return {
    orders,
    suppliers: db.suppliers,
    materials: db.materials.rows.map(row => row.data).filter(row => row.VL_ID),
    today: filmDate_(new Date()),
    setup: db.setup,
    receiptWarning: db.receiptWarning
  };
}

function filmDescribeDetails_(db, rows) {
  if (!rows.length) return [];
  const names = CONFIG.SHEET_NAMES;
  const { ORDERS: order, CUSTOMERS: customer, PRODUCTS: product } = CONFIG.COLUMNS;
  const index = (book, name, key, value) => new Map(
    filmTable_(book, [name], [key, value]).rows
      .filter(({ data }) => data[key]).map(({ data }) => [data[key], data[value]]));
  const customers = index(db.master, names.CUSTOMERS, customer.ID, customer.NAME);
  const orders = index(db.operation, names.ORDERS, order.ID, order.CUSTOMER_ID);
  const products = index(db.master, names.PRODUCTS, product.ID, product.NAME);
  return rows.map(({ data }) => ({
    id: data.ID_Chi_Tiet,
    customer: customers.get(orders.get(data.ID_DonHang)) || '',
    product: products.get(data.Ma_SP) || '',
    qty: data.So_Luong_Dat || '',
    delivery: data.Ngay_Giao || '',
    assigned: !!data.DN_ID,
    editable: data.Trang_Thai === CONFIG.LINE_STATUS.NEW
  }));
}

function filmDetails_(operation, headersOnly = false, include) {
  const detail = CONFIG.COLUMNS.ORDER_DETAILS;
  return filmTable_(operation, [CONFIG.SHEET_NAMES.ORDER_DETAILS],
    [detail.ID, detail.STATUS], [detail.FILM_LINE_ID, detail.ORDER_ID, detail.PRODUCT_ID,
      detail.QTY_ORDER, detail.DELIVERY_DATE], include, headersOnly);
}

function filmAssignmentRows_(operation, lineId, includeAvailable) {
  const details = filmDetails_(operation, false, (cells, positions) => {
    const detail = CONFIG.COLUMNS.ORDER_DETAILS;
    const link = filmText_(cells[positions.get(detail.FILM_LINE_ID)]);
    return !!filmText_(cells[positions.get(detail.ID)]) &&
      ((lineId && link === lineId) ||
        (includeAvailable && !link && filmText_(cells[positions.get(detail.STATUS)]) === CONFIG.LINE_STATUS.NEW));
  });
  if (!details.positions.has(CONFIG.COLUMNS.ORDER_DETAILS.FILM_LINE_ID)) {
    throw new Error(CONFIG.SHEET_NAMES.ORDER_DETAILS + ' thiếu cột DN_ID.');
  }
  return details;
}

function filmAssertAssignmentLine_(operation, lineId) {
  const names = CONFIG.SHEET_NAMES;
  const lines = filmTable_(operation, [names.FILM_LINES], ['DN_ID', 'Trang_Thai'], ['PO_NVL_ID']);
  const matches = lines.rows.filter(({ data }) => data.DN_ID === lineId);
  if (matches.length !== 1 || matches[0].data.Trang_Thai === 'Hủy') {
    throw new Error('Dòng đặt màng không tồn tại hoặc không thể sửa.');
  }
  const poId = matches[0].data.PO_NVL_ID;
  if (poId && operation.getSheetByName(names.FILM_PO)) {
    const headers = filmTable_(operation, [names.FILM_PO], ['PO_NVL_ID', 'Trang_Thai']);
    if (headers.rows.some(({ data }) => data.PO_NVL_ID === poId && data.Trang_Thai === 'Hủy')) {
      throw new Error('Đơn đặt màng đã hủy.');
    }
  }
}

function getFilmOrderFormData(lineIds) {
  const db = filmAccess_();
  const material = filmMaterials_(db.master, ['DM_MAU_SAC']);
  const colors = [...new Set(material.rows.map(({ data }) => data.DM_MAU_SAC).filter(Boolean))];
  const today = filmDate_(new Date());
  if (lineIds === null) {
    return { today, colors, choices: filmDescribeDetails_(db, filmAssignmentRows_(db.operation, '', true).rows) };
  }
  return { today, colors, assignments: filmOrderAssignments_(db, lineIds) };
}

function getFilmOrderAssignments(lineIds) {
  return filmOrderAssignments_(filmAccess_(), lineIds);
}

function filmOrderAssignments_(db, lineIds) {
  if (!Array.isArray(lineIds)) throw new Error('Danh sách quy cách không hợp lệ.');
  const ids = new Set(lineIds.map(filmText_).filter(Boolean));
  const detail = CONFIG.COLUMNS.ORDER_DETAILS;
  const { rows } = filmDetails_(db.operation, false, (cells, positions) =>
    !!filmText_(cells[positions.get(detail.ID)]) &&
    ids.has(filmText_(cells[positions.get(detail.FILM_LINE_ID)])));
  const entries = filmDescribeDetails_(db, rows);
  const byLine = new Map();
  rows.forEach(({ data }, index) => {
    if (!byLine.has(data.DN_ID)) byLine.set(data.DN_ID, []);
    byLine.get(data.DN_ID).push(entries[index]);
  });
  return Array.from(byLine);
}

function getFilmAssignedDetails(lineId) {
  const db = filmAccess_();
  const id = filmText_(lineId);
  if (!id) throw new Error('Thiếu mã dòng đặt màng.');
  return filmDescribeDetails_(db, filmAssignmentRows_(db.operation, id, false).rows);
}

function getFilmAssignmentOptions(lineId) {
  const db = filmAccess_();
  const id = filmText_(lineId);
  if (id) filmAssertAssignmentLine_(db.operation, id);
  return filmDescribeDetails_(db, filmAssignmentRows_(db.operation, id, true).rows);
}

function saveFilmAssignments(lineId, originalIds, selectedIds) {
  const id = filmText_(lineId);
  if (!id || !Array.isArray(originalIds) || !Array.isArray(selectedIds)) {
    throw new Error('Dữ liệu gán CTDH không hợp lệ.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const { operation } = filmAccess_();
    filmAssertAssignmentLine_(operation, id);
    const details = filmDetails_(operation);
    if (!details.positions.has(CONFIG.COLUMNS.ORDER_DETAILS.FILM_LINE_ID)) {
      throw new Error(CONFIG.SHEET_NAMES.ORDER_DETAILS + ' thiếu cột DN_ID.');
    }
    const original = new Set(originalIds);
    const selected = new Set(selectedIds);
    const linked = details.rows.filter(({ data }) => data.DN_ID === id);
    if (linked.length !== original.size || linked.some(({ data }) => !original.has(data.ID_Chi_Tiet))) {
      throw new Error('CTDH đã thay đổi. Hãy mở lại Sửa/Bổ sung.');
    }
    const byId = new Map();
    for (const row of details.rows) {
      const key = row.data.ID_Chi_Tiet;
      if (original.has(key) || selected.has(key)) byId.set(key, byId.has(key) ? null : row);
    }
    const requests = [];
    for (const key of new Set([...original, ...selected])) {
      if (original.has(key) === selected.has(key)) continue;
      const row = byId.get(key);
      if (!row || !key || row.data.Trang_Thai !== CONFIG.LINE_STATUS.NEW ||
          (row.data.DN_ID && row.data.DN_ID !== id)) {
        throw new Error('Không thể đổi gán CTDH ' + key + '. Hãy mở lại Sửa/Bổ sung.');
      }
      filmWrite_(requests, details, row.index, { DN_ID: selected.has(key) ? id : '' });
    }
    if (requests.length) Sheets.Spreadsheets.batchUpdate({ requests }, operation.getId());
  } finally {
    lock.releaseLock();
  }
}

function filmText_(value) {
  return String(value == null ? '' : value).trim();
}

function filmDate_(value) {
  return value instanceof Date
    ? Utilities.formatDate(value, CONFIG.TIMEZONE, 'yyyy-MM-dd')
    : filmText_(value);
}

function filmInputDate_(value, required) {
  const day = filmText_(value);
  if (!day && !required) return '';
  const parsed = new Date(day + 'T00:00:00Z');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(day) || !Number.isFinite(+parsed) ||
      parsed.toISOString().slice(0, 10) !== day) {
    throw new Error('Ngày không hợp lệ.');
  }
  return day;
}

function filmTable_(book, names, columns, optional = [], include = () => true, headersOnly = false) {
  const sheet = names.map(name => book.getSheetByName(name)).find(Boolean);
  if (!sheet) throw new Error('Thiếu tab ' + names[0] + '.');
  const values = headersOnly && sheet.getLastRow()
    ? sheet.getRange(1, 1, Math.min(6, sheet.getLastRow()), sheet.getLastColumn()).getValues()
    : headersOnly ? [] : sheet.getDataRange().getValues();
  const index = values.slice(0, 6).findIndex(row => row.map(filmText_).includes(columns[0]));
  if (index < 0) throw new Error(sheet.getName() + ' thiếu cột ' + columns[0] + '.');
  const headers = values[index].map(filmText_);
  const positions = new Map();
  for (const name of columns.concat(optional)) {
    const position = headers.indexOf(name);
    if (position !== headers.lastIndexOf(name)) {
      throw new Error(sheet.getName() + ' có cột ' + name + ' bị trùng.');
    }
    if (position >= 0) positions.set(name, position);
    else if (columns.includes(name)) throw new Error(sheet.getName() + ' thiếu cột ' + name + '.');
  }
  const rows = [];
  for (let rowIndex = index + 1; !headersOnly && rowIndex < values.length; rowIndex++) {
    const cells = values[rowIndex];
    if (!include(cells, positions)) continue;
    const data = {};
    for (const [name, col] of positions) data[name] = filmDate_(cells[col]);
    rows.push({ index: rowIndex, data });
  }
  return { sheet, positions, rows };
}

function filmAccess_() {
  const env = getSecureEnv_();
  const master = SpreadsheetApp.openById(env.ID_MASTER);
  const operation = SpreadsheetApp.openById(env.ID_OPERATION);
  const names = CONFIG.SHEET_NAMES;
  const employee = CONFIG.COLUMNS.EMPLOYEES;
  const email = filmText_(Session.getActiveUser().getEmail()).toLowerCase();
  if (!email) throw new Error('Đăng nhập bằng tài khoản nhân viên để mở Đặt Màng.');
  const user = filmTable_(master, [names.EMPLOYEES],
    [employee.ID, employee.EMAIL, employee.STATUS, employee.DEPARTMENT, employee.ACCESS])
    .rows.map(row => row.data).find(row => row[employee.ID] &&
      row[employee.EMAIL].toLowerCase() === email &&
      row[employee.STATUS].toLowerCase() === CONFIG.EMPLOYEE_STATUS.ACTIVE.toLowerCase());
  if (!user || (!user[employee.DEPARTMENT].toLowerCase().includes(CONFIG.DEPARTMENTS.SALES.toLowerCase()) &&
      !/\b(Admin|Manager)\b/i.test(user[employee.ACCESS]))) {
    throw new Error('Đặt Màng yêu cầu quyền Kinh doanh, Admin hoặc Manager.');
  }

  return { master, operation, userId: user[employee.ID] };
}

function filmMaterials_(master, extraColumns = []) {
  return filmTable_(master, [CONFIG.SHEET_NAMES.FILM_MATERIALS, '01_04_ThongSoVatLieu'],
    FilmColumns_.material.concat(extraColumns));
}

function filmDatabase_(listOnly = false) {
  const { master, operation, userId } = filmAccess_();
  const names = CONFIG.SHEET_NAMES;
  const material = filmMaterials_(master);
  const suppliers = filmTable_(master, [names.FILM_SUPPLIERS, '01_03_DM_NhaCungCap'], FilmColumns_.supplier)
    .rows.map(row => row.data).filter(row => /^NCC(0[1-9]|[1-9]\d)$/.test(row.Ma_NCC) &&
      /màng/i.test(row.Loai_Cung_Cap));
  const lines = filmTable_(operation, [names.FILM_LINES], FilmColumns_.line,
    FilmColumns_.lineNew.concat('Ngay_Dat', 'Ma_NCC'));
  const detail = CONFIG.COLUMNS.ORDER_DETAILS;
  const details = filmDetails_(operation, listOnly);
  const setup = [];
  let headers;
  try {
    headers = filmTable_(operation, [names.FILM_PO], FilmColumns_.po);
  } catch (err) {
    setup.push(err.message);
  }
  for (const col of FilmColumns_.lineNew) {
    if (!lines.positions.has(col)) setup.push(names.FILM_LINES + ' thiếu cột ' + col + '.');
  }
  if (!details.positions.has(detail.FILM_LINE_ID)) setup.push(names.ORDER_DETAILS + ' thiếu cột DN_ID.');

  const receipts = new Map();
  let receiptWarning = '';
  try {
    const incoming = filmTable_(operation, [names.FILM_RECEIPTS], FilmColumns_.receipt);
    const seen = new Set();
    for (const { data: row } of incoming.rows) {
      if (!row.DN_ID) continue;
      const kg = Number(row.So_KG_Nhan);
      if (!row.N_NVL_ID || seen.has(row.N_NVL_ID) ||
          !row.So_KG_Nhan || !Number.isFinite(kg) || kg < 0) {
        throw new Error('Kiểm tra mã phiếu và số kg của các dòng kho có DN_ID.');
      }
      seen.add(row.N_NVL_ID);
      const linked = receipts.get(row.DN_ID) || { kg: 0, ids: [] };
      linked.kg += kg;
      linked.ids.push(row.N_NVL_ID);
      receipts.set(row.DN_ID, linked);
    }
  } catch (err) {
    receiptWarning = 'Chưa đọc được tiến độ nhận. ' + err.message;
    receipts.clear();
  }
  return { master, operation, materials: material, suppliers, headers, lines, details,
    receipts, setup, receiptWarning, userId };
}

function filmOrders_(db) {
  const orders = new Map();
  const materialById = new Map(db.materials.rows.map(({ data }) => [data.VL_ID, data]));
  for (const { data } of db.headers ? db.headers.rows : []) {
    if (data.PO_NVL_ID) orders.set(data.PO_NVL_ID, { ...data, lines: [] });
  }
  for (const { data } of db.lines.rows) {
    if (!data.DN_ID) continue;
    const id = data.PO_NVL_ID || data.DN_ID;
    if (!orders.has(id)) {
      orders.set(id, { PO_NVL_ID: id, Ngay_Dat: data.Ngay_Dat, Ma_NCC: data.Ma_NCC,
        Trang_Thai: data.Trang_Thai, legacy: true, lines: [] });
    }
    const receipt = db.receipts.get(data.DN_ID);
    orders.get(id).lines.push({ ...data, material: materialById.get(data.VL_ID) || {},
      received: db.receiptWarning ? null : (receipt ? receipt.kg : 0),
      receiptIds: receipt ? receipt.ids : [] });
  }
  return Array.from(orders.values()).map(order => {
    order.ordered = order.lines.reduce((sum, line) => sum + (Number(line.So_Luong) || 0), 0);
    order.received = db.receiptWarning ? null : order.lines.reduce((sum, line) => sum + line.received, 0);
    order.progress = order.Trang_Thai === 'Hủy' ? 'Hủy' : db.receiptWarning ? 'Chưa rõ' :
      order.lines.length && order.lines.every(line =>
        Number(line.So_Luong) > 0 && line.received + 1e-8 >= Number(line.So_Luong)) ? 'Đã nhận đủ' :
      order.received > 0 ? 'Nhận một phần' : 'Chờ giao';
    order.revision = Utilities.base64EncodeWebSafe(Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256, JSON.stringify(order)));
    return order;
  }).sort((a, b) => (b.Ngay_Dat || '').localeCompare(a.Ngay_Dat || '') ||
    b.PO_NVL_ID.localeCompare(a.PO_NVL_ID));
}

function filmNumber_(value, positive, label) {
  const input = filmText_(value).replace(',', '.');
  if (!input && !positive) return '';
  const number = Number(input);
  if (!/^\d+(?:\.\d+)?$/.test(input) || !Number.isFinite(number) || (positive && number <= 0)) {
    throw new Error(label + ' không hợp lệ.');
  }
  return number;
}

function filmMaterialKey_(data) {
  const normalize = field => filmText_(data[field]).normalize('NFC').toLowerCase();
  const dimension = field => {
    const value = normalize(field).replace(',', '.');
    return /^\d+(?:\.\d+)?$/.test(value) ? Number(value) : value;
  };
  return JSON.stringify([normalize('Cấp Độ'), normalize('Màu Sắc'),
    normalize('Tĩnh Điện'), dimension('Độ Dày'), dimension('Khổ Nhựa')]);
}

function filmHighest_(table, column, prefix, digits) {
  let highest = 0;
  for (const { data } of table.rows) {
    const id = filmText_(data[column]);
    const suffix = id.slice(prefix.length);
    if (id.startsWith(prefix) && suffix.length === digits && /^\d+$/.test(suffix)) {
      highest = Math.max(highest, Number(suffix));
    }
  }
  return highest;
}

function filmSequence_(props, table, col, kind, day) {
  const prefix = kind + '_' + day + '_';
  const key = 'FILM_SEQUENCE_' + kind;
  const stored = (props.getProperty(key) || '').split(':');
  let highest = Math.max(filmHighest_(table, col, prefix, 2), stored[0] === day ? Number(stored[1]) : 0);
  return {
    next() {
      if (highest >= 99) throw new Error(kind + ' đã đạt giới hạn 99 mã trong ngày.');
      return prefix + String(++highest).padStart(2, '0');
    },
    reserve() {
      props.setProperty(key, day + ':' + highest);
    }
  };
}

function filmCell_(value, isDate) {
  if (value === '') return {};
  if (isDate) {
    const serial = Date.parse(value + 'T00:00:00Z') / 86400000 + 25569;
    return { userEnteredValue: { numberValue: serial },
      userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } };
  }
  return { userEnteredValue: typeof value === 'number' ? { numberValue: value } : { stringValue: value } };
}

function filmWrite_(requests, table, rowIndex, data) {
  const sheetId = table.sheet.getSheetId();
  if (rowIndex === null) {
    const cells = Array.from({ length: Math.max(...table.positions.values()) + 1 }, () => ({}));
    for (const [name, value] of Object.entries(data)) {
      cells[table.positions.get(name)] = filmCell_(value, name.startsWith('Ngay_'));
    }
    requests.push({ appendCells: { sheetId, rows: [{ values: cells }], fields: '*' } });
    return;
  }
  for (const [name, value] of Object.entries(data)) {
    requests.push({ updateCells: {
      start: { sheetId, rowIndex, columnIndex: table.positions.get(name) },
      rows: [{ values: [filmCell_(value, name.startsWith('Ngay_'))] }],
      fields: name.startsWith('Ngay_') && value ? 'userEnteredValue,userEnteredFormat.numberFormat' : 'userEnteredValue'
    } });
  }
}

function saveFilmOrder(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return filmSaveLocked_(payload);
  } finally {
    lock.releaseLock();
  }
}

function filmSaveLocked_(payload) {
  if (!payload || !Array.isArray(payload.lines) || !payload.lines.length || payload.lines.length > 99) {
    throw new Error('Đơn cần từ 1 đến 99 dòng quy cách.');
  }
  const db = filmDatabase_();
  if (db.setup.length) throw new Error(db.setup.join(' '));
  const props = PropertiesService.getScriptProperties();
  const orderId = filmText_(payload.id);
  const token = filmText_(payload.token);
  const retryKey = 'FILM_LAST_CREATE_' + db.userId;
  const lastCreate = (props.getProperty(retryKey) || '').split(':');
  if (!orderId && !/^[a-z0-9-]{16,64}$/i.test(token)) throw new Error('Mã phiên tạo đơn không hợp lệ.');
  if (!orderId && lastCreate[0] === token &&
      db.headers.rows.some(row => row.data.PO_NVL_ID === lastCreate[1])) {
    return { id: lastCreate[1] };
  }

  const existing = orderId ? filmOrders_(db).find(order => order.PO_NVL_ID === orderId) : null;
  if (orderId && (!existing || existing.legacy || existing.Trang_Thai === 'Hủy' ||
      existing.revision !== payload.revision)) {
    throw new Error('Đơn đã thay đổi hoặc không thể sửa. Hãy tải lại danh sách.');
  }
  if (existing && db.receiptWarning) throw new Error(db.receiptWarning);
  const supplier = filmText_(payload.supplier);
  if (!db.suppliers.some(row => row.Ma_NCC === supplier) && (!existing || existing.Ma_NCC !== supplier)) {
    throw new Error('Chọn nhà cung cấp màng nhựa hợp lệ.');
  }
  const date = filmInputDate_(payload.date, true);
  if (existing && existing.lines.some(line => line.receiptIds.length) &&
      (existing.Ma_NCC !== supplier || existing.Ngay_Dat !== date)) {
    throw new Error('Đơn đã có phiếu nhận. Không thể đổi nhà cung cấp hoặc ngày đặt.');
  }

  const day = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyMMdd');
  const poSequence = filmSequence_(props, db.headers, 'PO_NVL_ID', 'PO_NVL', day);
  const dnSequence = filmSequence_(props, db.lines, 'DN_ID', 'DN', day);
  const poId = orderId || poSequence.next();
  const oldLines = new Map(existing ? existing.lines.map(line => [line.DN_ID, line]) : []);
  const usedLines = new Set();
  const targetDetails = new Map();
  const detailRows = new Map();
  const existingAssignments = new Map();
  for (const row of db.details.rows) {
    if (!row.data.ID_Chi_Tiet) continue;
    detailRows.set(row.data.ID_Chi_Tiet, detailRows.has(row.data.ID_Chi_Tiet) ? null : row);
    if (oldLines.has(row.data.DN_ID)) {
      if (!existingAssignments.has(row.data.DN_ID)) existingAssignments.set(row.data.DN_ID, []);
      existingAssignments.get(row.data.DN_ID).push(row.data.ID_Chi_Tiet);
    }
  }

  const materials = new Map();
  let highestMaterial = filmHighest_(db.materials, 'VL_ID', 'VL', 4);
  const masterWrites = [];
  let nextMaterialRow = db.materials.sheet.getLastRow();
  const writes = [];
  for (const { data } of db.materials.rows) {
    if (!data['Độ Dày'] || !data['Khổ Nhựa']) continue;
    const key = filmMaterialKey_(data);
    const current = materials.get(key);
    if (!current || (!current.VL_ID && data.VL_ID)) materials.set(key, data);
  }
  for (const line of payload.lines) {
    if (!line) throw new Error('Dòng quy cách không hợp lệ.');
    const suppliedId = filmText_(line.id);
    const old = oldLines.get(suppliedId);
    if (suppliedId && (!old || usedLines.has(suppliedId))) throw new Error('DN_ID không thuộc đơn hoặc bị trùng.');
    const id = suppliedId || dnSequence.next();
    usedLines.add(id);
    const resin = filmText_(line.resin);
    if (!resin) throw new Error('Chọn loại nhựa.');
    const qty = filmNumber_(line.qty, true, 'Số lượng (kg)');
    const material = {
      'Cấp Độ': filmText_(line.grade), 'Màu Sắc': filmText_(line.color),
      'Tĩnh Điện': filmText_(line.esd),
      'Độ Dày': filmNumber_(line.thickness, true, 'Độ dày'),
      'Khổ Nhựa': filmNumber_(line.width, true, 'Khổ nhựa')
    };
    const key = filmMaterialKey_(material);
    const sameMaterial = old && filmMaterialKey_(old.material) === key;
    if (!['', '10^6-10^9', '10^9-10^11'].includes(material['Tĩnh Điện']) && !sameMaterial) {
      throw new Error('Giá trị tĩnh điện không hợp lệ.');
    }
    const matched = materials.get(key);
    let vl = sameMaterial ? old.VL_ID : (matched && matched.VL_ID);
    if (!vl) {
      if (highestMaterial >= 9999) throw new Error('Danh mục đã đạt giới hạn VL9999.');
      vl = 'VL' + String(++highestMaterial).padStart(4, '0');
      materials.set(key, { ...material, VL_ID: vl });
      const blank = matched && db.materials.rows.find(row => row.data === matched);
      filmWrite_(masterWrites, db.materials, blank ? blank.index : nextMaterialRow++,
        blank ? { VL_ID: vl } : { ...material, VL_ID: vl });
    }
    if (old && old.receiptIds.length &&
        (old.VL_ID !== vl || old.Loai_Nhua !== resin || qty + 1e-8 < old.received)) {
      throw new Error(id + ' đã có phiếu nhận. Giữ quy cách và số kg đặt ít nhất bằng số kg đã nhận.');
    }
    // Existing lines can save their spec without overwriting separately edited assignments.
    const assignedIds = existingAssignments.get(id) || [];
    if (old && line.originalIds !== undefined) {
      if (!Array.isArray(line.originalIds)) throw new Error('Danh sách CTDH ban đầu không hợp lệ.');
      const original = new Set(line.originalIds);
      if (assignedIds.length !== original.size || assignedIds.some(id => !original.has(id))) {
        throw new Error('CTDH của ' + id + ' đã thay đổi. Hãy mở lại đơn.');
      }
    }
    const detailIds = old && line.detailIds === undefined ? assignedIds : line.detailIds;
    if (!Array.isArray(detailIds)) throw new Error('Dòng quy cách không hợp lệ.');
    for (const detailId of detailIds) {
      const detail = detailRows.get(detailId);
      if (!detail || targetDetails.has(detailId)) throw new Error('CTDH thiếu hoặc bị gán trùng.');
      const prior = detail.data.DN_ID;
      if (prior !== id && (detail.data.Trang_Thai !== CONFIG.LINE_STATUS.NEW || (prior && !oldLines.has(prior)))) {
        throw new Error('CTDH ' + detailId + ' đã được gán hoặc đã qua trạng thái Chờ sản xuất.');
      }
      targetDetails.set(detailId, id);
    }
    const priorRow = old && db.lines.rows.find(row => row.data.DN_ID === id);
    const data = {
      DN_ID: id, PO_NVL_ID: poId, Loai_Nhua: resin, VL_ID: vl, So_Luong: qty,
      Don_Gia: filmNumber_(line.price, false, 'Đơn giá'),
      Ngay_Giao_Du_Kien: filmInputDate_(line.delivery, false), Ghi_Chu: filmText_(line.note),
      Trang_Thai: old ? old.Trang_Thai : 'Chờ nhập'
    };
    if (db.lines.positions.has('Ngay_Dat')) data.Ngay_Dat = date;
    if (db.lines.positions.has('Ma_NCC')) data.Ma_NCC = supplier;
    filmWrite_(writes, db.lines, priorRow ? priorRow.index : null, data);
  }

  if (Array.from(oldLines.keys()).some(id => !usedLines.has(id))) {
    throw new Error('Giữ các dòng đã lưu khi sửa đơn.');
  }
  for (const row of db.details.rows) {
    const id = row.data.ID_Chi_Tiet;
    const prior = row.data.DN_ID;
    const next = targetDetails.get(id) || '';
    if (prior === next || (!oldLines.has(prior) && !next)) continue;
    if (row.data.Trang_Thai !== CONFIG.LINE_STATUS.NEW || detailRows.get(id) !== row) {
      throw new Error('Không thể đổi gán CTDH ' + id + '. Hãy tải lại danh sách.');
    }
    filmWrite_(writes, db.details, row.index, { DN_ID: next });
  }
  const headerRow = existing && db.headers.rows.find(row => row.data.PO_NVL_ID === poId);
  filmWrite_(writes, db.headers, headerRow ? headerRow.index : null, {
    PO_NVL_ID: poId, Ngay_Dat: date, Ma_NCC: supplier,
    Nguoi_Dat: existing ? existing.Nguoi_Dat : db.userId, Ghi_Chu: filmText_(payload.note),
    Trang_Thai: existing ? existing.Trang_Thai : 'Chờ giao'
  });

  // Reserve before writing so an interrupted request cannot reissue an ID.
  poSequence.reserve();
  dnSequence.reserve();
  if (!orderId) props.setProperty(retryKey, token + ':' + poId);
  // The master is a separate workbook. A saved specification remains valid if
  // the operation batch fails, and the next attempt will match it by attributes.
  if (masterWrites.length) {
    const extraRows = nextMaterialRow - db.materials.sheet.getMaxRows();
    if (extraRows > 0) {
      masterWrites.unshift({ appendDimension: {
        sheetId: db.materials.sheet.getSheetId(), dimension: 'ROWS', length: extraRows
      } });
    }
    Sheets.Spreadsheets.batchUpdate({ requests: masterWrites }, db.master.getId());
  }
  Sheets.Spreadsheets.batchUpdate({ requests: writes }, db.operation.getId());
  return { id: poId };
}
