const FilmColumns_ = {
  po: ['PO_NVL_ID', 'Ngay_Dat', 'Ma_NCC', 'Nguoi_Dat', 'Ghi_Chu', 'Trang_Thai'],
  line: ['DN_ID', 'Loai_Nhua', 'VL_ID', 'So_Luong', 'Don_Gia', 'Ghi_Chu', 'Trang_Thai'],
  lineNew: ['PO_NVL_ID', 'Ngay_Giao_Du_Kien'],
  material: ['VL_ID', 'Cấp Độ', 'Màu Sắc', 'Tĩnh Điện', 'Độ Dày', 'Khổ Nhựa'],
  supplier: ['Ma_NCC', 'Ten_NCC', 'Loai_Cung_Cap'],
  receipt: ['N_NVL_ID', 'DN_ID', 'So_KG_Nhan']
};

function getFilmOrders() {
  const db = filmDatabase_();
  return {
    orders: filmOrders_(db),
    suppliers: db.suppliers,
    materials: db.materials.rows.map(row => row.data).filter(row => row.VL_ID),
    details: db.details.rows.map(row => row.data).filter(row =>
      row.ID_Chi_Tiet && (row.Trang_Thai === CONFIG.LINE_STATUS.NEW || row.DN_ID)),
    today: filmDate_(new Date()),
    setup: db.setup,
    receiptWarning: db.receiptWarning
  };
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

function filmTable_(book, names, columns, optional = []) {
  const sheet = names.map(name => book.getSheetByName(name)).find(Boolean);
  if (!sheet) throw new Error('Thiếu tab ' + names[0] + '.');
  const values = sheet.getDataRange().getValues();
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
  const rows = values.slice(index + 1).map((cells, offset) => {
    const data = {};
    for (const [name, col] of positions) data[name] = filmDate_(cells[col]);
    return { index: index + offset + 1, data };
  });
  return { sheet, positions, rows };
}

function filmDatabase_() {
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

  const material = filmTable_(master, [names.FILM_MATERIALS, '01_04_ThongSoVatLieu'], FilmColumns_.material);
  const suppliers = filmTable_(master, [names.FILM_SUPPLIERS, '01_03_DM_NhaCungCap'], FilmColumns_.supplier)
    .rows.map(row => row.data).filter(row => /^NCC(0[1-9]|[1-9]\d)$/.test(row.Ma_NCC) &&
      /màng/i.test(row.Loai_Cung_Cap));
  const lines = filmTable_(operation, [names.FILM_LINES], FilmColumns_.line,
    FilmColumns_.lineNew.concat('Ngay_Dat', 'Ma_NCC'));
  const detail = CONFIG.COLUMNS.ORDER_DETAILS;
  const details = filmTable_(operation, [names.ORDER_DETAILS],
    [detail.ID, detail.STATUS], [detail.FILM_LINE_ID, detail.ORDER_ID, detail.PRODUCT_ID,
      detail.QTY_ORDER, detail.DELIVERY_DATE]);
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
    receipts, setup, receiptWarning, userId: user[employee.ID] };
}

function filmOrders_(db) {
  const orders = new Map();
  const materialById = new Map(db.materials.rows.map(({ data }) => [data.VL_ID, data]));
  const assigned = new Map();
  for (const { data } of db.details.rows) {
    if (!data.DN_ID || !data.ID_Chi_Tiet) continue;
    if (!assigned.has(data.DN_ID)) assigned.set(data.DN_ID, []);
    assigned.get(data.DN_ID).push(data.ID_Chi_Tiet);
  }
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
      detailIds: assigned.get(data.DN_ID) || [],
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
  for (const row of db.details.rows) {
    if (!row.data.ID_Chi_Tiet) continue;
    detailRows.set(row.data.ID_Chi_Tiet, detailRows.has(row.data.ID_Chi_Tiet) ? null : row);
  }

  const materials = new Map();
  let highestMaterial = filmHighest_(db.materials, 'VL_ID', 'VL', 4);
  const masterWrites = [];
  const writes = [];
  for (const { data } of db.materials.rows) {
    if (!data['Độ Dày'] || !data['Khổ Nhựa']) continue;
    const key = filmMaterialKey_(data);
    const current = materials.get(key);
    if (!current || (!current.VL_ID && data.VL_ID)) materials.set(key, data);
  }
  for (const line of payload.lines) {
    if (!line || !Array.isArray(line.detailIds)) throw new Error('Dòng quy cách không hợp lệ.');
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
      filmWrite_(masterWrites, db.materials, blank ? blank.index : null,
        blank ? { VL_ID: vl } : { ...material, VL_ID: vl });
    }
    if (old && old.receiptIds.length &&
        (old.VL_ID !== vl || old.Loai_Nhua !== resin || qty + 1e-8 < old.received)) {
      throw new Error(id + ' đã có phiếu nhận. Giữ quy cách và số kg đặt ít nhất bằng số kg đã nhận.');
    }
    for (const detailId of line.detailIds) {
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
  if (masterWrites.length) Sheets.Spreadsheets.batchUpdate({ requests: masterWrites }, db.master.getId());
  Sheets.Spreadsheets.batchUpdate({ requests: writes }, db.operation.getId());
  return { id: poId };
}
