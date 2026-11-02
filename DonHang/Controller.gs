/**
 * FILE: Controller.gs
 * Điều phối sự kiện giao diện, Web Modal và Tự động điền dữ liệu
 */

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚡ QUẢN LÝ ĐƠN HÀNG')
    .addItem('➕ Tạo Đơn Hàng (Web Form)', 'openCreateOrderModal')
    .addItem('🔄 Đồng bộ Danh mục Sản phẩm & NV', 'manualSyncDropdowns')
    .addSeparator()
    .addItem('🧹 Xóa trắng biểu mẫu nhập', 'clearInputSheetAfterSave')
    .addToUi();
}

/**
 * HÀM LẤY DATA CHO MODAL
 * - Đọc cột Link Bản Vẽ / PO sẵn có từ sheet TAO DON
 * - Đối chiếu SP CŨ vs SP MỚI chuẩn xác qua Cache Master DB
 */
/**
 * THAY THẾ: Lấy dữ liệu từ Form TAO DON để hiển thị lên Modal
 * - Sử dụng resolveCol kết hợp Schema và Fallback trực tiếp để không bị trượt cột
 * - Đọc trực tiếp Cache A2:A vào Set trong RAM để kiểm tra Sản Phẩm Mới/Cũ
 */
function getOrderFormDataForModal() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return JSON.stringify({ poGroups: [], salesStaffList: [] });
    const sheet = ss.getSheetByName(CONFIG.SHEET_TAO_DON);
    if (!sheet) return JSON.stringify({ poGroups: [], salesStaffList: [] });

    // Quét tiêu đề dòng 7 sheet TAO DON
    const sheetInfo = getDynamicSheetInfo(sheet, CONFIG.HEADER_ROW_TAO_DON, 'Sản Phẩm');
    if (!sheetInfo || sheetInfo.lastRow < sheetInfo.startDataRow) {
      return JSON.stringify({ poGroups: [], salesStaffList: [] });
    }

    // Helper gộp từ khóa Schema và Fallback an toàn
    const resolveCol = (table, key, fallbacks) => {
      let headers = [];
      if (typeof getFormHeaders === 'function') {
        headers = getFormHeaders(table, key) || [];
      }
      const merged = Array.from(new Set([...headers, ...fallbacks]));
      return getColIndex(sheetInfo, merged);
    };

    // Đọc Cache kiểm tra Sản Phẩm và nạp Nhân viên KD
  let salesStaffList = [];
  const existingProductSet = new Set();
  const cacheProductMap = {};
  const cacheSheet = ss.getSheetByName(CONFIG.SHEET_CACHE);

  if (cacheSheet && cacheSheet.getLastRow() >= 2) {
    const cacheRowsCount = cacheSheet.getLastRow() - 1;
    const cacheAll = cacheSheet.getRange(2, 1, cacheRowsCount, Math.max(cacheSheet.getLastColumn(), 11)).getValues();

    for (let i = 0; i < cacheAll.length; i++) {
      const pName = String(cacheAll[i][0] || '').trim();
      if (pName) {
        const pKey = normalizeKey(pName);
        existingProductSet.add(pKey);
        cacheProductMap[pKey] = {
          codeSP: cacheAll[i][2] || '',
          unit: cacheAll[i][3] || '',
          price: cacheAll[i][4] || '',
          giaKhuon: cacheAll[i][5] || '',
          mang: cacheAll[i][6] || '',
          type: cacheAll[i][7] || '',
          donggoi: cacheAll[i][8] || ''
        };
      }

      const staffVal = String(cacheAll[i][10] || '').trim();
      if (staffVal) {
        const parts = staffVal.split(' - ');
        salesStaffList.push({ display: staffVal, maNV: parts[0] ? parts[0].trim() : staffVal, hoTen: parts[1] ? parts[1].trim() : staffVal });
      }
    }
  }

    const toYMD = function(val) {
      if (!val) return '';
      if (val instanceof Date) return Utilities.formatDate(val, "Asia/Ho_Chi_Minh", "yyyy-MM-dd");
      let s = String(val).trim();
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
      let parts = s.split('/');
      if (parts.length === 3) return `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
      return s;
    };

    const data = sheet.getRange(sheetInfo.startDataRow, 1, sheetInfo.lastRow - sheetInfo.startDataRow + 1, sheetInfo.lastCol).getValues();

    // Định vị cột trên form TAO DON bằng Fallback an toàn
    const c_sp       = resolveCol('CHI_TIET_DH', 'masp', ['Sản Phẩm', 'Tên Sản Phẩm', 'Mã Sản Phẩm']);
    const c_kh       = resolveCol('DON_HANG', 'maKH', ['Khách Hàng', 'Tên Khách Hàng', 'Ma_KH']);
    const c_code     = resolveCol('SAN_PHAM_MASTER', 'codeSanpham', ['Code Sản Phẩm', 'Code SP']);
    const c_sl       = resolveCol('CHI_TIET_DH', 'qty', ['Số Lượng', 'SL']);
    const c_gia      = resolveCol('CHI_TIET_DH', 'price', ['Đơn Giá', 'Price', 'Giá']);
    const c_po       = resolveCol('DON_HANG', 'soPO', ['Số PO', 'PO', 'So_PO']);
    const c_ngaydat  = resolveCol('DON_HANG', 'ngayDat', ['Ngày Đặt', 'Ngay_Dat']);
    const c_ngaygiao = resolveCol('CHI_TIET_DH', 'deliveryDate', ['Ngày Giao', 'Ngay_Giao']);
    const c_dvt      = resolveCol('SAN_PHAM_MASTER', 'donvi', ['Đơn Vị', 'ĐVT']);
    const c_nv       = resolveCol('CHI_TIET_DH', 'manv', ['Nhân Viên', 'Nhân Viên KD']);
    const c_ghichu   = resolveCol('CHI_TIET_DH', 'note', ['Ghi chú', 'Ghi Chú']);
    const c_link     = resolveCol('DON_HANG', 'fileUrl', ['Link PO', 'Link Bản vẽ/PO', 'Link']);
    const c_mang     = resolveCol('SAN_PHAM_MASTER', 'mang', ['Màng']);
    const c_giakhuon = resolveCol('SAN_PHAM_MASTER', 'giakhuon', ['Giá Khuôn', 'Gia Khuon']);

    let poGroupsMap = {};

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const sp = c_sp >= 0 ? String(row[c_sp] || '').trim() : '';
      if (!sp) continue;

      const kh = c_kh >= 0 ? String(row[c_kh] || '').trim() : '';
      const po = (c_po >= 0 && String(row[c_po] || '').trim()) ? String(row[c_po]).trim() : generateAutoPO(kh, new Date());
      const ngayDatYMD = c_ngaydat >= 0 ? toYMD(row[c_ngaydat]) : '';
      const ngayGiaoYMD = c_ngaygiao >= 0 ? toYMD(row[c_ngaygiao]) : '';
      const staffVal = c_nv >= 0 ? String(row[c_nv] || '').trim() : '';
      const existingLinkVal = c_link >= 0 ? String(row[c_link] || '').trim() : '';

      if (!poGroupsMap[po]) {
        poGroupsMap[po] = {
          customerName: kh,
          soPO: po,
          autoPONumber: po,
          salesPerson: staffVal,
          ngayDat: ngayDatYMD,
          ngayGiaoDK: ngayGiaoYMD,
          orderNote: c_ghichu >= 0 ? String(row[c_ghichu] || '').trim() : '',
          existingFileUrl: existingLinkVal,
          items: []
        };
      }

      if (!poGroupsMap[po].existingFileUrl && existingLinkVal) {
        poGroupsMap[po].existingFileUrl = existingLinkVal;
      }

      const codeVal = c_code >= 0 ? String(row[c_code] || '').trim() : '';
      const slVal = c_sl >= 0 ? (Number(row[c_sl]) || 0) : 0;
      const giaVal = c_gia >= 0 ? (Number(row[c_gia]) || 0) : 0;
      const dvtVal = c_dvt >= 0 ? String(row[c_dvt] || '').trim() : '';
      const noteVal = c_ghichu >= 0 ? String(row[c_ghichu] || '').trim() : '';
      const mangVal = c_mang >= 0 ? String(row[c_mang] || '').trim() : '';
      const isNewProduct = !existingProductSet.has(normalizeKey(sp));
      const giaKhuonVal = c_giakhuon >= 0 ? (Number(row[c_giakhuon]) || 0) : 0;

      const cachedItem = cacheProductMap[normalizeKey(sp)] || {};
      poGroupsMap[po].items.push({
        codeSP: codeVal || cachedItem.codeSP || '',
        prodName: sp,
        unit: dvtVal || cachedItem.unit || '',
        qty: slVal,
        price: giaVal || cachedItem.price || 0,
        giaKhuon: giaKhuonVal || cachedItem.giaKhuon || 0,
        mang: mangVal || cachedItem.mang || '',
        type: cachedItem.type || '',
        donggoi: cachedItem.donggoi || '',
        note: noteVal,
        deliveryDate: ngayGiaoYMD,
        isNewProduct: isNewProduct
      });
    }

    return JSON.stringify({ poGroups: Object.values(poGroupsMap), salesStaffList: salesStaffList });
  } catch (err) {
    Logger.log("Lỗi getOrderFormDataForModal: " + err.toString());
    return JSON.stringify({ poGroups: [], salesStaffList: [] });
  }
}

/**
 * TỰ ĐỘNG ĐIỀN DỮ LIỆU KHI CHỌN SẢN PHẨM HOẶC SỬA FORM TAO DON
 * - Cơ chế Fallback đa tầng: Ưu tiên getFormHeaders từ Schema kết hợp từ khóa chuẩn thực tế
 * - Đọc trực tiếp ma trận Cache A2:E vào RAM (Batch Processing), không quét lệch tiêu đề
 * - Batch Write 1 lần duy nhất bằng setValues()
 */
/**
 * SỰ KIỆN ONEDIT: TỰ ĐỘNG ĐIỀN THÔNG TIN KHI CHỌN SẢN PHẨM TRÊN SHEET TAO DON
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SHEET_TAO_DON) return;

  // 1. Quét thông tin dòng tiêu đề sheet TAO DON (Dòng 7)
  const sheetInfo = getDynamicSheetInfo(sheet, CONFIG.HEADER_ROW_TAO_DON, 'Sản Phẩm');
  const rowStart = e.range.getRow();
  const colStart = e.range.getColumn();
  const numRows = e.range.getNumRows();
  const colEnd = colStart + e.range.getNumColumns() - 1;

  if (rowStart < sheetInfo.startDataRow || colStart > sheetInfo.lastCol) return;

  // Helper an toàn gộp từ khóa Schema và Fallback trực tiếp
  const resolveCol = (table, key, fallbacks) => {
    let headers = [];
    if (typeof getFormHeaders === 'function') {
      headers = getFormHeaders(table, key) || [];
    }
    const merged = Array.from(new Set([...headers, ...fallbacks]));
    return getColIndex(sheetInfo, merged);
  };

  // Xác định vị trí cột động trên TAO DON
  const idxSP       = resolveCol('CHI_TIET_DH', 'prodName', ['Sản Phẩm', 'Tên Sản Phẩm']);
  const idxKH       = resolveCol('DON_HANG', 'customerName', ['Khách Hàng', 'Tên Khách Hàng']);
  const idxCode     = resolveCol('CHI_TIET_DH', 'codeSP', ['Code Sản Phẩm', 'Code SP']);
  const idxDVT      = resolveCol('CHI_TIET_DH', 'unit', ['Đơn Vị', 'ĐVT']);
  const idxGia      = resolveCol('CHI_TIET_DH', 'price', ['Đơn Giá', 'Giá Bán', 'Giá']);
  const idxNV       = resolveCol('DON_HANG', 'salesPerson', ['Nhân Viên', 'Nhân Viên KD']);
  const idxNgayDat  = resolveCol('DON_HANG', 'ngayDat', ['Ngày Đặt', 'Ngay Dat']);
  const idxNgayGiao = resolveCol('CHI_TIET_DH', 'deliveryDate', ['Ngày Giao', 'Ngay Giao']);
  const idxPO       = resolveCol('DON_HANG', 'soPO', ['Số PO', 'PO']);
  const idxSL       = resolveCol('CHI_TIET_DH', 'qty', ['Số Lượng', 'SL']);
  const idxThanhTien= resolveCol('CHI_TIET_DH', 'lineTotal', ['Thành Tiền']);

  // Kiểm tra phạm vi chỉnh sửa của người dùng
  const isSpEdited  = (idxSP >= 0 && colStart <= (idxSP + 1) && colEnd >= (idxSP + 1));
  const isKhEdited  = (idxKH >= 0 && colStart <= (idxKH + 1) && colEnd >= (idxKH + 1));
  const isSlEdited  = (idxSL >= 0 && colStart <= (idxSL + 1) && colEnd >= (idxSL + 1));
  const isGiaEdited = (idxGia >= 0 && colStart <= (idxGia + 1) && colEnd >= (idxGia + 1));

  if (!isSpEdited && !isKhEdited && !isSlEdited && !isGiaEdited) return;

  // 2. Đọc ma trận vùng dữ liệu đang chỉnh sửa
  const targetRange = sheet.getRange(rowStart, 1, numRows, sheetInfo.lastCol);
  const targetValues = targetRange.getValues();

  // 3. Nạp danh mục từ _CACHE_DANH_MUC vào RAM (Đọc trực tiếp A2:E)
  let productDict = {};
  if (isSpEdited) {
    const cacheSheet = sheet.getParent().getSheetByName(CONFIG.SHEET_CACHE);
    if (cacheSheet && cacheSheet.getLastRow() >= 2) {
      const cacheRows = cacheSheet.getLastRow() - 1;
      const cacheData = cacheSheet.getRange(2, 1, cacheRows, 5).getValues();

      // Cột A: Tên SP | B: Khách Hàng | C: Code SP | D: Đơn Vị | E: Đơn Giá
      for (let i = 0; i < cacheData.length; i++) {
        const spName = String(cacheData[i][0] || '').trim();
        if (spName) {
          productDict[normalizeKey(spName)] = {
            khachHang: cacheData[i][1] || '',
            codeSP: cacheData[i][2] || '',
            dvt: cacheData[i][3] || '',
            donGia: cacheData[i][4] || ''
          };
        }
      }
    }
  }

  const todayDateObj = new Date();
  const deliveryDateObj = new Date(todayDateObj.getTime() + (CONFIG.DEFAULT_DELIVERY_DAYS * 24 * 60 * 60 * 1000));
  let hasChanges = false;

  // 4. Duyệt mảng và tính toán nội suy trong RAM
  for (let r = 0; r < numRows; r++) {
    // NGHIỆP VỤ 1: CHỌN / SỬA SẢN PHẨM
    if (isSpEdited && idxSP >= 0) {
      const spInput = String(targetValues[r][idxSP] || '').trim();
      if (spInput) {
        const info = productDict[normalizeKey(spInput)];

        if (idxNgayDat >= 0) targetValues[r][idxNgayDat] = todayDateObj;
        if (idxNgayGiao >= 0) targetValues[r][idxNgayGiao] = deliveryDateObj;

        if (info) {
          if (idxKH >= 0) targetValues[r][idxKH] = info.khachHang;
          if (idxCode >= 0) targetValues[r][idxCode] = info.codeSP;
          if (idxDVT >= 0 && idxDVT !== idxGia) targetValues[r][idxDVT] = info.dvt;
          if (idxGia >= 0 && idxGia !== idxDVT) targetValues[r][idxGia] = info.donGia;
          if (idxPO >= 0) targetValues[r][idxPO] = generateAutoPO(info.khachHang, todayDateObj);
        } else {
          // Sản phẩm mới: xóa trắng các ô liên quan để nhập mới
          if (idxKH >= 0) targetValues[r][idxKH] = '';
          if (idxCode >= 0) targetValues[r][idxCode] = '';
          if (idxDVT >= 0 && idxDVT !== idxGia) targetValues[r][idxDVT] = '';
          if (idxGia >= 0 && idxGia !== idxDVT) targetValues[r][idxGia] = '';
          if (idxPO >= 0) targetValues[r][idxPO] = '';
        }
        hasChanges = true;
      } else {
        // Xóa sản phẩm: dọn sạch hàng
        if (idxKH >= 0) targetValues[r][idxKH] = '';
        if (idxCode >= 0) targetValues[r][idxCode] = '';
        if (idxDVT >= 0) targetValues[r][idxDVT] = '';
        if (idxGia >= 0) targetValues[r][idxGia] = '';
        if (idxNV >= 0) targetValues[r][idxNV] = '';
        if (idxNgayDat >= 0) targetValues[r][idxNgayDat] = '';
        if (idxNgayGiao >= 0) targetValues[r][idxNgayGiao] = '';
        if (idxPO >= 0) targetValues[r][idxPO] = '';
        if (idxThanhTien >= 0) targetValues[r][idxThanhTien] = '';
        hasChanges = true;
      }
    }

    // NGHIỆP VỤ 2: THAY ĐỔI KHÁCH HÀNG -> SINH LẠI SỐ PO
    if (isKhEdited && idxKH >= 0 && idxPO >= 0) {
      const khInput = String(targetValues[r][idxKH] || '').trim();
      if (khInput) {
        targetValues[r][idxPO] = generateAutoPO(khInput, todayDateObj);
        hasChanges = true;
      }
    }

    // NGHIỆP VỤ 3: TỰ TÍNH THÀNH TIỀN = SỐ LƯỢNG * ĐƠN GIÁ
    if (idxThanhTien >= 0 && idxSL >= 0 && idxGia >= 0) {
      const sl = parseFloat(targetValues[r][idxSL]) || 0;
      const gia = parseFloat(targetValues[r][idxGia]) || 0;
      if (sl > 0 && gia > 0) {
        targetValues[r][idxThanhTien] = sl * gia;
        hasChanges = true;
      } else {
        targetValues[r][idxThanhTien] = '';
        hasChanges = true;
      }
    }
  }

  // 5. Ghi hàng loạt xuống Sheet 1 lần duy nhất
  if (hasChanges) {
    targetRange.setValues(targetValues);
  }
}

function openCreateOrderModal() {
  try {
    // 1. Tạo Template từ file HTML
    const template = HtmlService.createTemplateFromFile('OrderModal');
    
    // 2. Lấy dữ liệu đơn hàng từ sheet TAO DON dạng JSON chuỗi
    const rawJson = getOrderFormDataForModal();
    
    // 3. Mã hóa Base64 để chống lỗi ký tự đặc biệt, dấu nháy kép, xuống dòng khi truyền sang client
    template.initialDataJson = Utilities.base64Encode(rawJson, Utilities.Charset.UTF_8);
    
    // 4. Đánh giá template và mở Dialog
    const htmlOutput = template.evaluate()
      .setWidth(1150)
      .setHeight(700)
      .setTitle('TẠO ĐƠN HÀNG MỚI - NAPACO MANAGEMENT');
      
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'TẠO ĐƠN HÀNG MỚI - NAPACO MANAGEMENT');
  } catch (err) {
    Logger.log("❌ Lỗi openCreateOrderModal: " + err.message);
    SpreadsheetApp.getUi().alert("Lỗi mở giao diện tạo đơn: " + err.message);
  }
}


function clearInputSheetAfterSave() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_TAO_DON);
  if (!sheet) return;
  const sheetInfo = getDynamicSheetInfo(sheet, CONFIG.HEADER_ROW_TAO_DON, 'Sản Phẩm');
  sheet.getRange(sheetInfo.startDataRow, 1, CONFIG.MAX_ROWS_TAO_DON, sheetInfo.lastCol).clearContent();
  SpreadsheetApp.getActiveSpreadsheet().toast('Bảng nhập liệu đã được làm sạch!', '✅ Hoàn tất');
}

function saveOrderFromModal(payload) {
  try {
    if (!payload) {
      throw new Error("Payload rỗng hoặc không hợp lệ từ giao diện.");
    }

    // 1. Chuẩn hóa cấu trúc payload: trích xuất mảng danh sách đơn hàng
    const poGroups = payload.poGroups || payload.orders || (Array.isArray(payload) ? payload : null);

    if (!poGroups || !Array.isArray(poGroups) || poGroups.length === 0) {
      throw new Error("Không tìm thấy dữ liệu đơn PO nào trong payload để lưu.");
    }

    // 2. Gọi hàm thực thi backend ghi database
    const rawResult = executeSaveOrdersWithSchema({ poGroups: poGroups });

    // 3. Chuẩn hóa kết quả trả về dưới dạng Plain Object (chống lỗi serialize của google.script.run)
    const safeResult = {
      success: true,
      orderCount: (rawResult && typeof rawResult.orderCount === 'number') ? rawResult.orderCount : poGroups.length,
      itemCount: (rawResult && typeof rawResult.itemCount === 'number') ? rawResult.itemCount : 0,
      newProductCount: (rawResult && typeof rawResult.newProductCount === 'number') ? rawResult.newProductCount : 0,
      summaryText: (rawResult && rawResult.summaryText) ? String(rawResult.summaryText) : ''
     // newCacheRows: (rawResult && Array.isArray(rawResult.newCacheRows)) ? rawResult.newCacheRows : []
    };

    Logger.log(`✅ [CONTROLLER] Lưu thành công: ${safeResult.orderCount} PO, ${safeResult.itemCount} items.`);
    return safeResult;

  } catch (err) {
    Logger.log("❌ Lỗi tại saveOrderFromModal: " + err.message);
    throw new Error(err.message);
  }
}

/**
 * Chạy ngầm dọn dẹp form và đồng bộ danh mục sau khi đóng Modal
 * Được gọi bất đồng bộ từ Client (Fire-and-forget), không làm nghẽn giao diện người dùng
 */
function runPostSaveTasks(newCacheRows) {
  try {
    Logger.log("🚀 [BACKGROUND] Bắt đầu tác vụ ngầm: Dọn dẹp form & Đồng bộ danh mục...");
    
    // 1. Tối ưu mở file Operation qua phiên Active hoặc ID
    const activeSS = SpreadsheetApp.getActiveSpreadsheet();
    const ssOp = (activeSS && activeSS.getId() === CONFIG.ID_OPERATION)
      ? activeSS
      : SpreadsheetApp.openById(CONFIG.ID_OPERATION);

    // 2. Xóa dữ liệu đã nhập trên sheet TAO DON (từ dòng 8 trở xuống)
    const sheetTaoDon = ssOp.getSheetByName(CONFIG.SHEET_TAO_DON);
    if (sheetTaoDon) {
      const startR = (CONFIG.HEADER_ROW_TAO_DON || 7) + 1; // Dòng 8
      const curLastRow = sheetTaoDon.getLastRow();
      if (curLastRow >= startR) {
        const numRows = curLastRow - startR + 1;
        const numCols = Math.max(sheetTaoDon.getLastColumn(), 26);
        // Xóa trắng toàn bộ nội dung dữ liệu vừa nhập, giữ nguyên định dạng, dropdown & công thức
        sheetTaoDon.getRange(startR, 1, numRows, numCols).clearContent();
        Logger.log(`🧹 [BACKGROUND] Đã xóa ${numRows} dòng dữ liệu trên sheet ${CONFIG.SHEET_TAO_DON}.`);
      }
    }

    // 3. Đồng bộ danh mục (Fast-Append vào Cache) nếu có sản phẩm/khách hàng mới phát sinh
    if (newCacheRows && Array.isArray(newCacheRows) && newCacheRows.length > 0) {
      const sheetCache = ssOp.getSheetByName(CONFIG.SHEET_CACHE_PRODUCTS || 'CACHE_SAN_PHAM');
      if (sheetCache) {
        const lastRowCache = sheetCache.getLastRow();
        const startWriteRow = lastRowCache + 1;
        sheetCache.getRange(startWriteRow, 1, newCacheRows.length, newCacheRows[0].length).setValues(newCacheRows);
        Logger.log(`⚡ [BACKGROUND] Đã đồng bộ nối đuôi ${newCacheRows.length} sản phẩm mới vào Cache.`);
      }
    }

    Logger.log("✅ [BACKGROUND] Hoàn tất toàn bộ tác vụ ngầm thành công.");
  } catch (err) {
    Logger.log("⚠️ [BACKGROUND] Lỗi trong tác vụ ngầm: " + err.message);
  }
}

function debugOnEditColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_TAO_DON);
  if (!sheet) {
    Logger.log("❌ Không tìm thấy sheet: " + CONFIG.SHEET_TAO_DON);
    return;
  }
  
  const anchorSP = getFormHeaders('CHI_TIET_DH', 'prodName')[0] || 'Sản Phẩm';
  const sheetInfo = getDynamicSheetInfo(sheet, CONFIG.HEADER_ROW_TAO_DON, anchorSP);
  
  Logger.log("--- KẾT QUẢ QUÉT CỘT FORM TAO DON ---");
  Logger.log("Dòng tiêu đề nhận diện được: " + sheetInfo.headerRow);
  Logger.log("Cột Sản Phẩm (idxSP): " + getColIndex(sheetInfo, getFormHeaders('CHI_TIET_DH', 'prodName')));
  Logger.log("Cột Khách Hàng (idxKH): " + getColIndex(sheetInfo, getFormHeaders('DON_HANG', 'customerName')));
  Logger.log("Cột Code SP (idxCode): " + getColIndex(sheetInfo, getFormHeaders('CHI_TIET_DH', 'codeSP')));
  Logger.log("Cột Đơn Vị (idxDVT): " + getColIndex(sheetInfo, getFormHeaders('CHI_TIET_DH', 'unit')));
  Logger.log("Cột Đơn Giá (idxGia): " + getColIndex(sheetInfo, getFormHeaders('CHI_TIET_DH', 'price')));
  Logger.log("Cột Số PO (idxPO): " + getColIndex(sheetInfo, getFormHeaders('DON_HANG', 'soPO')));
  
  const cacheSheet = ss.getSheetByName(CONFIG.SHEET_CACHE);
  if (!cacheSheet) {
    Logger.log("❌ Không tìm thấy sheet Cache: " + CONFIG.SHEET_CACHE);
  } else {
    Logger.log("--- KIỂM TRA SHEET CACHE ---");
    Logger.log("Số dòng Cache hiện có: " + cacheSheet.getLastRow());
    if (cacheSheet.getLastRow() >= 1) {
      Logger.log("Tiêu đề dòng 1 Cache: " + JSON.stringify(cacheSheet.getRange(1, 1, 1, Math.min(cacheSheet.getLastColumn(), 7)).getValues()[0]));
    }
  }
}

/**
 * HÀM TEST KIỂM TRA CHỈ SỐ CỘT SAU KHI CẬP NHẬT
 * Mục tiêu: Tất cả các cột phải trả về index >= 0 (không còn bị -1)
 */
function verifyOnEditSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_TAO_DON);
  if (!sheet) {
    Logger.log("❌ Không tìm thấy sheet: " + CONFIG.SHEET_TAO_DON);
    return;
  }

  const sheetInfo = getDynamicSheetInfo(sheet, CONFIG.HEADER_ROW_TAO_DON, 'Sản Phẩm');
  
  const resolveCol = (table, key, fallbacks) => {
    let headers = [];
    if (typeof getFormHeaders === 'function') {
      headers = getFormHeaders(table, key) || [];
    }
    const merged = Array.from(new Set([...headers, ...fallbacks]));
    return getColIndex(sheetInfo, merged);
  };

  Logger.log("=== KẾT QUẢ ĐỐI SOÁT CỘT SHEET TAO DON ===");
  Logger.log("Dòng tiêu đề nhận dạng: " + sheetInfo.headerRow);
  Logger.log("Cột Sản Phẩm (idxSP): "       + resolveCol('CHI_TIET_DH', 'prodName', ['Sản Phẩm', 'Tên Sản Phẩm']));
  Logger.log("Cột Khách Hàng (idxKH): "     + resolveCol('DON_HANG', 'customerName', ['Khách Hàng', 'Tên Khách Hàng']));
  Logger.log("Cột Code SP (idxCode): "      + resolveCol('CHI_TIET_DH', 'codeSP', ['Code Sản Phẩm', 'Code SP']));
  Logger.log("Cột Đơn Vị (idxDVT): "        + resolveCol('CHI_TIET_DH', 'unit', ['Đơn Vị', 'ĐVT']));
  Logger.log("Cột Đơn Giá (idxGia): "       + resolveCol('CHI_TIET_DH', 'price', ['Đơn Giá', 'Giá Bán', 'Giá']));
  Logger.log("Cột Số PO (idxPO): "          + resolveCol('DON_HANG', 'soPO', ['Số PO', 'PO']));
  Logger.log("Cột Số Lượng (idxSL): "       + resolveCol('CHI_TIET_DH', 'qty', ['Số Lượng', 'SL']));
  Logger.log("Cột Thành Tiền (idxThanhTien): " + resolveCol('CHI_TIET_DH', 'lineTotal', ['Thành Tiền']));
}

/**
 * HÀM TEST ĐỘC LẬP: Kiểm tra dữ liệu đẩy lên Modal Web Form
 */
function testModalDataPayload() {
  Logger.log("=== BẮT ĐẦU KIỂM TRA GETORDERFORMDATAFORMODAL ===");
  try {
    const rawJson = getOrderFormDataForModal();
    Logger.log("1. Kết quả getOrderFormDataForModal: " + (rawJson ? "THÀNH CÔNG" : "RỖNG"));
    Logger.log("2. Độ dài chuỗi JSON: " + (rawJson ? rawJson.length : 0) + " ký tự");
    
    if (rawJson) {
      const parsed = JSON.parse(rawJson);
      Logger.log("3. Số nhóm PO tìm thấy: " + (parsed.poGroups ? parsed.poGroups.length : 0));
      Logger.log("4. Số nhân viên KD tìm thấy: " + (parsed.salesStaffList ? parsed.salesStaffList.length : 0));
      
      // Test thử việc Base64 Encode
      const encoded = Utilities.base64Encode(rawJson, Utilities.Charset.UTF_8);
      Logger.log("5. Mã hóa Base64: THÀNH CÔNG (" + encoded.length + " bytes)");
    }
  } catch (err) {
    Logger.log("❌ LỖI NGHIÊM TRỌNG TẠI SERVER: " + err.toString());
    Logger.log("Dòng lỗi: " + err.stack);
  }
}

