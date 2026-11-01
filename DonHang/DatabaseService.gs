/**
 * FILE: DatabaseService.gs
 * Chuyên trách tương tác với DB Master và Cache, thực thi lưu đơn theo TABLE_MAPPING_SCHEMA
 */

function manualSyncDropdowns() {
  syncFormDropdowns();
  SpreadsheetApp.getActiveSpreadsheet().toast('Đã cập nhật Danh mục & Nhân viên!', '✅ Hoàn tất');
}

function syncFormDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cacheSheet = ss.getSheetByName(CONFIG.SHEET_CACHE);
  if (!cacheSheet) cacheSheet = ss.insertSheet(CONFIG.SHEET_CACHE);

  const masterSS = SpreadsheetApp.openById(CONFIG.ID_MASTER);

  // --- 1. ĐỌC KHÁCH HÀNG ---
  const khSheet = masterSS.getSheetByName(CONFIG.SHEET_KHACH_HANG);
  const khInfo = getDynamicSheetInfo(khSheet, 1, 'Ma_KH');
  const khMap = {};
  if (khInfo.lastRow >= khInfo.startDataRow) {
    const khData = khSheet.getRange(khInfo.startDataRow, 1, khInfo.lastRow - khInfo.startDataRow + 1, khInfo.lastCol).getValues();
    const c_makh = getColIndex(khInfo, ['Ma_KH', 'customer_id']);
    const c_namekh = getColIndex(khInfo, ['name', 'Tên Khách Hàng', 'Khách Hàng']);
    for (let i = 0; i < khData.length; i++) {
      if (c_makh >= 0) {
        const ma = String(khData[i][c_makh]).trim();
        if (ma) khMap[ma] = c_namekh >= 0 ? String(khData[i][c_namekh]).trim() : '';
      }
    }
  }

  // --- 2. ĐỌC NHÂN VIÊN KINH DOANH ---
  const nvSheet = masterSS.getSheetByName(CONFIG.SHEET_NHAN_VIEN);
  const listNvKinhDoanh = [];
  if (nvSheet) {
    const nvInfo = getDynamicSheetInfo(nvSheet, 6, 'Mã NV');
    if (nvInfo.lastRow >= nvInfo.startDataRow) {
      const nvData = nvSheet.getRange(nvInfo.startDataRow, 1, nvInfo.lastRow - nvInfo.startDataRow + 1, nvInfo.lastCol).getValues();
      const c_manv = getColIndex(nvInfo, ['Mã NV']);
      const c_tennv = getColIndex(nvInfo, ['Họ và tên']);
      const c_phongban = getColIndex(nvInfo, ['Phòng ban']);
      const c_trangthai = getColIndex(nvInfo, ['Trạng thái']);
      for (let i = 0; i < nvData.length; i++) {
        if (c_manv >= 0 && c_tennv >= 0) {
          const ma = String(nvData[i][c_manv]).trim();
          const ten = String(nvData[i][c_tennv]).trim();
          const pb = c_phongban >= 0 ? String(nvData[i][c_phongban]).trim().toLowerCase() : '';
          const tt = c_trangthai >= 0 ? String(nvData[i][c_trangthai]).trim().toLowerCase() : '';
          if (ma && pb.includes('kinh doanh') && tt.includes('đang hoạt động')) {
            listNvKinhDoanh.push([`${ma} - ${ten}`]);
          }
        }
      }
    }
  }

  // --- 3. ĐỌC SẢN PHẨM & GHI VÀO CACHE ---
  const spSheet = masterSS.getSheetByName(CONFIG.SHEET_SAN_PHAM);
  if (!spSheet) return;

  const spInfo = getDynamicSheetInfo(spSheet, 1, 'name_napaco');
  if (spInfo.lastRow < spInfo.startDataRow) return;

  const spData = spSheet.getRange(spInfo.startDataRow, 1, spInfo.lastRow - spInfo.startDataRow + 1, spInfo.lastCol).getValues();
  const c_sp = getColIndex(spInfo, ['name_napaco', 'Tên Sản Phẩm']);
  const c_kh = getColIndex(spInfo, ['customer_id', 'Khách Hàng']);
  const c_code = getColIndex(spInfo, ['Code_Sanpham', 'Code Sản Phẩm']);
  const c_dvt = getColIndex(spInfo, ['donvi', 'Đơn Vị']);
  const c_gia = getColIndex(spInfo, ['dongia', 'Đơn Giá']);

  const cacheHeadersSP = [['Tên Sản Phẩm', 'Khách Hàng', 'Code Sản Phẩm', 'Đơn Vị', 'Đơn Giá']];
  const cacheRowsSP = [];
  for (let i = 0; i < spData.length; i++) {
    const row = spData[i];
    const tenSP = c_sp >= 0 ? row[c_sp] : '';
    if (String(tenSP).trim() !== '') {
      const customerId = c_kh >= 0 ? String(row[c_kh]).trim() : '';
      cacheRowsSP.push([
        tenSP,
        khMap[customerId] || customerId,
        c_code >= 0 ? row[c_code] : '',
        c_dvt >= 0 ? row[c_dvt] : '',
        c_gia >= 0 ? row[c_gia] : ''
      ]);
    }
  }

  cacheSheet.clear();
  cacheSheet.getRange(1, 1, 1, cacheHeadersSP[0].length).setValues(cacheHeadersSP).setFontWeight("bold");
  if (cacheRowsSP.length > 0) {
    cacheSheet.getRange(2, 1, cacheRowsSP.length, cacheRowsSP[0].length).setValues(cacheRowsSP);
  }

  cacheSheet.getRange(1, 7).setValue("Danh Sách NV KD").setFontWeight("bold");
  if (listNvKinhDoanh.length > 0) {
    cacheSheet.getRange(2, 7, listNvKinhDoanh.length, 1).setValues(listNvKinhDoanh);
  }

  // --- 4. DATA VALIDATION CHO TAO DON ---
  const taoDonSheet = ss.getSheetByName(CONFIG.SHEET_TAO_DON);
  if (!taoDonSheet) return;

  const taoDonInfo = getDynamicSheetInfo(taoDonSheet, CONFIG.HEADER_ROW_TAO_DON, 'Sản Phẩm');
  const idxSpTaoDon = getColIndex(taoDonInfo, ['Sản Phẩm', 'Tên Sản Phẩm']);
  if (idxSpTaoDon >= 0 && cacheRowsSP.length > 0) {
    const spRange = cacheSheet.getRange(2, 1, cacheRowsSP.length, 1);
    const spRule = SpreadsheetApp.newDataValidation().requireValueInRange(spRange, true).setAllowInvalid(true).build();
    taoDonSheet.getRange(taoDonInfo.startDataRow, idxSpTaoDon + 1, CONFIG.MAX_ROWS_TAO_DON, 1).setDataValidation(spRule);
  }

  const idxNvTaoDon = getColIndex(taoDonInfo, ['Nhân Viên KD', 'Nhân Viên']);
  if (idxNvTaoDon >= 0 && listNvKinhDoanh.length > 0) {
    const nvRange = cacheSheet.getRange(2, 7, listNvKinhDoanh.length, 1);
    const nvRule = SpreadsheetApp.newDataValidation().requireValueInRange(nvRange, true).setAllowInvalid(true).build();
    taoDonSheet.getRange(taoDonInfo.startDataRow, idxNvTaoDon + 1, CONFIG.MAX_ROWS_TAO_DON, 1).setDataValidation(nvRule);
  }
}

/**
 * THỰC THI LƯU ĐƠN HÀNG DỰA TRÊN TABLE_MAPPING_SCHEMA
 * - Sinh mã: DH_yymmdd_xx, CT_yymmdd_xx
 * - Lưu MaKH (KHxxxx) vào DonHang
 * - Lưu manv và ma_sp (SPxxxx) vào ChiTietDH
 * - Dùng Link có sẵn nếu không upload file mới
 */
/**
 * THAY THẾ: Lưu đơn hàng đồng bộ theo Schema 3NF
 * - Chi tiết đơn hàng lưu masp (Mã SPxxxx)
 * - Tự động điền sl_dachay: 0, sl_dagiao: 0, sl_conlai: qty
 * - Batch Write toàn bộ mảng dữ liệu xuống Sheet
 */
function executeSaveOrdersWithSchema(payload) {
  const tTotalStart = Date.now();
  Logger.log("⏱️ ================= BẮT ĐẦU LƯU ĐƠN (TỐI ƯU TỐC ĐỘ CAO) =================");

  // MỐC 0: Khóa chống xung đột LockService
  const t0 = Date.now();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error('Hệ thống đang bận, vui lòng thử lại sau vài giây!');
  }
  Logger.log(`⏱️ [SERVER] 0. LockService: ${Date.now() - t0} ms`);

  try {
    const poGroups = payload.poGroups || [];
    if (!poGroups.length) throw new Error('Không có dữ liệu đơn hàng để lưu!');

    const now = new Date();

    // MỐC 1: Mở 2 Spreadsheet qua ID
    const t1 = Date.now();
    const targetOpId = CONFIG.ID_OPERATION;
    const targetMasterId = CONFIG.ID_MASTER;
    
    const activeSS = SpreadsheetApp.getActiveSpreadsheet();
    const ssOp = (activeSS && activeSS.getId() === targetOpId)
      ? activeSS
      : SpreadsheetApp.openById(targetOpId);
      
    const ssMaster = SpreadsheetApp.openById(targetMasterId);

    const shDH = ssOp.getSheetByName(CONFIG.SHEET_DON_HANG);
    const shCT = ssOp.getSheetByName(CONFIG.SHEET_CHI_TIET_DH);
    const shSP = ssMaster.getSheetByName(CONFIG.SHEET_SAN_PHAM);
    const shKH = ssMaster.getSheetByName(CONFIG.SHEET_KHACH_HANG);
    if (!shDH || !shCT) throw new Error('Không tìm thấy bảng 01.DonHang hoặc 02.ChiTietDH!');

    // MỐC 1: MỞ FILE DATABASE VỚI ĐỊNH DANH ID CHUẨN XÁC
    
    Logger.log(`⏱️ [SERVER] 1. Mở 2 Spreadsheet qua ID: ${Date.now() - t1} ms (ssOp: ${ssOp.getName()} - ${ssOp.getId()})`);
  
    

    /* MỐC 1: TỐI ƯU CỰC HẠN - TÁI SỬ DỤNG PHIÊN ACTIVE, CHỈ MỞ MASTER QUA ID
    const t1 = Date.now();
    const activeSS = SpreadsheetApp.getActiveSpreadsheet();
    const ssOp = (activeSS && activeSS.getId() === CONFIG.SHEET_DON_HANG)
      ? activeSS
      : SpreadsheetApp.openById(CONFIG.SHEET_DON_HANG);
    const ssMaster = SpreadsheetApp.openById(CONFIG.ID_MASTER);
    Logger.log(`⏱️ [SERVER] 1. Mở Spreadsheet qua ID: ${Date.now() - t1} ms`);
*/

   // MỐC 2: TỐI ƯU ONE-PASS TOÀN DIỆN - ĐỌC RAM 4 LẦN DUY NHẤT, QUÉT ID TRONG BỘ NHỚ
    const t2 = Date.now();

    // Chuẩn bị chuỗi ngày để quét regex ID
    const yy = String(now.getFullYear()).slice(2);
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');
    const dateStr = `${yy}${mm}${dd}`;

    // Helper tạo cấu trúc sheetInfo chuẩn hóa 1-based colMap và tự động cắt biên lastCol thực tế trong RAM
    const createSheetInfoFromHeaders = (rawHeaders, headerRow, lastRow) => {
      const colMap = {};
      let trueLastCol = 0;
      if (Array.isArray(rawHeaders)) {
        for (let c = 0; c < rawHeaders.length; c++) {
          const val = rawHeaders[c];
          if (val !== undefined && val !== null && String(val).trim() !== '') {
            colMap[normalizeKey(String(val))] = c + 1; // 1-based index đồng bộ chuẩn getColIndex()
            trueLastCol = c + 1;
          }
        }
      }
      return {
        headerRow: headerRow,
        startDataRow: headerRow + 1,
        lastRow: lastRow,
        lastCol: trueLastCol,
        headers: Array.isArray(rawHeaders) ? rawHeaders.slice(0, trueLastCol) : [],
        colMap: colMap
      };
    };

    // 1. Bảng 01.DonHang: Đọc RAM 1 lần duy nhất (Lấy tiêu đề động & Quét ID trong RAM)
    const dhRawData = shDH.getDataRange().getValues();
    const dhHeaderIdx = TABLE_MAPPING_SCHEMA.DON_HANG.headerRow - 1;
    const headersDH = dhRawData.length > dhHeaderIdx ? dhRawData[dhHeaderIdx] : [];
    const infoDH = createSheetInfoFromHeaders(headersDH, TABLE_MAPPING_SCHEMA.DON_HANG.headerRow, dhRawData.length);
    const cDH_MaDH = getColIndex(infoDH, getTargetHeaders('DON_HANG', 'maDH'));

    let curSeqDH = 0;
    if (cDH_MaDH >= 0 && dhRawData.length >= infoDH.startDataRow) {
      const startIdxDH = Math.max(infoDH.startDataRow - 1, dhRawData.length - 100);
      const regexDH = new RegExp(`^DH_${dateStr}_(\\d+)$`);
      for (let i = startIdxDH; i < dhRawData.length; i++) {
        const m = String(dhRawData[i][cDH_MaDH] || '').trim().match(regexDH);
        if (m) {
          const s = parseInt(m[1], 10);
          if (s > curSeqDH) curSeqDH = s;
        }
      }
    }

    // 2. Bảng 02.ChiTietDH: Đọc RAM 1 lần duy nhất (Lấy tiêu đề động & Quét ID trong RAM)
    const ctRawData = shCT.getDataRange().getValues();
    const ctHeaderIdx = TABLE_MAPPING_SCHEMA.CHI_TIET_DH.headerRow - 1;
    const headersCT = ctRawData.length > ctHeaderIdx ? ctRawData[ctHeaderIdx] : [];
    const infoCT = createSheetInfoFromHeaders(headersCT, TABLE_MAPPING_SCHEMA.CHI_TIET_DH.headerRow, ctRawData.length);
    const cCT_MaCT = getColIndex(infoCT, getTargetHeaders('CHI_TIET_DH', 'maCT'));

    let curSeqCT = 0;
    if (cCT_MaCT >= 0 && ctRawData.length >= infoCT.startDataRow) {
      const startIdxCT = Math.max(infoCT.startDataRow - 1, ctRawData.length - 100);
      const regexCT = new RegExp(`^CT_${dateStr}_(\\d+)$`);
      for (let i = startIdxCT; i < ctRawData.length; i++) {
        const m = String(ctRawData[i][cCT_MaCT] || '').trim().match(regexCT);
        if (m) {
          const s = parseInt(m[1], 10);
          if (s > curSeqCT) curSeqCT = s;
        }
      }
    }

    // 3. Bảng 02_SAN_PHAM: Đọc RAM 1 lần duy nhất (Nạp danh mục & Max SPxxxx)
    const existingProducts = {};
    const productMapByName = {};
    let curSeqSP = 0;
    let infoSP = null;

    if (shSP) {
      const spRawData = shSP.getDataRange().getValues();
      const spHeaderIdx = TABLE_MAPPING_SCHEMA.SAN_PHAM_MASTER.headerRow - 1;
      const headersSP = spRawData.length > spHeaderIdx ? spRawData[spHeaderIdx] : [];
      infoSP = createSheetInfoFromHeaders(headersSP, TABLE_MAPPING_SCHEMA.SAN_PHAM_MASTER.headerRow, spRawData.length);

      const cSP_MaSP = getColIndex(infoSP, getTargetHeaders('SAN_PHAM_MASTER', 'spId'));
      const cSP_Name = getColIndex(infoSP, getTargetHeaders('SAN_PHAM_MASTER', 'nameNapaco'));

      for (let i = infoSP.startDataRow - 1; i < spRawData.length; i++) {
        const row = spRawData[i];
        if (cSP_MaSP >= 0) {
          const m = String(row[cSP_MaSP] || '').match(/^SP(\d+)$/);
          if (m) {
            const num = parseInt(m[1], 10);
            if (num > curSeqSP) curSeqSP = num;
          }
        }
        if (cSP_Name >= 0) {
          const pName = normalizeKey(row[cSP_Name]);
          const spIdVal = cSP_MaSP >= 0 ? String(row[cSP_MaSP] || '').trim() : '';
          if (pName) {
            existingProducts[pName] = true;
            if (spIdVal) productMapByName[pName] = spIdVal;
          }
        }
      }
    }

    // 4. Bảng 01_KHACH_HANG: Đọc RAM 1 lần duy nhất (Nạp danh mục & Max KHxxxx)
    const existingCustomers = {};
    const customerMapByName = {};
    let curSeqKH = 0;
    let infoKH = null;

    if (shKH) {
      const khRawData = shKH.getDataRange().getValues();
      const khHeaderIdx = TABLE_MAPPING_SCHEMA.KHACH_HANG_MASTER.headerRow - 1;
      const headersKH = khRawData.length > khHeaderIdx ? khRawData[khHeaderIdx] : [];
      infoKH = createSheetInfoFromHeaders(headersKH, TABLE_MAPPING_SCHEMA.KHACH_HANG_MASTER.headerRow, khRawData.length);

      const cKH_MaKH = getColIndex(infoKH, getTargetHeaders('KHACH_HANG_MASTER', 'Ma_KH'));
      const cKH_Name = getColIndex(infoKH, getTargetHeaders('KHACH_HANG_MASTER', 'name'));

      for (let k = infoKH.startDataRow - 1; k < khRawData.length; k++) {
        const row = khRawData[k];
        if (cKH_MaKH >= 0) {
          const m = String(row[cKH_MaKH] || '').match(/^KH(\d+)$/);
          if (m) {
            const num = parseInt(m[1], 10);
            if (num > curSeqKH) curSeqKH = num;
          }
        }
        if (cKH_Name >= 0) {
          const ten = normalizeKey(row[cKH_Name]);
          const ma = cKH_MaKH >= 0 ? String(row[cKH_MaKH] || '').trim() : '';
          if (ten) {
            existingCustomers[ten] = true;
            if (ma) customerMapByName[ten] = ma;
          }
        }
      }
    }
    Logger.log(`⏱️ [SERVER] 2. Quét Header & nạp dữ liệu vào RAM (One-Pass 4 Bảng): ${Date.now() - t2} ms`);

    // // code hiện tại: MỐC 3: Xử lý logic nghiệp vụ trong RAM (Giữ nguyên toàn bộ phía sau)

    // MỐC 3: Xử lý logic nghiệp vụ trong RAM
    const t3 = Date.now();
    let storageFolder = null;
    try {
      storageFolder = DriveApp.getFolderById(CONFIG.ID_STORAGES);
    } catch (e) {
      Logger.log("Lỗi mở thư mục Storages: " + e.message);
    }

    const orderRows = [];
    const detailRows = [];
    const newProductRows = [];
    const newCustomerRows = [];
    const newCacheRows = []; // Dữ liệu bổ sung trực tiếp vào Cache siêu tốc
    const summaryPOList = [];
    let totalItems = 0;
    let timeUploadDrive = 0;

    for (let g = 0; g < poGroups.length; g++) {
      const grp = poGroups[g];
      const soPO = String(grp.soPO || '').trim();
      summaryPOList.push(soPO);

      curSeqDH++;
      const maDH = `DH_${dateStr}_${String(curSeqDH).padStart(2, '0')}`;

      // Upload file nếu có
      let fileUrl = grp.existingFileUrl || '';
      if (grp.fileAttachment && grp.fileAttachment.base64 && storageFolder) {
        const tUp = Date.now();
        try {
          const fileBytes = Utilities.base64Decode(grp.fileAttachment.base64);
          const ext = grp.fileAttachment.extension || 'pdf';
          const blob = Utilities.newBlob(fileBytes, grp.fileAttachment.mimeType || 'application/octet-stream', `${soPO}.${ext}`);
          const savedFile = storageFolder.createFile(blob);
          fileUrl = savedFile.getUrl();
        } catch (errUpload) {
          Logger.log(`Lỗi upload file: ${errUpload.message}`);
        }
        timeUploadDrive += (Date.now() - tUp);
      }

      // Xử lý Khách Hàng mới
      const custNorm = normalizeKey(grp.customerName);
      let currentMaKH = customerMapByName[custNorm] || '';
      if (custNorm && !existingCustomers[custNorm] && shKH && infoKH) {
        curSeqKH++;
        currentMaKH = `KH${String(curSeqKH).padStart(4, '0')}`;
        const khObj = {
          Ma_KH: currentMaKH,
          name: grp.customerName
        };
        newCustomerRows.push(buildRowBySchema(TABLE_MAPPING_SCHEMA.KHACH_HANG_MASTER, infoKH, khObj));
        existingCustomers[custNorm] = true;
        customerMapByName[custNorm] = currentMaKH;
      }

      const salesRaw = String(grp.salesPerson || '').trim();
      const maNV = salesRaw.includes(' - ') ? salesRaw.split(' - ')[0].trim() : salesRaw;
      const items = grp.items || [];

      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        totalItems++;
        curSeqCT++;
        const maCT = `CT_${dateStr}_${String(curSeqCT).padStart(2, '0')}`;
        const qty = Number(item.qty) || 0;
        const price = Number(item.price) || 0;

        // Xử lý Sản Phẩm mới
        const pNorm = normalizeKey(item.prodName);
        let currentMaSP = productMapByName[pNorm] || '';
        if (pNorm && !existingProducts[pNorm] && shSP && infoSP) {
          curSeqSP++;
          currentMaSP = `SP${String(curSeqSP).padStart(4, '0')}`;
          const spObj = {
            spId: currentMaSP,
            customerId: currentMaKH,
            mang: item.mang || '',
            codeSanpham: item.codeSP || '',
            nameNapaco: item.prodName || '',
            donvi: item.unit || '',
            dongia: price
          };
          newProductRows.push(buildRowBySchema(TABLE_MAPPING_SCHEMA.SAN_PHAM_MASTER, infoSP, spObj));
          
          // Tạo bản ghi nạp nhanh vào Cache: [Tên SP, Khách Hàng, Code SP, Đơn Vị, Đơn Giá]
          newCacheRows.push([
            item.prodName || '',
            grp.customerName || '',
            item.codeSP || '',
            item.unit || '',
            price || ''
          ]);

          existingProducts[pNorm] = true;
          productMapByName[pNorm] = currentMaSP;
        }

        const detailObj = {
          maCT: maCT,
          maDH: maDH,
          masp: currentMaSP,
          qty: qty,
          price: price,
          sl_dachay: 0,
          sl_dagiao: 0,
          sl_conlai: qty,
          note: item.note || grp.orderNote || '',
          deliveryDate: item.deliveryDate || grp.ngayGiaoDK || '',
          manv: maNV,
          status: CONFIG.DEFAULT_STATUS_DETAIL
        };
        detailRows.push(buildRowBySchema(TABLE_MAPPING_SCHEMA.CHI_TIET_DH, infoCT, detailObj));
      }

      // Chuẩn hóa ngày tháng: Khóa mốc 12:00:00 trưa triệt tiêu độ lệch múi giờ GMT+7
      const parseDateSafe = (dVal) => {
        if (!dVal) return '';
        if (dVal instanceof Date && !isNaN(dVal.getTime())) {
          return new Date(dVal.getFullYear(), dVal.getMonth(), dVal.getDate(), 12, 0, 0);
        }
        const s = String(dVal).trim();
        if (!s) return '';
        
        // Nhận dạng chuẩn ISO từ input type="date" (YYYY-MM-DD hoặc YYYY/MM/DD)
        const mISO = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
        if (mISO) {
          return new Date(parseInt(mISO[1], 10), parseInt(mISO[2], 10) - 1, parseInt(mISO[3], 10), 12, 0, 0);
        }
        
        // Nhận dạng chuẩn Việt Nam (DD/MM/YYYY hoặc DD-MM-YYYY)
        const mVN = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
        if (mVN) {
          return new Date(parseInt(mVN[3], 10), parseInt(mVN[2], 10) - 1, parseInt(mVN[1], 10), 12, 0, 0);
        }
        
        const fallback = new Date(s);
        return isNaN(fallback.getTime()) ? '' : new Date(fallback.getFullYear(), fallback.getMonth(), fallback.getDate(), 12, 0, 0);
      };

      const finalNgayDat = parseDateSafe(grp.ngayDat || grp.orderDate) || new Date(now.getFullYear(), now.getMonth(), now.getDate(), 12, 0, 0);
      const finalNgayGiao = parseDateSafe(grp.ngayGiaoDK || grp.deliveryDate || grp.ngayGiao) || '';

      const orderObj = {
        maDH: maDH,
        maKH: currentMaKH,
        soPO: soPO,
        ngayDat: finalNgayDat,
        ngayGiaoDK: finalNgayGiao,
        fileUrl: fileUrl
      };
      orderRows.push(buildRowBySchema(TABLE_MAPPING_SCHEMA.DON_HANG, infoDH, orderObj));
    }
    Logger.log(`⏱️ [SERVER] 3. Tính toán RAM + Sinh ID: ${Date.now() - t3} ms (Upload Drive: ${timeUploadDrive} ms)`);

    // MỐC 4: Batch Writing 4 bảng xuống Sheet (setValues 1 lần duy nhất mỗi bảng)
    const t4 = Date.now();
    if (newCustomerRows.length > 0 && shKH) {
      shKH.getRange(shKH.getLastRow() + 1, 1, newCustomerRows.length, infoKH.lastCol).setValues(newCustomerRows);
    }
    if (newProductRows.length > 0 && shSP) {
      shSP.getRange(shSP.getLastRow() + 1, 1, newProductRows.length, infoSP.lastCol).setValues(newProductRows);
    }
    if (orderRows.length > 0) {
      shDH.getRange(shDH.getLastRow() + 1, 1, orderRows.length, infoDH.lastCol).setValues(orderRows);
    }
    if (detailRows.length > 0) {
      shCT.getRange(shCT.getLastRow() + 1, 1, detailRows.length, infoCT.lastCol).setValues(detailRows);
    }
    Logger.log(`⏱️ [SERVER] 4. Batch Write 4 bảng setValues(): ${Date.now() - t4} ms`);

    /* MỐC 5: XÓA TRẮNG DỮ LIỆU FORM NHẬP TRÊN GIAO DIỆN ĐANG MỞ (DIRECT FLUSH)
    const t5 = Date.now();
    try {
      const sheetTaoDon = ssOp.getSheetByName(CONFIG.SHEET_TAO_DON);
      if (sheetTaoDon) {
        const startR = (CONFIG.HEADER_ROW_TAO_DON || 7) + 1; // Mặc định từ dòng 8
        const curLastRow = sheetTaoDon.getLastRow();
        if (curLastRow >= startR) {
          const numRows = curLastRow - startR + 1;
          const numCols = Math.max(sheetTaoDon.getLastColumn(), 26);
          // Xóa trực tiếp nội dung dữ liệu, giữ nguyên định dạng viền, dropdown và tiêu đề
          sheetTaoDon.getRange(startR, 1, numRows, numCols).clearContent();
        }
      }
    } catch (errClear) {
      Logger.log("⚠️ Cảnh báo dọn form TAO DON: " + errClear.message);
    }
    Logger.log(`⏱️ [SERVER] 5. Dọn sheet form TAO DON (Fast-Clear): ${Date.now() - t5} ms`);
    */

    // MỐC 5: XÓA TRẮNG FORM TAO DON (BUFFER CỐ ĐỊNH 50 DÒNG, TỐN ~1-2 MS)
    const t5 = Date.now();
    try {
      const sheetTaoDonName = CONFIG.SHEET_TAO_DON || 'TAO DON';
      const sheetTaoDon = ssOp.getSheetByName(sheetTaoDonName);
      if (sheetTaoDon) {
        const startR = (CONFIG.HEADER_ROW_TAO_DON || 7) + 1; // Mặc định từ dòng 8
        const maxBufferRows = 50; // Quét sạch tối đa 50 dòng theo quy chuẩn
        const maxCols = Math.max(sheetTaoDon.getLastColumn(), 26);
        
        // Xóa trực tiếp nội dung dữ liệu, giữ nguyên viền, dropdown và định dạng
        sheetTaoDon.getRange(startR, 1, maxBufferRows, maxCols).clearContent();
        Logger.log(`🧹 [SERVER] Đã dọn sạch 50 dòng form trên sheet "${sheetTaoDonName}".`);
      } else {
        Logger.log(`⚠️ [SERVER] Không tìm thấy sheet có tên: "${sheetTaoDonName}" trong file Operation.`);
      }
    } catch (errClear) {
      Logger.log("⚠️ Cảnh báo dọn form TAO DON: " + errClear.message);
    }
    Logger.log(`⏱️ [SERVER] 5. Dọn sheet form TAO DON: ${Date.now() - t5} ms`);

    // MỐC 6: ĐỒNG BỘ NẠP NỐI ĐUÔI VÀO SHEET CACHE (DYNAMIC COLUMNS)
    const t6 = Date.now();
    try {
      if (newCacheRows && Array.isArray(newCacheRows) && newCacheRows.length > 0) {
        const sheetCacheName = CONFIG.SHEET_CACHE || '_CACHE_DANH_MUC';
        const sheetCache = ssOp.getSheetByName(sheetCacheName);
        
        if (sheetCache) {
          // 1. Quét Header dòng 1 để thiết lập ánh xạ cột động (Dynamic Columns)
          const lastColCache = Math.max(sheetCache.getLastColumn(), 6);
          const headerValues = sheetCache.getRange(1, 1, 1, lastColCache).getValues()[0];
          const cacheColMap = {};
          
          headerValues.forEach((h, idx) => {
            if (h) {
              const k = normalizeKey(String(h));
              cacheColMap[k] = idx + 1; // 1-based index
            }
          });

          // Helper tìm vị trí cột động theo nhiều tên alias khả dĩ
          const getColIdx = (aliases) => {
            for (let name of aliases) {
              const k = normalizeKey(name);
              if (cacheColMap[k]) return cacheColMap[k] - 1; // 0-based cho mảng
            }
            return -1;
          };

          const colTenSP = getColIdx(['Tên Sản Phẩm', 'Ten_San_Pham', 'Sản Phẩm']);
          const colKhachHang = getColIdx(['Khách Hàng', 'Khach_Hang', 'Tên Khách']);
          const colCodeSP = getColIdx(['Code Sản Phẩm', 'Code_San_Pham', 'Mã SP Khách']);
          const colDonVi = getColIdx(['Đơn Vị', 'Don_Vi', 'DVT']);
          const colDonGia = getColIdx(['Đơn Giá', 'Don_Gia', 'Giá Bán']);

          // 2. Chuyển đổi newCacheRows sang ma trận dữ liệu khớp chuẩn theo vị trí cột động
          const appendMatrix = [];
          for (let r = 0; r < newCacheRows.length; r++) {
            const rawItem = newCacheRows[r]; // [prodName, customerName, codeSP, unit, price]
            const rowArr = new Array(lastColCache).fill('');
            
            if (colTenSP >= 0) rowArr[colTenSP] = rawItem[0] || '';
            if (colKhachHang >= 0) rowArr[colKhachHang] = rawItem[1] || '';
            if (colCodeSP >= 0) rowArr[colCodeSP] = rawItem[2] || '';
            if (colDonVi >= 0) rowArr[colDonVi] = rawItem[3] || '';
            if (colDonGia >= 0) rowArr[colDonGia] = rawItem[4] || '';

            appendMatrix.push(rowArr);
          }

          // 3. Batch Write ghi nối đuôi vào dòng cuối của sheet Cache
          const startWriteRow = sheetCache.getLastRow() + 1;
          sheetCache.getRange(startWriteRow, 1, appendMatrix.length, lastColCache).setValues(appendMatrix);
          Logger.log(`⚡ [SERVER] Đã đồng bộ nối đuôi ${appendMatrix.length} sản phẩm mới vào "${sheetCacheName}".`);
        } else {
          Logger.log(`⚠️ [SERVER] Không tìm thấy sheet Cache có tên: "${sheetCacheName}" trong file Operation.`);
        }
      } else {
        Logger.log("ℹ️ [SERVER] Không có sản phẩm mới phát sinh để nạp vào Cache.");
      }
    } catch (errCache) {
      Logger.log("⚠️ Cảnh báo nạp Cache: " + errCache.message);
    }
    Logger.log(`⏱️ [SERVER] 6. Cập nhật Cache siêu tốc (Fast-Append): ${Date.now() - t6} ms`);


    

    const totalServer = Date.now() - tTotalStart;
    Logger.log(`🏁 [SERVER] === TỔNG THỜI GIAN SERVER XỬ LÝ: ${(totalServer / 1000).toFixed(2)} GIÂY (${totalServer} ms) ===`);

    return {
      success: true,
      orderCount: poGroups.length,
      itemCount: totalItems,
      newProductCount: newProductRows.length,
      summaryText: summaryPOList.join(', ')
    };
  } finally {
    // Luôn giải phóng khóa an toàn
    if (lock) {
      try { lock.releaseLock(); } catch (e) {}
    }
  }
}