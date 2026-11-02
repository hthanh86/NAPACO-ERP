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
  const c_giakhuon = getColIndex(spInfo, ['Gia Khuon', 'Giá Khuôn', 'gia_khuon']);
  const c_mang = getColIndex(spInfo, ['mang', 'Màng']);
  const c_type = getColIndex(spInfo, ['type', 'Type', 'Loại']);
  const c_donggoi = getColIndex(spInfo, ['donggoi', 'Đóng Gói', 'dong_goi']);

  const cacheHeadersSP = [['Tên Sản Phẩm', 'Khách Hàng', 'Code Sản Phẩm', 'Đơn Vị', 'Đơn Giá', 'Giá Khuôn', 'Màng', 'Type', 'Đóng Gói']];
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
        c_gia >= 0 ? row[c_gia] : '',
        c_giakhuon >= 0 ? row[c_giakhuon] : '',
        c_mang >= 0 ? row[c_mang] : '',
        c_type >= 0 ? row[c_type] : '',
        c_donggoi >= 0 ? row[c_donggoi] : ''
      ]);
    }
  }

  cacheSheet.clear();
  cacheSheet.getRange(1, 1, 1, cacheHeadersSP[0].length).setValues(cacheHeadersSP).setFontWeight("bold");
  if (cacheRowsSP.length > 0) {
    cacheSheet.getRange(2, 1, cacheRowsSP.length, cacheRowsSP[0].length).setValues(cacheRowsSP);
  }

  cacheSheet.getRange(1, 11).setValue("Danh Sách NV KD").setFontWeight("bold");
  if (listNvKinhDoanh.length > 0) {
    cacheSheet.getRange(2, 11, listNvKinhDoanh.length, 1).setValues(listNvKinhDoanh);
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
    const nvRange = cacheSheet.getRange(2, 11, listNvKinhDoanh.length, 1);
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
 * THỰC THI LƯU ĐƠN HÀNG DỰA TRÊN TABLE_MAPPING_SCHEMA
 * - Rút tên cột động từ Config.gs qua getTargetHeaders()
 * - Nhận diện sản phẩm cũ bằng khóa (name_napaco + customer_id)
 * - Tự động cập nhật Code sản phẩm, Đơn vị, Đơn giá cho sản phẩm cũ (chỉ ghi đè khi khác rỗng và khác 0)
 * - Tự động thêm sản phẩm mới nếu chưa có
 * - Xóa sạch biểu mẫu nhập liệu trên sheet TAO DON sau khi lưu thành công
 * - Batch Write toàn bộ mảng dữ liệu xuống Sheet (1 lần duy nhất mỗi bảng)
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

    // MỐC 1: Mở các Spreadsheet qua ID cấu hình
    const t1 = Date.now();
    const targetOpId = CONFIG.ID_OPERATION;
    const targetMasterId = CONFIG.ID_MASTER;
    
    const activeSS = SpreadsheetApp.getActiveSpreadsheet();
    const ssOp = (activeSS && activeSS.getId() === targetOpId)
      ? activeSS
      : SpreadsheetApp.openById(targetOpId);
      
    const ssMaster = (activeSS && activeSS.getId() === targetMasterId)
      ? activeSS
      : SpreadsheetApp.openById(targetMasterId);

    const shDH = ssOp.getSheetByName(CONFIG.SHEET_DON_HANG);
    const shCT = ssOp.getSheetByName(CONFIG.SHEET_CHI_TIET_DH);
    const shSP = ssMaster.getSheetByName(CONFIG.SHEET_SAN_PHAM);
    const shKH = ssMaster.getSheetByName(CONFIG.SHEET_KHACH_HANG);
    if (!shDH || !shCT) throw new Error('Không tìm thấy bảng 01.DonHang hoặc 02.ChiTietDH!');

    Logger.log(`⏱️ [SERVER] 1. Mở Spreadsheet qua ID: ${Date.now() - t1} ms`);

    // MỐC 2: TỐI ƯU ONE-PASS TOÀN DIỆN - ĐỌC RAM & LẬP BẢN ĐỒ VỊ TRÍ CỘT THEO SCHEMA
    const t2 = Date.now();

    const yy = String(now.getFullYear()).slice(2);
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');
    const dateStr = `${yy}${mm}${dd}`;

    const getTargetKeys = (tableName, key, fallbacks) => {
      let headers = [];
      if (typeof getTargetHeaders === 'function') {
        headers = getTargetHeaders(tableName, key) || [];
      }
      return Array.from(new Set([...headers, ...fallbacks]));
    };

    const infoDH = getDynamicSheetInfo(shDH, 1, 'Ma_DH');
    const infoCT = getDynamicSheetInfo(shCT, 1, 'Ma_CT');
    const infoKH = shKH ? getDynamicSheetInfo(shKH, 1, 'Ma_KH') : null;
    const infoSP = shSP ? getDynamicSheetInfo(shSP, 1, 'name_napaco') : null;

    // 1. Quét sequence DH
    let curSeqDH = 0;
    if (infoDH.lastRow >= infoDH.startDataRow) {
      const colMaDH = getColIndex(infoDH, getTargetKeys('DON_HANG', 'maDH', ['Mã Đơn Hàng', 'Ma_DH', 'Mã ĐH']));
      if (colMaDH >= 0) {
        const valsDH = shDH.getRange(infoDH.startDataRow, colMaDH + 1, infoDH.lastRow - infoDH.startDataRow + 1, 1).getValues();
        const regDH = new RegExp(`^DH_${dateStr}_(\\d+)$`);
        for (let r = 0; r < valsDH.length; r++) {
          const m = String(valsDH[r][0]).match(regDH);
          if (m) {
            const num = parseInt(m[1], 10);
            if (num > curSeqDH) curSeqDH = num;
          }
        }
      }
    }

    // 2. Quét sequence CT
    let curSeqCT = 0;
    if (infoCT.lastRow >= infoCT.startDataRow) {
      const colMaCT = getColIndex(infoCT, getTargetKeys('CHI_TIET_DH', 'maCT', ['Mã Chi Tiết', 'ID_Chi_Tiet', 'Mã CT']));
      if (colMaCT >= 0) {
        const valsCT = shCT.getRange(infoCT.startDataRow, colMaCT + 1, infoCT.lastRow - infoCT.startDataRow + 1, 1).getValues();
        const regCT = new RegExp(`^CT_${dateStr}_(\\d+)$`);
        for (let r = 0; r < valsCT.length; r++) {
          const m = String(valsCT[r][0]).match(regCT);
          if (m) {
            const num = parseInt(m[1], 10);
            if (num > curSeqCT) curSeqCT = num;
          }
        }
      }
    }

    // 3. Đọc & map Khách Hàng từ DB Master
    let curSeqKH = 0;
    const customerMapByName = {};
    const existingCustomers = {};
    if (shKH && infoKH && infoKH.lastRow >= infoKH.startDataRow) {
      const colMaKH = getColIndex(infoKH, getTargetKeys('KHACH_HANG_MASTER', 'Ma_KH', ['Ma_KH', 'customer_id']));
      const colTenKH = getColIndex(infoKH, getTargetKeys('KHACH_HANG_MASTER', 'name', ['name', 'Tên Khách Hàng', 'Khách Hàng']));
      const khData = shKH.getRange(infoKH.startDataRow, 1, infoKH.lastRow - infoKH.startDataRow + 1, infoKH.lastCol).getValues();
      for (let r = 0; r < khData.length; r++) {
        const makh = colMaKH >= 0 ? String(khData[r][colMaKH]).trim() : '';
        const tenkh = colTenKH >= 0 ? String(khData[r][colTenKH]).trim() : '';
        if (makh) {
          const m = makh.match(/^KH(\d+)$/i);
          if (m) {
            const num = parseInt(m[1], 10);
            if (num > curSeqKH) curSeqKH = num;
          }
        }
        if (tenkh) {
          const nKey = normalizeKey(tenkh);
          customerMapByName[nKey] = makh;
          existingCustomers[nKey] = true;
        }
      }
    }

    // 4. Đọc & map Sản Phẩm từ DB Master
    let curSeqSP = 0;
    let spExistingData = [];
    const productMapByCompositeKey = {};
    const productMapByNameOnly = {};

    const colIdxSP_ID = infoSP ? getColIndex(infoSP, getTargetKeys('SAN_PHAM_MASTER', 'spId', ['SP_ID', 'sp_id', 'Mã Sản Phẩm', 'Mã SP'])) : -1;
    const colIdxSP_Cust = infoSP ? getColIndex(infoSP, getTargetKeys('SAN_PHAM_MASTER', 'customerId', ['customer_id', 'Khách Hàng', 'Ma_KH'])) : -1;
    const colIdxSP_Name = infoSP ? getColIndex(infoSP, getTargetKeys('SAN_PHAM_MASTER', 'nameNapaco', ['name_napaco', 'Tên Sản Phẩm', 'Sản Phẩm'])) : -1;
    const colIdxSP_Code = infoSP ? getColIndex(infoSP, getTargetKeys('SAN_PHAM_MASTER', 'codeSanpham', ['Code_Sanpham', 'Code Sản Phẩm', 'code_sp'])) : -1;
    const colIdxSP_DonVi = infoSP ? getColIndex(infoSP, getTargetKeys('SAN_PHAM_MASTER', 'donvi', ['donvi', 'Đơn Vị', 'DVT'])) : -1;
    const colIdxSP_DonGia = infoSP ? getColIndex(infoSP, getTargetKeys('SAN_PHAM_MASTER', 'dongia', ['dongia', 'Đơn Giá', 'Gia_Ban'])) : -1;
    const colIdxSP_GiaKhuon = infoSP ? getColIndex(infoSP, getTargetKeys('SAN_PHAM_MASTER', 'giakhuon', ['Gia Khuon', 'Giá Khuôn', 'gia_khuon'])) : -1;
    const colIdxSP_Mang = infoSP ? getColIndex(infoSP, getTargetKeys('SAN_PHAM_MASTER', 'mang', ['mang', 'Màng'])) : -1;
    const colIdxSP_Type = infoSP ? getColIndex(infoSP, getTargetKeys('SAN_PHAM_MASTER', 'type', ['type', 'Type', 'Loại'])) : -1;
    const colIdxSP_DongGoi = infoSP ? getColIndex(infoSP, getTargetKeys('SAN_PHAM_MASTER', 'donggoi', ['donggoi', 'Đóng Gói', 'dong_goi'])) : -1;

    if (shSP && infoSP && infoSP.lastRow >= infoSP.startDataRow) {
      spExistingData = shSP.getRange(infoSP.startDataRow, 1, infoSP.lastRow - infoSP.startDataRow + 1, infoSP.lastCol).getValues();
      for (let r = 0; r < spExistingData.length; r++) {
        const row = spExistingData[r];
        const masp = colIdxSP_ID >= 0 ? String(row[colIdxSP_ID]).trim() : '';
        const tensp = colIdxSP_Name >= 0 ? String(row[colIdxSP_Name]).trim() : '';
        const custId = colIdxSP_Cust >= 0 ? String(row[colIdxSP_Cust]).trim() : '';

        if (masp) {
          const m = masp.match(/^SP(\d+)$/i);
          if (m) {
            const num = parseInt(m[1], 10);
            if (num > curSeqSP) curSeqSP = num;
          }
        }

        if (tensp) {
          const pKey = normalizeKey(tensp);
          const cKey = normalizeKey(custId);

          if (cKey) {
            productMapByCompositeKey[pKey + '___' + cKey] = { spId: masp, rIdx: r };
          }
          if (!productMapByNameOnly[pKey]) {
            productMapByNameOnly[pKey] = { spId: masp, rIdx: r };
          }
        }
      }
    }

    Logger.log(`⏱️ [SERVER] 2. Quét RAM & lập Index theo Schema: ${Date.now() - t2} ms`);

    // MỐC 3: TÍNH TOÁN RAM + XỬ LÝ UPSERT SẢN PHẨM & ĐƠN HÀNG
    const t3 = Date.now();
    let timeUploadDrive = 0;

    const newCustomerRows = [];
    const newProductRows = [];
    const newCacheRows = [];
    const orderRows = [];
    const detailRows = [];
    let hasUpdatedExistingProducts = false;

    for (let g = 0; g < poGroups.length; g++) {
      const grp = poGroups[g];
      curSeqDH++;
      const maDH = `DH_${dateStr}_${String(curSeqDH).padStart(2, '0')}`;
      const soPO = grp.soPO || '';

      // Upload file PO nếu có
      let fileUrl = grp.existingFileUrl || '';
      if (grp.fileBase64 && grp.fileName) {
        const tUp = Date.now();
        fileUrl = uploadPOFileToDrive(grp.fileBase64, grp.fileName, grp.customerName || 'KH', soPO, maDH);
        timeUploadDrive += (Date.now() - tUp);
      }

      // Xử lý Khách Hàng
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
        curSeqCT++;
        const maCT = `CT_${dateStr}_${String(curSeqCT).padStart(2, '0')}`;
        const qty = Number(item.qty) || 0;
        const price = Number(item.price !== undefined ? item.price : item.dongia) || 0;
        const giaKhuon = Number(item.giaKhuon !== undefined ? item.giaKhuon : item.giakhuon) || 0;

        const pNorm = normalizeKey(item.prodName);
        const keyWithMaKH = pNorm + '___' + normalizeKey(currentMaKH);
        const keyWithCustName = pNorm + '___' + custNorm;
        let currentMaSP = '';

        if (pNorm && shSP && infoSP) {
          // Đối soát sản phẩm đa tầng
          const matchEntry = productMapByCompositeKey[keyWithMaKH] || 
                             productMapByCompositeKey[keyWithCustName] || 
                             productMapByNameOnly[pNorm];

          if (matchEntry) {
            // --- SẢN PHẨM CŨ: CẬP NHẬT THÔNG TIN NẾU KHÁC RỖNG VÀ KHÁC 0 ---
            currentMaSP = matchEntry.spId;
            const targetRow = spExistingData[matchEntry.rIdx];

            if (targetRow) {
              let rowModified = false;

              // 1. Code Sản Phẩm
              const newCode = String(item.codeSP !== undefined ? item.codeSP : (item.codeSanpham !== undefined ? item.codeSanpham : '')).trim();
              if (newCode !== '' && colIdxSP_Code >= 0) {
                if (String(targetRow[colIdxSP_Code]).trim() !== newCode) {
                  targetRow[colIdxSP_Code] = newCode;
                  rowModified = true;
                }
              }

              // 2. Đơn Vị Tính
              const newUnit = String(item.unit !== undefined ? item.unit : (item.donvi !== undefined ? item.donvi : '')).trim();
              if (newUnit !== '' && colIdxSP_DonVi >= 0) {
                if (String(targetRow[colIdxSP_DonVi]).trim() !== newUnit) {
                  targetRow[colIdxSP_DonVi] = newUnit;
                  rowModified = true;
                }
              }

              // 3. Đơn Giá (chỉ cập nhật khi khác rỗng và > 0)
              const rawPrice = item.price !== undefined ? item.price : item.dongia;
              const numPrice = Number(rawPrice);
              if (rawPrice !== '' && rawPrice !== null && !isNaN(numPrice) && numPrice > 0 && colIdxSP_DonGia >= 0) {
                if (Number(targetRow[colIdxSP_DonGia]) !== numPrice) {
                  targetRow[colIdxSP_DonGia] = numPrice;
                  rowModified = true;
                }
              }

              // 4. Giá Khuôn (chỉ cập nhật khi khác rỗng và > 0)
        if (giaKhuon > 0 && colIdxSP_GiaKhuon >= 0) {
          if (Number(targetRow[colIdxSP_GiaKhuon]) !== giaKhuon) {
            targetRow[colIdxSP_GiaKhuon] = giaKhuon;
            rowModified = true;
          }
        }

        // 5. Màng (cập nhật nếu có dữ liệu mới)
        const newMang = String(item.mang !== undefined ? item.mang : '').trim();
        if (newMang !== '' && colIdxSP_Mang >= 0) {
          if (String(targetRow[colIdxSP_Mang]).trim() !== newMang) {
            targetRow[colIdxSP_Mang] = newMang;
            rowModified = true;
          }
        }

        // 6. Type (cập nhật nếu có dữ liệu mới)
        const newType = String(item.type !== undefined ? item.type : '').trim();
        if (newType !== '' && colIdxSP_Type >= 0) {
          if (String(targetRow[colIdxSP_Type]).trim() !== newType) {
            targetRow[colIdxSP_Type] = newType;
            rowModified = true;
          }
        }

        // 7. Đóng Gói (cập nhật nếu có dữ liệu mới)
        const newDongGoi = String(item.donggoi !== undefined ? item.donggoi : '').trim();
        if (newDongGoi !== '' && colIdxSP_DongGoi >= 0) {
          if (String(targetRow[colIdxSP_DongGoi]).trim() !== newDongGoi) {
            targetRow[colIdxSP_DongGoi] = newDongGoi;
            rowModified = true;
          }
        }

        // 8. Bổ sung customer_id vào sản phẩm nếu trước đó chưa gán
        if (currentMaKH && colIdxSP_Cust >= 0 && !String(targetRow[colIdxSP_Cust]).trim()) {
          targetRow[colIdxSP_Cust] = currentMaKH;
          rowModified = true;
        }

        if (rowModified) {
          hasUpdatedExistingProducts = true;
        }
      }
    } else {
      // --- SẢN PHẨM MỚI: TẠO MÃ SP VÀ THÊM VÀO MASTER ---
      curSeqSP++;
      currentMaSP = `SP${String(curSeqSP).padStart(4, '0')}`;
      const spObj = {
        spId: currentMaSP,
        customerId: currentMaKH,
        mang: item.mang || '',
        type: item.type || '',
        donggoi: item.donggoi || '',
        codeSanpham: item.codeSP || item.codeSanpham || '',
        nameNapaco: item.prodName || '',
        donvi: item.unit || item.donvi || '',
        dongia: price,
        giakhuon: giaKhuon
      };
      newProductRows.push(buildRowBySchema(TABLE_MAPPING_SCHEMA.SAN_PHAM_MASTER, infoSP, spObj));

            newCacheRows.push([
              item.prodName || '',
              grp.customerName || '',
              item.codeSP || item.codeSanpham || '',
              item.unit || item.donvi || '',
              price || ''
            ]);

            const newIndex = spExistingData.length + newProductRows.length - 1;
            productMapByCompositeKey[keyWithMaKH] = { spId: currentMaSP, rIdx: newIndex };
            productMapByNameOnly[pNorm] = { spId: currentMaSP, rIdx: newIndex };
          }
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

      // Chuẩn hóa ngày tháng
      const parseDateSafe = (dVal) => {
        if (!dVal) return '';
        if (dVal instanceof Date && !isNaN(dVal.getTime())) {
          return new Date(dVal.getFullYear(), dVal.getMonth(), dVal.getDate(), 12, 0, 0);
        }
        const s = String(dVal).trim();
        if (!s) return '';
        const mISO = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
        if (mISO) {
          return new Date(parseInt(mISO[1], 10), parseInt(mISO[2], 10) - 1, parseInt(mISO[3], 10), 12, 0, 0);
        }
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
    Logger.log(`⏱️ [SERVER] 3. Xử lý RAM & Upsert: ${Date.now() - t3} ms (Upload Drive: ${timeUploadDrive} ms)`);

    // MỐC 4: BATCH WRITING TẤT CẢ CÁC BẢNG (setValues 1 LẦN DUY NHẤT MỖI BẢNG)
    const t4 = Date.now();
    
    // 1. Ghi đè lại bảng sản phẩm trong RAM xuống Sheet Master nếu có sản phẩm cũ được update
    if (hasUpdatedExistingProducts && shSP && spExistingData.length > 0) {
      shSP.getRange(infoSP.startDataRow, 1, spExistingData.length, infoSP.lastCol).setValues(spExistingData);
      Logger.log("✅ [SERVER] Đã cập nhật lại thông tin các sản phẩm cũ trên DB Master.");
    }
    
    // 2. Thêm khách hàng mới
    if (newCustomerRows.length > 0 && shKH) {
      shKH.getRange(shKH.getLastRow() + 1, 1, newCustomerRows.length, infoKH.lastCol).setValues(newCustomerRows);
    }
    
    // 3. Thêm sản phẩm mới
    if (newProductRows.length > 0 && shSP) {
      shSP.getRange(shSP.getLastRow() + 1, 1, newProductRows.length, infoSP.lastCol).setValues(newProductRows);
    }
    
    // 4. Thêm đơn hàng
    if (orderRows.length > 0) {
      shDH.getRange(shDH.getLastRow() + 1, 1, orderRows.length, infoDH.lastCol).setValues(orderRows);
    }
    
    // 5. Thêm chi tiết đơn hàng
    if (detailRows.length > 0) {
      shCT.getRange(shCT.getLastRow() + 1, 1, detailRows.length, infoCT.lastCol).setValues(detailRows);
    }
    Logger.log(`⏱️ [SERVER] 4. Batch Write: ${Date.now() - t4} ms`);

    // MỐC 5: XÓA TRẮNG FORM TAO DON TRÊN GIAO DIỆN HIỆN HÀNH
    const t5 = Date.now();
    try {
      const sheetTaoDonName = CONFIG.SHEET_TAO_DON || 'TAO DON';
      let sheetTaoDon = null;
      if (activeSS) sheetTaoDon = activeSS.getSheetByName(sheetTaoDonName);
      if (!sheetTaoDon) sheetTaoDon = ssOp.getSheetByName(sheetTaoDonName);

      if (sheetTaoDon) {
        const startR = (CONFIG.HEADER_ROW_TAO_DON || 7) + 1;
        const curLastRow = sheetTaoDon.getLastRow();
        const maxBufferRows = curLastRow >= startR ? (curLastRow - startR + 1) : (CONFIG.MAX_ROWS_TAO_DON || 50);
        const maxCols = Math.max(sheetTaoDon.getLastColumn(), 26);
        
        sheetTaoDon.getRange(startR, 1, maxBufferRows, maxCols).clearContent();
        Logger.log(`🧹 [SERVER] Đã dọn sạch biểu mẫu nhập liệu trên sheet "${sheetTaoDonName}".`);
      }
    } catch (errClear) {
      Logger.log("⚠️ Cảnh báo dọn form TAO DON: " + errClear.message);
    }
    Logger.log(`⏱️ [SERVER] 5. Dọn form hoàn tất: ${Date.now() - t5} ms`);

    // MỐC 6: TỰ ĐỘNG ĐỒNG BỘ CACHE
    if (hasUpdatedExistingProducts || newProductRows.length > 0 || newCustomerRows.length > 0) {
      try {
        syncFormDropdowns();
      } catch (eSync) {
        Logger.log("⚠️ Lỗi đồng bộ Cache: " + eSync.message);
      }
    }

    Logger.log(`⏱️ [SERVER] TỔNG THỜI GIAN THỰC THI: ${Date.now() - tTotalStart} ms`);
    return {
      success: true,
      message: `Đã lưu thành công ${orderRows.length} PO và cập nhật danh mục sản phẩm!`
    };

  } finally {
    lock.releaseLock();
    Logger.log("🔒 [SERVER] Đã giải phóng LockService.");
  }
}