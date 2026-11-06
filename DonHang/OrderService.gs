/**
 * ============================================================================
 * NAPACO ENTERPRISE OMS - ORDER SERVICE & REPOSITORY LAYER
 * File: OrderService.gs
 * Tầng kiến trúc: Tầng 2 & 3 - Business Logic & Data Repository
 * Nhiệm vụ:
 *  1. Xác thực người dùng kép (Hybrid Email/Dropdown) đối chiếu bảng '07.NhanVien'.
 *  2. Đọc dữ liệu nháp từ sheet 'TAO DON' (A8:I), gom nhóm nhiều PO của nhiều khách hàng,
 *     tự động gán Ngày đặt = TODAY() và Ngày giao = TODAY() + 10 ngày.
 *  3. Lưu liên tầng (Flow-based Upsert) có LockService (30,000 ms):
 *     - Tự tạo Khách hàng mới (KHxxxx) vào '01_KHACH_HANG'.
 *     - Tự tạo Sản phẩm mới (SPxxxx) hoặc cập nhật thông tin vào '02_SAN_PHAM'.
 *     - Tải tệp đính kèm PO lên Google Drive folder ID_STORAGES.
 *     - Sinh mã đơn hàng 'DH_yymmdd_xx' và mã chi tiết 'CT_yymmdd_xx'.
 *     - Ghi theo lô (Batch Write) vào '01.DonHang' và '02.ChiTietDH' kèm cơ chế Rollback.
 *     - Xóa trắng dữ liệu vùng A8:I trên sheet 'TAO DON' sau khi lưu thành công.
 *  4. Tra cứu kép (Dual-View): PO Header View và Line-Item Flat View.
 * ============================================================================
 */

const OrderService = {

  /**
   * Mở kết nối tới File 01. DB MASTER từ ScriptProperties
   * @private
   */
  getMasterSpreadsheet_: function() {
    const env = getSecureEnv_();
    return SpreadsheetApp.openById(env.ID_MASTER);
  },

  /**
   * Mở kết nối tới File 02. DB OPERATION từ ScriptProperties
   * @private
   */
  getOperationSpreadsheet_: function() {
    const env = getSecureEnv_();
    return SpreadsheetApp.openById(env.ID_OPERATION);
  },

  /**
   * Mở kết nối tới Thư mục lưu trữ File PO (ID_STORAGES)
   * @private
   */
  getStorageFolder_: function() {
    const env = getSecureEnv_();
    return DriveApp.getFolderById(env.ID_STORAGES);
  },

  /**
   * Xóa bộ nhớ đệm Cache khi có cập nhật
   */
  clearSystemCache: function() {
    try {
      const cache = CacheService.getScriptCache();
      cache.removeAll(['NAPACO_MASTER_CATALOG_V2']);
    } catch (e) {
      console.warn('Không thể xóa cache:', e.message);
    }
  },

  /**
   * Đọc danh mục Master (Khách hàng, Sản phẩm, Nhân viên) và Xác thực người dùng kép
   */
  getCatalogAndAuthData: function() {
    const masterSS = this.getMasterSpreadsheet_();

    // 1. Đọc bảng 01_KHACH_HANG: [Ma_KH, name, address, contacter, phone]
    const custSheet = masterSS.getSheetByName(CONFIG.SHEET_NAMES.CUSTOMERS);
    if (!custSheet) throw new Error('Không tìm thấy bảng ' + CONFIG.SHEET_NAMES.CUSTOMERS);
    const custValues = custSheet.getDataRange().getValues();
    const customers = [];
    const customerByNameMap = {};
    const customerByIdMap = {};

    for (let i = 1; i < custValues.length; i++) {
      const maKH = String(custValues[i][0] || '').trim();
      const name = String(custValues[i][1] || '').trim();
      if (!maKH || !name) continue;
      const cusObj = {
        maKH: maKH,
        name: name,
        address: String(custValues[i][2] || '').trim(),
        contacter: String(custValues[i][3] || '').trim(),
        phone: String(custValues[i][4] || '').trim()
      };
      customers.push(cusObj);
      customerByNameMap[name.toLowerCase()] = cusObj;
      customerByIdMap[maKH] = cusObj;
    }

    // 2. Đọc bảng 02_SAN_PHAM (16 cột):
    // [SP_ID, customer_id, mang, type, Code_Sanpham, name_napaco, cavity, tonkho, donvi, dai, rong, cao, imageURL, donggoi, Gia Khuon, dongia]
    const prodSheet = masterSS.getSheetByName(CONFIG.SHEET_NAMES.PRODUCTS);
    if (!prodSheet) throw new Error('Không tìm thấy bảng ' + CONFIG.SHEET_NAMES.PRODUCTS);
    const prodValues = prodSheet.getDataRange().getValues();
    const products = [];
    const productByIdMap = {};

    for (let j = 1; j < prodValues.length; j++) {
      const spId = String(prodValues[j][0] || '').trim();
      const nameNapaco = String(prodValues[j][5] || '').trim();
      if (!spId || !nameNapaco) continue;
      const prodObj = {
        spId: spId,
        customerId: String(prodValues[j][1] || '').trim(),
        mang: String(prodValues[j][2] || '').trim(),
        type: String(prodValues[j][3] || '').trim(),
        codeSP: String(prodValues[j][4] || '').trim(),
        prodName: nameNapaco,
        cavity: Number(prodValues[j][6]) || 0,
        tonkho: Number(prodValues[j][7]) || 0,
        unit: String(prodValues[j][8] || 'Cái').trim() || 'Cái',
        dai: prodValues[j][9] || '',
        rong: prodValues[j][10] || '',
        cao: prodValues[j][11] || '',
        imageURL: String(prodValues[j][12] || '').trim(),
        donggoi: String(prodValues[j][13] || '').trim(),
        giaKhuon: Number(prodValues[j][14]) || 0,
        price: Number(prodValues[j][15]) || 0
      };
      products.push(prodObj);
      productByIdMap[spId] = prodObj;
    }

    // 3. Đọc bảng 07.NhanVien & Xác thực người dùng kép (Hybrid Authentication)
    // Cột: [Ma_NV, name, phong_ban, role, email, sdt, quyen_truy_cap, ngay_tao, trang_thai]
    const empSheet = masterSS.getSheetByName(CONFIG.SHEET_NAMES.EMPLOYEES);
    if (!empSheet) throw new Error('Không tìm thấy bảng ' + CONFIG.SHEET_NAMES.EMPLOYEES);
    const empValues = empSheet.getDataRange().getValues();

    let sessionEmail = '';
    try {
      sessionEmail = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
    } catch (e) {
      sessionEmail = '';
    }

    const salesStaffList = [];
    const allActiveStaffList = [];
    const employeeByIdMap = {};
    let matchedEmployee = null;

    for (let k = 1; k < empValues.length; k++) {
      const maNV = String(empValues[k][0] || '').trim();
      const empName = String(empValues[k][1] || '').trim();
      // Bỏ qua các dòng thống kê cuối bảng (VD: 'Tổng số nhân sự')
      if (!maNV || !/^NV/i.test(maNV) || !empName) continue;

      const phongBan = String(empValues[k][2] || '').trim();
      const role = String(empValues[k][3] || '').trim();
      const email = String(empValues[k][4] || '').trim().toLowerCase();
      const sdt = String(empValues[k][5] || '').trim();
      const quyenTruyCap = String(empValues[k][6] || '').trim();
      const trangThai = String(empValues[k][8] || '').trim();

      const empObj = {
        maNV: maNV,
        hoTen: empName,
        phongBan: phongBan,
        role: role,
        email: email,
        sdt: sdt,
        quyenTruyCap: quyenTruyCap,
        trangThai: trangThai,
        display: maNV + ' - ' + empName
      };

      employeeByIdMap[maNV] = empObj;

      if (trangThai === 'Đang hoạt động') {
        allActiveStaffList.push(empObj);
        if (phongBan.toLowerCase().indexOf('kinh doanh') !== -1 || quyenTruyCap.indexOf('Admin') !== -1 || quyenTruyCap.indexOf('Manager') !== -1) {
          salesStaffList.push(empObj);
        }
        if (sessionEmail && email && sessionEmail === email) {
          matchedEmployee = empObj;
        }
      }
    }

    const finalSalesDropdown = salesStaffList.length > 0 ? salesStaffList : allActiveStaffList;

    const now = new Date();
    const todayStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const deliveryDateObj = new Date(now.getTime() + CONFIG.DEFAULT_DELIVERY_DAYS * 86400000);
    const defaultDeliveryStr = Utilities.formatDate(deliveryDateObj, CONFIG.TIMEZONE, 'yyyy-MM-dd');

    return {
      customers: customers,
      customerByNameMap: customerByNameMap,
      customerByIdMap: customerByIdMap,
      products: products,
      productByIdMap: productByIdMap,
      salesStaffList: finalSalesDropdown,
      employeeByIdMap: employeeByIdMap,
      currentUser: {
        email: sessionEmail,
        isAuthenticated: Boolean(matchedEmployee),
        isLockedReadonly: Boolean(matchedEmployee),
        maNV: matchedEmployee ? matchedEmployee.maNV : '',
        hoTen: matchedEmployee ? matchedEmployee.hoTen : '',
        role: matchedEmployee ? matchedEmployee.role : '',
        quyenTruyCap: matchedEmployee ? matchedEmployee.quyenTruyCap : '',
        display: matchedEmployee ? matchedEmployee.display : ''
      },
      serverDate: todayStr,
      defaultDeliveryDate: defaultDeliveryStr
    };
  },

  /**
   * Tra cứu nhanh thông tin Sản phẩm cho hàm onEdit trên sheet 'TAO DON'
   */
  lookupProductByNameAndCustomer: function(productName, customerName) {
    const catalog = this.getCatalogAndAuthData();
    const pNameLower = String(productName || '').trim().toLowerCase();
    const cNameLower = String(customerName || '').trim().toLowerCase();
    const cusObj = catalog.customerByNameMap[cNameLower];
    const targetCusId = cusObj ? cusObj.maKH : '';

    // Ưu tiên khớp cả tên sản phẩm và mã khách hàng
    let matched = catalog.products.find(function(p) {
      return p.prodName.toLowerCase() === pNameLower && (!targetCusId || p.customerId === targetCusId);
    });
    if (!matched) {
      matched = catalog.products.find(function(p) {
        return p.prodName.toLowerCase() === pNameLower;
      });
    }
    return matched || null;
  },

  /**
   * Xây dựng dữ liệu khởi tạo cho OrderModal.html
   * Quét vùng A8:I trên sheet 'TAO DON', gom nhóm theo (Khách hàng + Số PO),
   * tự động gán Ngày đặt = TODAY() và Ngày giao = TODAY() + 10 ngày.
   */
  buildModalInitialData: function(optSourceSpreadsheetId) {
    const catalog = this.getCatalogAndAuthData();
    const todayStr = catalog.serverDate;
    const defaultDeliveryStr = catalog.defaultDeliveryDate;
    const yymmdd = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyMMdd');

    let wsSheet = null;
    try {
      if (optSourceSpreadsheetId) {
        wsSheet = SpreadsheetApp.openById(optSourceSpreadsheetId).getSheetByName(CONFIG.SHEET_NAMES.ORDER_WORKSPACE);
      } else if (SpreadsheetApp.getActiveSpreadsheet()) {
        wsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.ORDER_WORKSPACE);
      }
    } catch (e) {
      wsSheet = null;
    }

    const poGroupsMap = {};
    const poGroupsList = [];

    if (wsSheet) {
      const lastRow = wsSheet.getLastRow();
      const startRow = CONFIG.WORKSPACE_CONFIG.DATA_START_ROW;
      if (lastRow >= startRow) {
        const numRows = lastRow - startRow + 1;
        const rawRows = wsSheet.getRange(startRow, 1, numRows, CONFIG.WORKSPACE_CONFIG.DATA_END_COL).getValues();

        for (let r = 0; r < rawRows.length; r++) {
          const row = rawRows[r];
          const customerName = String(row[0] || '').trim();
          const prodName = String(row[1] || '').trim();
          const codeSPInput = String(row[2] || '').trim();
          const unitInput = String(row[3] || '').trim();
          const priceInput = this.parseNumber_(row[4]);
          const qtyInput = this.parseNumber_(row[5]);
          const soPOInput = String(row[6] || '').trim();
          const noteInput = String(row[7] || '').trim();
          const linkPOInput = String(row[8] || '').trim();

          // Bỏ qua dòng trống hoàn toàn
          if (!customerName && !prodName) continue;

          const effectiveCustomer = customerName || 'Khách lẻ';
          const cleanCus = this.removeVietnameseTones_(effectiveCustomer)
            .toUpperCase()
            .replace(/[^A-Z0-9]+/g, '-')
            .replace(/^-+|-+$/g, '');
          const autoPONumber = CONFIG.ORDER_CODE_PREFIX + yymmdd + '-' + (cleanCus || 'KH');
          const effectivePO = soPOInput || autoPONumber;

          // Kiểm tra Khách hàng cũ / mới
          const existingCus = catalog.customerByNameMap[effectiveCustomer.toLowerCase()] || null;

          // Kiểm tra Sản phẩm cũ / mới
          let existingProd = null;
          if (prodName) {
            const pLower = prodName.toLowerCase();
            existingProd = catalog.products.find(function(p) {
              return p.prodName.toLowerCase() === pLower && (!existingCus || p.customerId === existingCus.maKH);
            });
            if (!existingProd) {
              existingProd = catalog.products.find(function(p) {
                return p.prodName.toLowerCase() === pLower;
              });
            }
          }

          const groupKey = effectiveCustomer.toLowerCase() + '__' + effectivePO.toUpperCase();
          if (!poGroupsMap[groupKey]) {
            const newGroup = {
              customerName: existingCus ? existingCus.name : effectiveCustomer,
              maKH: existingCus ? existingCus.maKH : '',
              isNewCustomer: !existingCus,
              soPO: effectivePO,
              autoPONumber: autoPONumber,
              salesPerson: catalog.currentUser.isAuthenticated
                ? catalog.currentUser.maNV
                : (catalog.salesStaffList[0] ? catalog.salesStaffList[0].maNV : ''),
              ngayDat: todayStr,
              ngayGiaoDK: defaultDeliveryStr,
              orderNote: '',
              existingFileUrl: linkPOInput || '',
              items: []
            };
            poGroupsMap[groupKey] = newGroup;
            poGroupsList.push(newGroup);
          }

          if (linkPOInput && !poGroupsMap[groupKey].existingFileUrl) {
            poGroupsMap[groupKey].existingFileUrl = linkPOInput;
          }

          poGroupsMap[groupKey].items.push({
            spId: existingProd ? existingProd.spId : '',
            prodName: existingProd ? existingProd.prodName : prodName,
            codeSP: codeSPInput || (existingProd ? existingProd.codeSP : ''),
            unit: unitInput || (existingProd ? existingProd.unit : CONFIG.DEFAULT_UNIT),
            qty: qtyInput > 0 ? qtyInput : 0,
            price: priceInput > 0 ? priceInput : (existingProd ? existingProd.price : 0),
            giaKhuon: existingProd ? existingProd.giaKhuon : 0,
            mang: existingProd ? existingProd.mang : '',
            type: existingProd ? existingProd.type : '1',
            donggoi: existingProd ? existingProd.donggoi : '',
            ngayDat: todayStr,
            ngayGiao: defaultDeliveryStr,
            deliveryDate: defaultDeliveryStr,
            note: noteInput,
            isNew: !existingProd,
            isNewProduct: !existingProd
          });
        }
      }
    }

    // Nếu mở Modal khi sheet TAO DON đang trống -> Khởi tạo sẵn 1 PO Card mặc định
    if (poGroupsList.length === 0) {
      const defaultPO = CONFIG.ORDER_CODE_PREFIX + yymmdd + '-NEW';
      poGroupsList.push({
        customerName: '',
        maKH: '',
        isNewCustomer: false,
        soPO: defaultPO,
        autoPONumber: defaultPO,
        salesPerson: catalog.currentUser.isAuthenticated
          ? catalog.currentUser.maNV
          : (catalog.salesStaffList[0] ? catalog.salesStaffList[0].maNV : ''),
        ngayDat: todayStr,
        ngayGiaoDK: defaultDeliveryStr,
        orderNote: '',
        existingFileUrl: '',
        items: [{
          spId: '',
          prodName: '',
          codeSP: '',
          unit: CONFIG.DEFAULT_UNIT,
          qty: 5000,
          price: 0,
          giaKhuon: 0,
          mang: '',
          type: '1',
          donggoi: '',
          ngayDat: todayStr,
          ngayGiao: defaultDeliveryStr,
          deliveryDate: defaultDeliveryStr,
          note: '',
          isNew: true,
          isNewProduct: true
        }]
      });
    }

    return {
      hasSourceSheet: Boolean(wsSheet),
      poGroups: poGroupsList,
      salesStaffList: catalog.salesStaffList,
      currentUser: catalog.currentUser,
      customers: catalog.customers.map(function(c) {
        return { maKH: c.maKH, name: c.name };
      }),
      products: catalog.products.map(function(p) {
        return {
          spId: p.spId,
          customerId: p.customerId,
          prodName: p.prodName,
          codeSP: p.codeSP,
          unit: p.unit,
          price: p.price,
          mang: p.mang,
          type: p.type,
          donggoi: p.donggoi,
          giaKhuon: p.giaKhuon
        };
      }),
      serverDate: todayStr,
      defaultDeliveryDate: defaultDeliveryStr
    };
  },

  /**
   * ==========================================================================
   * QUY TRÌNH LƯU LIÊN TẦNG (FLOW-BASED UPSERT SEQUENCE) CÓ LOCKSERVICE
   * ==========================================================================
   * Bước 1: Khóa chống xung đột (LockService.getScriptLock().waitLock(30000))
   * Bước 2: Xác thực dữ liệu đầu vào & xác thực nhân viên kinh doanh
   * Bước 3: Quét & Upsert Khách hàng mới (KHxxxx) vào '01_KHACH_HANG'
   * Bước 4: Quét & Upsert Sản phẩm mới (SPxxxx) vào '02_SAN_PHAM'
   * Bước 5: Upload file PO đính kèm (Base64 -> Blob) lên Drive 'ID_STORAGES'
   * Bước 6: Sinh mã Đơn hàng (DH_yymmdd_xx) & Chi tiết (CT_yymmdd_xx)
   * Bước 7: Ghi theo lô đồng thời (Batch Write) kèm cơ chế Rollback nếu lỗi
   * Bước 8: Xóa trắng dữ liệu vùng A8:I trên sheet 'TAO DON' & Giải phóng khóa
   */
  /**
   * Lưu giao dịch đơn hàng liên tầng (ACID Transaction)
   * Nơi duy nhất nắm giữ LockService độc quyền của hệ thống
   */
  saveOrderTransaction: function(payload) {
    const isDebug = typeof CONFIG !== 'undefined' && Boolean(CONFIG.DEBUG_MODE);
    if (isDebug) console.log("⏱️ [ORDER_SERVICE] Bắt đầu tiến trình lưu đơn hàng...");

    const lock = LockService.getScriptLock();
    const timeoutMs = (CONFIG && CONFIG.LOCK_TIMEOUT_MS) ? CONFIG.LOCK_TIMEOUT_MS : 30000;
    
    // 1. Xin khóa độc quyền (Tối đa 30s)
    let hasLock = false;
    try {
      hasLock = lock.tryLock(timeoutMs);
    } catch (e) {
      hasLock = false;
    }

    if (!hasLock) {
      throw new Error("Hệ thống đang bận xử lý giao dịch từ người dùng khác. Vui lòng thử lại sau vài giây! (ERR_LOCK_TIMEOUT)");
    }

    try {
      const masterSS = this.getMasterSpreadsheet_();
      const operSS = this.getOperationSpreadsheet_();

      const custSheet = masterSS.getSheetByName(CONFIG.SHEET_NAMES.CUSTOMERS);
      const prodSheet = masterSS.getSheetByName(CONFIG.SHEET_NAMES.PRODUCTS);
      const orderSheet = operSS.getSheetByName(CONFIG.SHEET_NAMES.ORDERS);
      const detailSheet = operSS.getSheetByName(CONFIG.SHEET_NAMES.ORDER_DETAILS);

      if (!orderSheet || !detailSheet) {
        throw new Error("Không tìm thấy các bảng giao dịch 01.DonHang hoặc 02.ChiTietDH!");
      }

      // 2. Tải bản đồ Khách hàng & Sản phẩm hiện có để kiểm tra cũ/mới
      const custData = custSheet.getDataRange().getValues();
      const khMap = {}; // Tên KH (lowercase) -> Ma_KH
      let maxKhSeq = 0;

      for (let i = 1; i < custData.length; i++) {
        const ma = String(custData[i][0] || '').trim();
        const ten = String(custData[i][1] || '').trim();
        if (ma && ten) khMap[ten.toLowerCase()] = ma;
        
        // Trích xuất số lớn nhất từ các mã KH (VD: KH0214 -> 214)
        const matchKh = ma.match(/KH(\d+)/i);
        if (matchKh) {
          const num = parseInt(matchKh[1], 10);
          if (num > maxKhSeq) maxKhSeq = num;
        }
      }

      const prodData = prodSheet.getDataRange().getValues();
      const prodMap = {}; // Tên SP + Mã KH -> SP_ID
      let maxSpSeq = 0;

      for (let j = 1; j < prodData.length; j++) {
        const spId = String(prodData[j][0] || '').trim();
        const cusId = String(prodData[j][1] || '').trim();
        const spName = String(prodData[j][5] || '').trim().toLowerCase();
        if (spId && spName) {
          prodMap[spName + '__' + cusId] = spId;
          if (!prodMap[spName]) prodMap[spName] = spId;
        }

        // Trích xuất số lớn nhất từ các mã SP (VD: SP2174 -> 2174)
        const matchSp = spId.match(/SP(\d+)/i);
        if (matchSp) {
          const num = parseInt(matchSp[1], 10);
          if (num > maxSpSeq) maxSpSeq = num;
        }
      }

      // Biến tăng tuần tự chuẩn xác cho các bản ghi mới phát sinh trong giao dịch
      let currentMaxKh = maxKhSeq;
      let currentMaxSp = maxSpSeq;

      const poGroups = payload.poGroups || [];
      const createdOrders = [];
      const headerRows = [];
      const detailRows = [];
      const newCustRows = [];
      const newProdRows = [];

      const todayStr = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "yyyy-MM-dd");
      const yymmdd = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "yyMMdd");

      // Đếm số thứ tự đơn hàng và chi tiết trong ngày
      let orderSeq = 1;
      let detailSeq = 1;

      // Quét đơn hàng hiện có để lấy sequence lớn nhất trong ngày
      const existingOrders = orderSheet.getDataRange().getValues();
      const orderRegex = new RegExp(`^DH_${yymmdd}_(\\d+)$`, 'i');
      for (let r = 1; r < existingOrders.length; r++) {
        const match = String(existingOrders[r][0] || '').match(orderRegex);
        if (match) {
          const num = parseInt(match[1], 10);
          if (num >= orderSeq) orderSeq = num + 1;
        }
      }

      const existingDetails = detailSheet.getDataRange().getValues();
      const detailRegex = new RegExp(`^CT_${yymmdd}_(\\d+)$`, 'i');
      for (let r = 1; r < existingDetails.length; r++) {
        const match = String(existingDetails[r][0] || '').match(detailRegex);
        if (match) {
          const num = parseInt(match[1], 10);
          if (num >= detailSeq) detailSeq = num + 1;
        }
      }

      // 3. Xử lý từng khối PO
      for (let g = 0; g < poGroups.length; g++) {
        const grp = poGroups[g];
        const cusName = String(grp.customerName || 'Khách lẻ').trim();
        const cusKey = cusName.toLowerCase();
        let targetMaKH = khMap[cusKey];

        // Nếu là Khách Hàng mới -> Tự sinh mã KHxxxx
        // Nếu là Khách Hàng mới -> Lấy số lớn nhất hiện tại + 1
        if (!targetMaKH) {
          currentMaxKh++;
          targetMaKH = 'KH' + String(currentMaxKh).padStart(4, '0');
          khMap[cusKey] = targetMaKH;
          newCustRows.push([targetMaKH, cusName, '', '', '']);
        }

        // Tải file đính kèm PO lên Drive nếu có
        /*let finalFileUrl = grp.existingFileUrl || '';
        if (grp.fileData && grp.fileData.base64) {
          try {
            const folder = this.getStorageFolder_();
            const decodedBytes = Utilities.base64Decode(grp.fileData.base64);
            const fileName = `PO_${grp.soPO}_${cusName}_${yymmdd}`.replace(/[^a-zA-Z0-9_\-\.]/g, '_');
            const blob = Utilities.newBlob(decodedBytes, grp.fileData.mimeType || 'application/pdf', fileName);
            const uploadedFile = folder.createFile(blob);
            finalFileUrl = uploadedFile.getUrl();
          } catch (fileErr) {
            console.warn("⚠️ Không thể upload tệp PO:", fileErr.message);
          }
        }*/
        // Tải file đính kèm PO lên Drive nếu có (Tên tệp = Số PO)
        let finalFileUrl = grp.existingFileUrl || '';
        if (grp.fileData && grp.fileData.base64) {
          try {
            const folder = this.getStorageFolder_();
            const decodedBytes = Utilities.base64Decode(grp.fileData.base64);
            const ext = (grp.fileData.fileName && grp.fileData.fileName.lastIndexOf('.') !== -1)
              ? grp.fileData.fileName.substring(grp.fileData.fileName.lastIndexOf('.'))
              : '.pdf';
            // Tên tệp lưu trên Drive lấy chính xác theo Số PO
            const safePoName = String(soPO).replace(/[\\/:*?"<>|]/g, '_');
            const finalFileName = safePoName + ext;
            
            const blob = Utilities.newBlob(decodedBytes, grp.fileData.mimeType || 'application/pdf', finalFileName);
            const uploadedFile = folder.createFile(blob);
            finalFileUrl = uploadedFile.getUrl();
          } catch (fileErr) {
            console.warn("⚠️ Không thể upload tệp PO:", fileErr.message);
          }
        }

        // Sinh mã đơn hàng: DH_yymmdd_xx
        const maDH = `DH_${yymmdd}_` + String(orderSeq++).padStart(2, '0');
        const ngayGiaoDK = grp.ngayGiaoDK || todayStr;
        const soPO = grp.soPO || `PO-${yymmdd}-${cusName}`;

        // Header: [Ma_DH, MaKH, NgayDat, NgayGiaoDK, So_PO, File_PO]
        headerRows.push([maDH, targetMaKH, todayStr, ngayGiaoDK, soPO, finalFileUrl]);

        // Xử lý các dòng sản phẩm con
        const items = grp.items || [];
        for (let it = 0; it < items.length; it++) {
          const item = items[it];
          const prodName = String(item.prodName || '').trim();
          const prodKey = prodName.toLowerCase() + '__' + targetMaKH;
          let targetSpId = prodMap[prodKey] || prodMap[prodName.toLowerCase()];

          // Nếu là Sản Phẩm mới -> Tự sinh mã SPxxxx
          // Nếu là Sản Phẩm mới -> Lấy số lớn nhất hiện tại + 1
          if (!targetSpId) {
            currentMaxSp++;
            targetSpId = 'SP' + String(currentMaxSp).padStart(4, '0');
            prodMap[prodKey] = targetSpId;
            newProdRows.push([
              targetSpId,
              targetMaKH,
              item.mang || '',
              item.type || '1',
              item.codeSP || '',
              prodName,
              0, // cavity
              0, // tonkho
              item.unit || 'Cái',
              '', '', '', // dai, rong, cao
              '', // imageURL
              item.donggoi || '',
              item.giaKhuon || 0,
              item.price || 0
            ]);
          }

          // Sinh mã chi tiết: CT_yymmdd_xx
          const idChiTiet = `CT_${yymmdd}_` + String(detailSeq++).padStart(2, '0');
          const slDat = Number(item.qty) || 0;
          const donGia = Number(item.price) || 0;
          const ngayGiao = item.deliveryDate || item.ngayGiao || ngayGiaoDK;
          const maNV = grp.salesPerson ? String(grp.salesPerson).split(' - ')[0].trim() : '';

          // Line Item: [ID_Chi_Tiet, ID_DonHang, Ma_SP, So_Luong_Dat, Don_Gia, So_Luong_Da_Chay, SL_Da_Giao, SL_Con_Lai, Ghi_Chu, Ngay_Giao, Ma_NV, Trang_Thai]
          detailRows.push([
            idChiTiet,
            maDH,
            targetSpId,
            slDat,
            donGia,
            0, // So_Luong_Da_Chay
            0, // SL_Da_Giao
            slDat, // SL_Con_Lai mặc định = So_Luong_Dat
            item.note || grp.orderNote || '',
            ngayGiao,
            maNV,
            'Chờ sản xuất'
          ]);
        }

        createdOrders.push({
          maDH: maDH,
          soPO: soPO,
          itemCount: items.length
        });
      }

      // 4. Batch Write nguyên tử (Atomic Commit)
      if (newCustRows.length > 0) {
        custSheet.getRange(custSheet.getLastRow() + 1, 1, newCustRows.length, newCustRows[0].length).setValues(newCustRows);
      }
      if (newProdRows.length > 0) {
        prodSheet.getRange(prodSheet.getLastRow() + 1, 1, newProdRows.length, newProdRows[0].length).setValues(newProdRows);
      }
      if (headerRows.length > 0) {
        orderSheet.getRange(orderSheet.getLastRow() + 1, 1, headerRows.length, headerRows[0].length).setValues(headerRows);
      }
      if (detailRows.length > 0) {
        detailSheet.getRange(detailSheet.getLastRow() + 1, 1, detailRows.length, detailRows[0].length).setValues(detailRows);
      }

      // 5. Dọn dẹp dòng dữ liệu nháp trên Sheet 'TAO DON'
      if (payload.sourceSpreadsheetId) {
        try {
          const clientSS = SpreadsheetApp.openById(payload.sourceSpreadsheetId);
          const wsSheet = clientSS.getSheetByName(CONFIG.SHEET_NAMES.ORDER_WORKSPACE || 'TAO DON');
          if (wsSheet && wsSheet.getLastRow() >= 8) {
            wsSheet.getRange('A8:I' + wsSheet.getLastRow()).clearContent();
          }
        } catch (wsErr) {
          console.warn("⚠️ Không thể xóa dòng nháp TAO DON:", wsErr.message);
        }
      }

      SpreadsheetApp.flush();
      if (isDebug) console.log("✅ [ORDER_SERVICE] Lưu đơn hàng thành công!");

      return {
        success: true,
        createdOrders: createdOrders,
        newCustomerCount: newCustRows.length,
        newProductCount: newProdRows.length,
        message: `Đã lưu thành công ${createdOrders.length} đơn hàng gồm ${detailRows.length} chi tiết mặt hàng!`
      };

    } finally {
      // 6. BẮT BUỘC: Luôn luôn giải phóng khóa trong khối finally
      lock.releaseLock();
      if (isDebug) console.log("🔒 [ORDER_SERVICE] Đã giải phóng ScriptLock an toàn.");
    }
  },

  /**
   * ==========================================================================
   * TRA CỨU ĐƠN HÀNG & CHI TIẾT MẶT HÀNG (DUAL-VIEW: HEADER & LINE_ITEM)
   * ==========================================================================
   */
  queryOrdersAndLineItems: function(searchParams) {
    const params = searchParams || {};
    const viewMode = String(params.viewMode || 'HEADER').toUpperCase() === 'LINE_ITEM' ? 'LINE_ITEM' : 'HEADER';
    const keyword = String(params.keyword || '').trim().toLowerCase();
    const maKHFilter = String(params.maKH || 'ALL').trim();
    const statusFilter = String(params.status || 'ALL').trim().toLowerCase();
    const fromDate = String(params.fromDate || '').trim();
    const toDate = String(params.toDate || '').trim();
    const page = Math.max(1, parseInt(params.page, 10) || 1);
    const pageSize = Math.max(1, parseInt(params.pageSize, 10) || 100);

    const catalog = this.getCatalogAndAuthData();
    const opSS = this.getOperationSpreadsheet_();
    const orderSheet = opSS.getSheetByName(CONFIG.SHEET_NAMES.ORDERS);
    const detailSheet = opSS.getSheetByName(CONFIG.SHEET_NAMES.ORDER_DETAILS);

    if (!orderSheet || !detailSheet) {
      throw new Error('Không tìm thấy bảng 01.DonHang hoặc 02.ChiTietDH!');
    }

    const orderRows = orderSheet.getDataRange().getValues();
    const detailRows = detailSheet.getDataRange().getValues();

    // 1. Ánh xạ toàn bộ Header Đơn hàng (01.DonHang)
    const orderMap = {};
    const orderHeadersList = [];

    for (let i = 1; i < orderRows.length; i++) {
      const maDH = String(orderRows[i][0] || '').trim();
      if (!maDH) continue;
      const maKH = String(orderRows[i][1] || '').trim();
      const ngayDat = this.formatSheetDate_(orderRows[i][2]);
      const ngayGiaoDK = this.formatSheetDate_(orderRows[i][3]);
      const soPO = String(orderRows[i][4] || '').trim();
      const filePO = String(orderRows[i][5] || '').trim();

      const cusObj = catalog.customerByIdMap[maKH];
      const khachHangName = cusObj ? cusObj.name : maKH;

      const headerObj = {
        maDH: maDH,
        maKH: maKH,
        khachHang: khachHangName,
        ngayDat: ngayDat,
        ngayGiaoDK: ngayGiaoDK,
        soPO: soPO,
        filePO: filePO,
        tongSLDat: 0,
        tongDaGiao: 0,
        tongConLai: 0,
        tongGiaTri: 0,
        itemCount: 0,
        salesPerson: '',
        trangThai: CONFIG.ORDER_STATUS.NEW,
        items: []
      };
      orderMap[maDH] = headerObj;
      orderHeadersList.push(headerObj);
    }

    // 2. Ánh xạ toàn bộ Dòng chi tiết mặt hàng (02.ChiTietDH)
    const allLineItems = [];
    for (let j = 1; j < detailRows.length; j++) {
      const idChiTiet = String(detailRows[j][0] || '').trim();
      const idDonHang = String(detailRows[j][1] || '').trim();
      if (!idChiTiet || !idDonHang) continue;

      const maSP = String(detailRows[j][2] || '').trim();
      const slDat = Number(detailRows[j][3]) || 0;
      const donGia = Number(detailRows[j][4]) || 0;
      const slDaChay = Number(detailRows[j][5]) || 0;
      const slDaGiao = Number(detailRows[j][6]) || 0;
      const slConLaiRaw = detailRows[j][7];
      const slConLai = (slConLaiRaw !== '' && slConLaiRaw !== null && slConLaiRaw !== undefined)
        ? Number(slConLaiRaw)
        : Math.max(0, slDat - slDaGiao);
      const ghiChu = String(detailRows[j][8] || '').trim();
      const ngayGiao = this.formatSheetDate_(detailRows[j][9]);
      const maNV = String(detailRows[j][10] || '').trim();
      const trangThaiRaw = String(detailRows[j][11] || CONFIG.LINE_STATUS.NEW).trim();

      const parentOrder = orderMap[idDonHang];
      const prodObj = catalog.productByIdMap[maSP];
      const empObj = catalog.employeeByIdMap[maNV];

      const sanPhamName = prodObj ? prodObj.prodName : maSP;
      const codeSP = prodObj ? prodObj.codeSP : '';
      const dvt = prodObj ? prodObj.unit : 'Cái';
      const thanhTien = slDat * donGia;

      const lineObj = {
        idChiTiet: idChiTiet,
        maDH: idDonHang,
        soPO: parentOrder ? parentOrder.soPO : '',
        filePO: parentOrder ? parentOrder.filePO : '',
        maKH: parentOrder ? parentOrder.maKH : (prodObj ? prodObj.customerId : ''),
        khachHang: parentOrder ? parentOrder.khachHang : '',
        ngayDat: parentOrder ? parentOrder.ngayDat : '',
        maSP: maSP,
        codeSP: codeSP,
        sanPham: sanPhamName,
        dvt: dvt,
        slDat: slDat,
        donGia: donGia,
        thanhTien: thanhTien,
        daChay: slDaChay,
        daGiao: slDaGiao,
        conLai: slConLai,
        ghiChu: ghiChu,
        ngayGiao: ngayGiao || (parentOrder ? parentOrder.ngayGiaoDK : ''),
        maNV: maNV,
        tenNV: empObj ? empObj.hoTen : maNV,
        trangThai: trangThaiRaw
      };

      allLineItems.push(lineObj);

      if (parentOrder) {
        parentOrder.items.push(lineObj);
        parentOrder.itemCount++;
        parentOrder.tongSLDat += slDat;
        parentOrder.tongDaGiao += slDaGiao;
        parentOrder.tongConLai += slConLai;
        parentOrder.tongGiaTri += thanhTien;
        if (!parentOrder.salesPerson && lineObj.tenNV) {
          parentOrder.salesPerson = lineObj.tenNV;
        }
      }
    }

    // 3. Tính toán Tiến độ & Trạng thái tổng hợp cho từng PO Header
    for (let h = 0; h < orderHeadersList.length; h++) {
      const ord = orderHeadersList[h];
      const progressPct = ord.tongSLDat > 0 ? Math.min(100, Math.round((ord.tongDaGiao / ord.tongSLDat) * 100)) : 0;
      ord.tienDo = progressPct + '% (' + ord.tongDaGiao.toLocaleString('vi-VN') + '/' + ord.tongSLDat.toLocaleString('vi-VN') + ')';
      ord.progressPercent = progressPct;

      if (ord.itemCount > 0 && ord.tongConLai <= 0 && ord.tongSLDat > 0) {
        ord.trangThai = 'Hoàn tất';
      } else if (ord.tongDaGiao > 0 || ord.items.some(function(it) { return it.daChay > 0 || it.trangThai.toLowerCase().indexOf('đang') !== -1; })) {
        ord.trangThai = 'Đang sản xuất';
      } else {
        ord.trangThai = 'Chờ sản xuất';
      }
    }

    // 4. Lọc theo chế độ xem (HEADER hoặc LINE_ITEM)
    if (viewMode === 'LINE_ITEM') {
      const filteredLines = allLineItems.filter(function(item) {
        if (maKHFilter && maKHFilter !== 'ALL' && item.maKH !== maKHFilter && item.khachHang !== maKHFilter) {
          return false;
        }
        if (statusFilter && statusFilter !== 'all') {
          if (item.trangThai.toLowerCase().indexOf(statusFilter) === -1) return false;
        }
        if (fromDate && item.ngayDat && item.ngayDat < fromDate) return false;
        if (toDate && item.ngayDat && item.ngayDat > toDate) return false;
        if (keyword) {
          const haystack = [
            item.idChiTiet, item.maDH, item.soPO, item.khachHang,
            item.maSP, item.codeSP, item.sanPham, item.ghiChu, item.tenNV
          ].join(' ').toLowerCase();
          if (haystack.indexOf(keyword) === -1) return false;
        }
        return true;
      });

      // Sắp xếp mới nhất lên đầu
      filteredLines.reverse();

      const total = filteredLines.length;
      const startIndex = (page - 1) * pageSize;
      const pagedData = filteredLines.slice(startIndex, startIndex + pageSize);

      return {
        success: true,
        viewMode: 'LINE_ITEM',
        total: total,
        summary: this.computeQueryMetrics_(orderHeadersList, filteredLines),
        customers: catalog.customers,
        data: pagedData
      };
    } else {
      const filteredHeaders = orderHeadersList.filter(function(ord) {
        if (maKHFilter && maKHFilter !== 'ALL' && ord.maKH !== maKHFilter && ord.khachHang !== maKHFilter) {
          return false;
        }
        if (statusFilter && statusFilter !== 'all') {
          if (ord.trangThai.toLowerCase().indexOf(statusFilter) === -1) return false;
        }
        if (fromDate && ord.ngayDat && ord.ngayDat < fromDate) return false;
        if (toDate && ord.ngayDat && ord.ngayDat > toDate) return false;
        if (keyword) {
          const itemMatch = ord.items.some(function(it) {
            return (it.sanPham + ' ' + it.codeSP + ' ' + it.maSP + ' ' + it.idChiTiet).toLowerCase().indexOf(keyword) !== -1;
          });
          const headerHaystack = [ord.maDH, ord.soPO, ord.khachHang, ord.maKH, ord.salesPerson].join(' ').toLowerCase();
          if (headerHaystack.indexOf(keyword) === -1 && !itemMatch) return false;
        }
        return true;
      });

      filteredHeaders.reverse();

      const total = filteredHeaders.length;
      const startIndex = (page - 1) * pageSize;
      const pagedData = filteredHeaders.slice(startIndex, startIndex + pageSize);

      return {
        success: true,
        viewMode: 'HEADER',
        total: total,
        summary: this.computeQueryMetrics_(filteredHeaders, allLineItems),
        customers: catalog.customers,
        data: pagedData
      };
    }
  },

  // ==========================================================================
  // CÁC HÀM HELPER NỘI BỘ (PRIVATE METHODS)
  // ==========================================================================

  /**
   * Chuẩn hóa cấu trúc payload từ Modal (hỗ trợ cả chuẩn poGroups và payload đơn)
   * @private
   */
  normalizeAndValidatePayload_: function(payload) {
    if (!payload) {
      throw new Error('Dữ liệu đơn hàng gửi lên trống!');
    }

    let rawGroups = [];
    if (Array.isArray(payload.poGroups)) {
      rawGroups = payload.poGroups;
    } else if (Array.isArray(payload.items)) {
      // Hỗ trợ tương thích nếu gửi từ form đơn 1 PO
      rawGroups = [{
        customerName: payload.customerName || '',
        soPO: payload.soPO || '',
        salesPerson: payload.salesEmployeeId || payload.salesPerson || '',
        ngayDat: payload.orderDate || payload.ngayDat || '',
        ngayGiaoDK: payload.deliveryDate || payload.ngayGiaoDK || '',
        orderNote: payload.notes || payload.orderNote || '',
        fileData: payload.fileData || null,
        items: payload.items.map(function(it) {
          return {
            prodName: it.prodName || it.productName || '',
            codeSP: it.codeSP || '',
            unit: it.unit || 'Cái',
            qty: Number(it.qty || it.quantity) || 0,
            price: Number(it.price || it.unitPrice) || 0,
            ngayDat: it.ngayDat || payload.orderDate || '',
            ngayGiao: it.ngayGiao || it.deliveryDate || payload.deliveryDate || '',
            note: it.note || it.techNotes || ''
          };
        })
      }];
    }

    if (rawGroups.length === 0) {
      throw new Error('Phải có ít nhất 1 đơn hàng (PO) trong giao dịch!');
    }

    const poCustomerCheck = {};
    for (let g = 0; g < rawGroups.length; g++) {
      const grp = rawGroups[g];
      const cusName = String(grp.customerName || '').trim();
      const soPO = String(grp.soPO || '').trim();

      if (!cusName) {
        throw new Error('Đơn PO #' + (g + 1) + ': Tên khách hàng không được để trống!');
      }
      if (!soPO) {
        throw new Error('Đơn PO #' + (g + 1) + ' (' + cusName + '): Số PO không được để trống!');
      }

      const poKey = soPO.toUpperCase();
      if (poCustomerCheck[poKey] && poCustomerCheck[poKey] !== cusName.toLowerCase()) {
        throw new Error('Số PO "' + soPO + '" đang bị gán trùng cho 2 khách hàng khác nhau trong cùng giao dịch!');
      }
      poCustomerCheck[poKey] = cusName.toLowerCase();

      if (!Array.isArray(grp.items) || grp.items.length === 0) {
        throw new Error('Đơn PO "' + soPO + '" phải có ít nhất 1 dòng sản phẩm!');
      }

      for (let i = 0; i < grp.items.length; i++) {
        const it = grp.items[i];
        if (!it.prodName || !String(it.prodName).trim()) {
          throw new Error('Đơn PO "' + soPO + '" - Dòng #' + (i + 1) + ': Tên sản phẩm không được để trống!');
        }
        const qty = Number(it.qty);
        if (isNaN(qty) || qty <= 0) {
          throw new Error('Đơn PO "' + soPO + '" - Dòng #' + (i + 1) + ' (' + it.prodName + '): Số lượng đặt phải lớn hơn 0!');
        }
        const price = Number(it.price);
        if (isNaN(price) || price < 0) {
          throw new Error('Đơn PO "' + soPO + '" - Dòng #' + (i + 1) + ' (' + it.prodName + '): Đơn giá không hợp lệ!');
        }
      }
    }

    return rawGroups;
  },

  /**
   * Đối chiếu Email phiên đăng nhập với bảng 07.NhanVien
   * Nếu khớp email nhân viên đang hoạt động -> Khóa cứng Ma_NV đó ở Backend
   * @private
   */
  resolveAuthenticatedSalesCode_: function(empSheet) {
    let email = '';
    try {
      email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
    } catch (e) {
      email = '';
    }
    if (!email) return '';

    const rows = empSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const maNV = String(rows[i][0] || '').trim();
      const rowEmail = String(rows[i][4] || '').trim().toLowerCase();
      const trangThai = String(rows[i][8] || '').trim();
      if (maNV && /^NV/i.test(maNV) && rowEmail === email && trangThai === 'Đang hoạt động') {
        return maNV;
      }
    }
    return '';
  },

  /**
   * Bóc tách Mã nhân viên (VD: 'NV0001 - Hoàng Thành' -> 'NV0001')
   * @private
   */
  extractStaffCode_: function(salesPersonStr) {
    if (!salesPersonStr) return '';
    const raw = String(salesPersonStr).trim();
    if (raw.indexOf(' - ') !== -1) {
      return raw.split(' - ')[0].trim();
    }
    return raw;
  },

  /**
   * Tải tệp đính kèm PO lên Google Drive (ID_STORAGES)
   * Tự động đổi tên chuẩn: PO_[Số PO]_[Tên KH].[ext]
   * @private
   */
  uploadPoAttachmentToDrive_: function(fileObj, soPO, customerName) {
    try {
      const base64Data = String(fileObj.base64 || '').replace(/^data:[^;]+;base64,/, '');
      if (!base64Data) return '';

      const decodedBytes = Utilities.base64Decode(base64Data);
      if (decodedBytes.length > CONFIG.MAX_FILE_SIZE_BYTES) {
        throw new Error('Tệp đính kèm của PO "' + soPO + '" vượt quá dung lượng tối đa 15MB!');
      }

      const origName = String(fileObj.fileName || fileObj.originalName || 'PO.pdf');
      const ext = origName.indexOf('.') !== -1 ? origName.split('.').pop().toLowerCase() : 'pdf';
      const cleanPO = this.removeVietnameseTones_(soPO).replace(/[^a-zA-Z0-9_-]/g, '_');
      const cleanCus = this.removeVietnameseTones_(customerName).replace(/[^a-zA-Z0-9_-]/g, '_');
      const standardFileName = 'PO_' + cleanPO + '_' + cleanCus + '.' + ext;

      const mimeType = fileObj.mimeType || this.guessMimeType_(ext);
      const blob = Utilities.newBlob(decodedBytes, mimeType, standardFileName);
      const folder = this.getStorageFolder_();
      const createdFile = folder.createFile(blob);

      return createdFile.getUrl();
    } catch (err) {
      console.warn('[Upload PO File Warning]: ' + err.message);
      return '';
    }
  },

  /**
   * Cập nhật thông tin sản phẩm cũ trong 02_SAN_PHAM nếu đơn hàng mới có bổ sung thông số
   * @private
   */
  updateExistingProductIfChanged_: function(prodSheet, matchedProd, item) {
    try {
      const rIdx = matchedProd.rowIndex;
      if (item.codeSP && item.codeSP !== matchedProd.codeSP) {
        prodSheet.getRange(rIdx, 5).setValue(item.codeSP);
        matchedProd.codeSP = item.codeSP;
      }
      if (Number(item.price) > 0 && Number(item.price) !== matchedProd.price) {
        prodSheet.getRange(rIdx, 16).setValue(Number(item.price));
        matchedProd.price = Number(item.price);
      }
      if (Number(item.giaKhuon) > 0 && Number(item.giaKhuon) !== matchedProd.giaKhuon) {
        prodSheet.getRange(rIdx, 15).setValue(Number(item.giaKhuon));
        matchedProd.giaKhuon = Number(item.giaKhuon);
      }
      if (item.mang && item.mang !== matchedProd.mang) {
        prodSheet.getRange(rIdx, 3).setValue(item.mang);
        matchedProd.mang = item.mang;
      }
      if (item.donggoi && item.donggoi !== matchedProd.donggoi) {
        prodSheet.getRange(rIdx, 14).setValue(item.donggoi);
        matchedProd.donggoi = item.donggoi;
      }
    } catch (e) {
      console.warn('[Update Product Master Warning]: ' + e.message);
    }
  },

  /**
   * Xóa trắng vùng dữ liệu A8:I trên sheet 'TAO DON' sau khi lưu đơn thành công
   * Giữ nguyên công thức, định dạng và Data Validation
   * @private
   */
  clearWorkspaceSheetAfterSave_: function(optSourceSpreadsheetId) {
    try {
      let ss = null;
      if (optSourceSpreadsheetId) {
        ss = SpreadsheetApp.openById(optSourceSpreadsheetId);
      } else {
        ss = SpreadsheetApp.getActiveSpreadsheet();
      }
      if (!ss) return;

      const wsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ORDER_WORKSPACE);
      if (!wsSheet) return;

      const maxRows = wsSheet.getMaxRows();
      if (maxRows >= CONFIG.WORKSPACE_CONFIG.DATA_START_ROW) {
        wsSheet.getRange(CONFIG.WORKSPACE_CONFIG.CLEAR_RANGE_A1).clearContent();
      }
    } catch (e) {
      console.warn('[Clear TAO DON Warning]: ' + e.message);
    }
  },

  /**
   * Tính toán các chỉ số thống kê tổng quan cho giao diện Tra cứu
   * @private
   */
  computeQueryMetrics_: function(headers, lines) {
    let totalRevenue = 0;
    let totalRemainingQty = 0;
    headers.forEach(function(h) {
      totalRevenue += (Number(h.tongGiaTri) || 0);
    });
    lines.forEach(function(l) {
      totalRemainingQty += (Number(l.conLai) || 0);
    });
    return {
      totalOrders: headers.length,
      totalLineItems: lines.length,
      totalRevenue: totalRevenue,
      totalRemainingQty: totalRemainingQty
    };
  },

  /**
   * Chuẩn hóa chuỗi ngày về yyyy-MM-dd
   * @private
   */
  normalizeDateStr_: function(val) {
    if (!val) return '';
    if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
      return Utilities.formatDate(val, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    const s = String(val).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (dmy) {
      return dmy[3] + '-' + this.padZero_(dmy[2], 2) + '-' + this.padZero_(dmy[1], 2);
    }
    const parsed = new Date(s);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    return '';
  },

  /**
   * Định dạng giá trị ngày đọc từ Google Sheet về yyyy-MM-dd
   * @private
   */
  formatSheetDate_: function(val) {
    return this.normalizeDateStr_(val);
  },

  /**
   * Chuyển yyyy-MM-dd sang yymmdd (VD: '2026-11-04' -> '261104')
   * @private
   */
  dateStrToYYMMDD_: function(dateStr) {
    const norm = this.normalizeDateStr_(dateStr);
    if (norm && /^\d{4}-\d{2}-\d{2}$/.test(norm)) {
      return norm.substring(2, 4) + norm.substring(5, 7) + norm.substring(8, 10);
    }
    return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyMMdd');
  },

  /**
   * Đệm số 0 ở đầu
   * @private
   */
  padZero_: function(num, size) {
    let s = String(num);
    while (s.length < size) s = '0' + s;
    return s;
  },

  /**
   * Chuyển đổi chuỗi số có dấu phân cách sang Number
   * @private
   */
  parseNumber_: function(val) {
    if (typeof val === 'number') return isNaN(val) ? 0 : val;
    if (!val) return 0;
    const cleaned = String(val).replace(/\./g, '').replace(/,/g, '.').replace(/[^\d.-]/g, '');
    const n = parseFloat(cleaned);
    return isNaN(n) ? 0 : n;
  },

  /**
   * Loại bỏ dấu tiếng Việt
   * @private
   */
  removeVietnameseTones_: function(str) {
    if (!str) return '';
    return String(str)
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/đ/g, 'd')
      .replace(/Đ/g, 'D');
  },

  /**
   * Xác định MimeType theo đuôi file
   * @private
   */
  guessMimeType_: function(ext) {
    const map = {
      'pdf': 'application/pdf',
      'png': 'image/png',
      'jpg': 'image/jpeg',
      'jpeg': 'image/jpeg',
      'doc': 'application/msword',
      'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'xls': 'application/vnd.ms-excel',
      'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    };
    return map[ext] || 'application/octet-stream';
  },

  /**
   * Đồng bộ danh mục Khách hàng, Sản phẩm & Nhân viên KD từ DB MASTER sang _CACHE_DANH_MUC
   * Cú pháp giữ nguyên chuẩn Object Literal bên trong OrderService
   */
  /**
 * NAPACO ENTERPRISE OMS - ORDER SERVICE
 * File: OrderService.gs
 * Target: Business Logic Tier
 */

  /**
   * Đồng bộ danh mục Khách hàng, Sản phẩm & Nhân viên KD từ DB MASTER sang _CACHE_DANH_MUC
   * Cú pháp giữ nguyên chuẩn Object Literal bên trong OrderService
   */
  /**
   * Đồng bộ danh mục Khách hàng, Sản phẩm & Nhân viên KD từ DB MASTER sang _CACHE_DANH_MUC
   * Cú pháp giữ nguyên chuẩn Object Literal bên trong OrderService
   */
  syncCacheCatalogFromDatabase: function(optTargetSpreadsheetId) {
    const tStart = Date.now();
    const isDebug = typeof CONFIG !== 'undefined' && Boolean(CONFIG.DEBUG_MODE);

    if (isDebug) console.log("⏱️ [BACKEND-SYNC] Bắt đầu đồng bộ danh mục chuẩn...");

    const targetSS = optTargetSpreadsheetId 
      ? SpreadsheetApp.openById(optTargetSpreadsheetId) 
      : SpreadsheetApp.getActiveSpreadsheet();
    if (!targetSS) throw new Error("Không xác định được file Client đích!");

    const masterSS = this.getMasterSpreadsheet_();

    // 1. Tạo bản đồ: Ma_KH -> Tên Khách Hàng
    const custSheet = masterSS.getSheetByName(CONFIG.SHEET_NAMES.CUSTOMERS);
    const custData = custSheet.getDataRange().getValues();
    const khMap = {};
    const allCustomerNames = [];
    for (let i = 1; i < custData.length; i++) {
      const ma = String(custData[i][0] || '').trim();
      const ten = String(custData[i][1] || '').trim();
      if (ma && ten) {
        khMap[ma] = ten;
        if (!allCustomerNames.includes(ten)) allCustomerNames.push(ten);
      }
    }

    // 2. Đọc bảng 07.NhanVien -> Nhận diện động dòng tiêu đề & cột
    const empSheet = masterSS.getSheetByName(CONFIG.SHEET_NAMES.EMPLOYEES);
    const empData = empSheet ? empSheet.getDataRange().getValues() : [];
    const salesStaffRows = [];

    let empHeaderRowIdx = -1;
    let colMaNV = -1, colTenNV = -1, colPhongBan = -1, colTrangThai = -1;

    for (let r = 0; r < Math.min(empData.length, 10); r++) {
      for (let c = 0; c < empData[r].length; c++) {
        const val = String(empData[r][c] || '').trim().toLowerCase();
        if (val === 'ma_nv' || val === 'mã nv') {
          empHeaderRowIdx = r;
          break;
        }
      }
      if (empHeaderRowIdx !== -1) break;
    }

    if (empHeaderRowIdx !== -1) {
      const empHeaders = empData[empHeaderRowIdx];
      for (let c = 0; c < empHeaders.length; c++) {
        const h = String(empHeaders[c] || '').trim().toLowerCase();
        if (h === 'ma_nv' || h === 'mã nv') colMaNV = c;
        if (h === 'name' || h === 'họ và tên' || h === 'tên nv') colTenNV = c;
        if (h === 'phong_ban' || h === 'phòng ban') colPhongBan = c;
        if (h === 'trang_thai' || h === 'trạng thái') colTrangThai = c;
      }

      for (let k = empHeaderRowIdx + 1; k < empData.length; k++) {
        const row = empData[k];
        const maNV = colMaNV !== -1 ? String(row[colMaNV] || '').trim() : '';
        const tenNV = colTenNV !== -1 ? String(row[colTenNV] || '').trim() : '';
        const pb = colPhongBan !== -1 ? String(row[colPhongBan] || '').trim().toLowerCase() : '';
        const tt = colTrangThai !== -1 ? String(row[colTrangThai] || '').trim().toLowerCase() : '';

        if (maNV.toUpperCase().startsWith('NV') && tt.includes('hoạt động') && pb.includes('kinh doanh')) {
          salesStaffRows.push([`${maNV} - ${tenNV}`]);
        }
      }
    }

    if (isDebug) {
      console.log("⏱️ [DEBUG_CHECK] 1. Dòng tiêu đề Nhân viên phát hiện tại Row index:", empHeaderRowIdx);
      console.log("⏱️ [DEBUG_CHECK] 2. Lọc được " + salesStaffRows.length + " nhân viên KD đang hoạt động: " + JSON.stringify(salesStaffRows));
    }

    // 3. Đọc Sản Phẩm và ghép ĐÚNG Khách Hàng sở hữu qua customer_id
    const prodSheet = masterSS.getSheetByName(CONFIG.SHEET_NAMES.PRODUCTS);
    const prodData = prodSheet.getDataRange().getValues();
    const cacheRows = [];
    for (let j = 1; j < prodData.length; j++) {
      const tenSP = String(prodData[j][5] || '').trim();
      if (!tenSP) continue;

      const cusId = String(prodData[j][1] || '').trim();
      const tenKH = khMap[cusId] || cusId;

      cacheRows.push([
        tenSP,
        tenKH,
        String(prodData[j][4] || '').trim(),
        String(prodData[j][8] || 'Cái').trim() || 'Cái',
        Number(prodData[j][15]) || 0
      ]);
    }

    // 4. Ghi đè vào sheet _CACHE_DANH_MUC
    let cacheSheet = targetSS.getSheetByName(CONFIG.SHEET_NAMES.CACHE_CATALOG || '_CACHE_DANH_MUC');
    if (!cacheSheet) cacheSheet = targetSS.insertSheet(CONFIG.SHEET_NAMES.CACHE_CATALOG || '_CACHE_DANH_MUC');
    cacheSheet.clearContents();

    const headers = [['Tên Sản Phẩm', 'Khách Hàng', 'Code SP', 'Đơn Vị', 'Đơn Giá Chuẩn (₫)']];
    cacheSheet.getRange(1, 1, 1, 5).setValues(headers).setFontWeight('bold');
    if (cacheRows.length > 0) {
      cacheSheet.getRange(2, 1, cacheRows.length, 5).setValues(cacheRows);
    }

    // Ghi Danh sách NV KD vào Cột G (Cột 7)
    cacheSheet.getRange(1, 7).setValue('Danh Sách NV KD').setFontWeight('bold');
    if (salesStaffRows.length > 0) {
      cacheSheet.getRange(2, 7, salesStaffRows.length, 1).setValues(salesStaffRows);
      if (isDebug) {
        console.log("⏱️ [DEBUG_CHECK] 3. Đã ghi thành công " + salesStaffRows.length + " nhân viên vào Cột G của sheet _CACHE_DANH_MUC.");
      }
    }

    // Cột K (11): Danh sách tất cả khách hàng (Dropdown cột Khách Hàng)
    cacheSheet.getRange(1, 11).setValue('Danh Sách Tất Cả KH').setFontWeight('bold');
    if (allCustomerNames.length > 0) {
      cacheSheet.getRange(2, 11, allCustomerNames.length, 1).setValues(allCustomerNames.map(c => [c]));
    }

    // 5. Gán Data Validation Dropdown cho sheet TAO DON theo tên cột động từ CONFIG
    const wsSheet = targetSS.getSheetByName(CONFIG.SHEET_NAMES.ORDER_WORKSPACE || 'TAO DON');
    if (wsSheet) {
      const headerRow = wsSheet.getRange(CONFIG.FORM_COLUMNS.HEADER_ROW, 1, 1, wsSheet.getLastColumn()).getValues()[0];
      const findCol = (aliases) => {
        for (let c = 0; c < headerRow.length; c++) {
          const text = String(headerRow[c] || '').trim().toLowerCase();
          if (aliases.some(a => text.includes(a.toLowerCase()))) return c + 1;
        }
        return -1;
      };

      const colKh = findCol(CONFIG.FORM_COLUMNS.ALIASES.CUSTOMER);
      const colSp = findCol(CONFIG.FORM_COLUMNS.ALIASES.PRODUCT);
      const startRow = CONFIG.FORM_COLUMNS.DATA_START_ROW;
      const numRows = 100;

      if (colKh > 0 && allCustomerNames.length > 0) {
        wsSheet.getRange(startRow, colKh, numRows, 1).setDataValidation(
          SpreadsheetApp.newDataValidation().requireValueInRange(cacheSheet.getRange(2, 11, allCustomerNames.length, 1), true).setAllowInvalid(true).build()
        );
      }
      if (colSp > 0 && cacheRows.length > 0) {
        wsSheet.getRange(startRow, colSp, numRows, 1).setDataValidation(
          SpreadsheetApp.newDataValidation().requireValueInRange(cacheSheet.getRange(2, 1, cacheRows.length, 1), true).setAllowInvalid(true).build()
        );
      }
    }

    SpreadsheetApp.flush();
    if (isDebug) console.log("⏱️ [BACKEND-SYNC] Hoàn tất sau " + (Date.now() - tStart) + " ms.");

    return {
      success: true,
      message: `Đã đồng bộ chính xác ${allCustomerNames.length} khách hàng, ${cacheRows.length} sản phẩm và ${salesStaffRows.length} nhân viên KD!`
    };
  },
}