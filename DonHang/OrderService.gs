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
  saveOrderTransaction: function(payload) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    } catch (lockErr) {
      const err = new Error('Hệ thống đang bận xử lý giao dịch từ người dùng khác. Vui lòng thử lại sau vài giây! (ERR_LOCK_TIMEOUT)');
      err.code = 'ERR_LOCK_TIMEOUT';
      throw err;
    }

    // Theo dõi các vùng đã ghi để Rollback nếu xảy ra lỗi giữa chừng
    const rollbackTracker = [];

    try {
      // 1. Chuẩn hóa & Validate Payload
      const normalizedGroups = this.normalizeAndValidatePayload_(payload);

      // 2. Mở kết nối Database Master & Operation
      const masterSS = this.getMasterSpreadsheet_();
      const opSS = this.getOperationSpreadsheet_();

      const custSheet = masterSS.getSheetByName(CONFIG.SHEET_NAMES.CUSTOMERS);
      const prodSheet = masterSS.getSheetByName(CONFIG.SHEET_NAMES.PRODUCTS);
      const empSheet = masterSS.getSheetByName(CONFIG.SHEET_NAMES.EMPLOYEES);
      const orderSheet = opSS.getSheetByName(CONFIG.SHEET_NAMES.ORDERS);
      const detailSheet = opSS.getSheetByName(CONFIG.SHEET_NAMES.ORDER_DETAILS);

      if (!custSheet || !prodSheet || !orderSheet || !detailSheet) {
        throw new Error('Không tìm thấy đầy đủ các bảng dữ liệu trong DB MASTER hoặc DB OPERATION!');
      }

      // 3. Kiểm tra xác thực người dùng kép ở Backend
      const verifiedSalesCode = this.resolveAuthenticatedSalesCode_(empSheet);

      // 4. Đọc dữ liệu hiện tại của 01_KHACH_HANG để sinh mã KHxxxx và tránh trùng
      const custData = custSheet.getDataRange().getValues();
      const custNameMap = {};
      let maxCustNum = 0;
      for (let i = 1; i < custData.length; i++) {
        const cId = String(custData[i][0] || '').trim();
        const cName = String(custData[i][1] || '').trim();
        if (cId) {
          const num = parseInt(cId.replace(/\D/g, ''), 10);
          if (!isNaN(num) && num > maxCustNum) maxCustNum = num;
        }
        if (cId && cName) {
          custNameMap[cName.toLowerCase()] = cId;
        }
      }

      // 5. Đọc dữ liệu hiện tại của 02_SAN_PHAM để sinh mã SPxxxx và đối chiếu cập nhật
      const prodData = prodSheet.getDataRange().getValues();
      let maxProdNum = 0;
      const prodLookupList = [];
      for (let j = 1; j < prodData.length; j++) {
        const spId = String(prodData[j][0] || '').trim();
        const cusId = String(prodData[j][1] || '').trim();
        const pName = String(prodData[j][5] || '').trim();
        if (spId && /^SP\d+$/i.test(spId)) {
          const pNum = parseInt(spId.substring(2), 10);
          if (!isNaN(pNum) && pNum > maxProdNum) maxProdNum = pNum;
        }
        if (spId && pName) {
          prodLookupList.push({
            rowIndex: j + 1,
            spId: spId,
            customerId: cusId,
            prodName: pName,
            prodNameLower: pName.toLowerCase(),
            mang: String(prodData[j][2] || '').trim(),
            type: String(prodData[j][3] || '').trim(),
            codeSP: String(prodData[j][4] || '').trim(),
            unit: String(prodData[j][8] || '').trim(),
            donggoi: String(prodData[j][13] || '').trim(),
            giaKhuon: Number(prodData[j][14]) || 0,
            price: Number(prodData[j][15]) || 0
          });
        }
      }

      // 6. Đọc dữ liệu hiện tại của 01.DonHang & 02.ChiTietDH để lấy số thứ tự lớn nhất theo từng ngày yymmdd
      const orderData = orderSheet.getDataRange().getValues();
      const dhSeqByDateMap = {};
      for (let r = 1; r < orderData.length; r++) {
        const maDH = String(orderData[r][0] || '').trim();
        const matchDH = maDH.match(/^DH_(\d{6})_(\d+)$/);
        if (matchDH) {
          const dKey = matchDH[1];
          const seq = parseInt(matchDH[2], 10);
          if (!dhSeqByDateMap[dKey] || seq > dhSeqByDateMap[dKey]) {
            dhSeqByDateMap[dKey] = seq;
          }
        }
      }

      const detailData = detailSheet.getDataRange().getValues();
      const ctSeqByDateMap = {};
      for (let d = 1; d < detailData.length; d++) {
        const idCT = String(detailData[d][0] || '').trim();
        const matchCT = idCT.match(/^CT_(\d{6})_(\d+)$/);
        if (matchCT) {
          const dKey = matchCT[1];
          const seq = parseInt(matchCT[2], 10);
          if (!ctSeqByDateMap[dKey] || seq > ctSeqByDateMap[dKey]) {
            ctSeqByDateMap[dKey] = seq;
          }
        }
      }

      // Mảng lưu các bản ghi mới cần Batch Write
      const newCustomerRows = [];
      const newProductRows = [];
      const newOrderHeaderRows = [];
      const newOrderDetailRows = [];
      const createdOrdersSummary = [];

      const todayFallback = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
      const defaultDelivFallback = Utilities.formatDate(
        new Date(Date.now() + CONFIG.DEFAULT_DELIVERY_DAYS * 86400000),
        CONFIG.TIMEZONE,
        'yyyy-MM-dd'
      );

      // 7. Duyệt từng khối PO trong giao dịch
      for (let g = 0; g < normalizedGroups.length; g++) {
        const grp = normalizedGroups[g];
        const cusNameClean = grp.customerName.trim();
        const cusKey = cusNameClean.toLowerCase();

        // --- BƯỚC 2 TRONG SPEC: UPSERT KHÁCH HÀNG ---
        let maKH = custNameMap[cusKey];
        if (!maKH) {
          maxCustNum++;
          maKH = CONFIG.KH_PREFIX + this.padZero_(maxCustNum, 4);
          custNameMap[cusKey] = maKH;
          // Schema 01_KHACH_HANG: [Ma_KH, name, address, contacter, phone]
          newCustomerRows.push([maKH, cusNameClean, '', '', '']);
        }

        // Xác định Nhân viên Kinh doanh phụ trách PO
        const salesCode = verifiedSalesCode || this.extractStaffCode_(grp.salesPerson);
        if (!salesCode) {
          throw new Error('Đơn PO "' + grp.soPO + '": Không xác định được Mã nhân viên kinh doanh phụ trách!');
        }

        // --- BƯỚC 4 TRONG SPEC: UPLOAD FILE PO LÊN DRIVE ---
        let filePoUrl = String(grp.existingFileUrl || '').trim();
        const fileObj = grp.fileData || (grp.fileBase64 ? { fileName: grp.fileName, base64: grp.fileBase64, mimeType: grp.mimeType } : null);
        if (fileObj && fileObj.base64) {
          filePoUrl = this.uploadPoAttachmentToDrive_(fileObj, grp.soPO, cusNameClean);
        }

        // Tính toán Ngày đặt của PO và Ngày giao dự kiến xa nhất (TC-02)
        let poNgayDat = this.normalizeDateStr_(grp.ngayDat) || todayFallback;
        let maxNgayGiaoDK = '';

        for (let i = 0; i < grp.items.length; i++) {
          const itemNgayDat = this.normalizeDateStr_(grp.items[i].ngayDat) || poNgayDat;
          const itemNgayGiao = this.normalizeDateStr_(grp.items[i].ngayGiao || grp.items[i].deliveryDate) || defaultDelivFallback;
          if (i === 0 && !grp.ngayDat) {
            poNgayDat = itemNgayDat;
          }
          if (!maxNgayGiaoDK || itemNgayGiao > maxNgayGiaoDK) {
            maxNgayGiaoDK = itemNgayGiao;
          }
        }
        if (!maxNgayGiaoDK) {
          maxNgayGiaoDK = this.normalizeDateStr_(grp.ngayGiaoDK) || defaultDelivFallback;
        }

        // --- BƯỚC 5 TRONG SPEC: SINH MÃ ĐƠN HÀNG (DH_yymmdd_xx) ---
        const yymmddOrder = this.dateStrToYYMMDD_(poNgayDat);
        const nextDhSeq = (dhSeqByDateMap[yymmddOrder] || 0) + 1;
        dhSeqByDateMap[yymmddOrder] = nextDhSeq;
        const maDH = CONFIG.DH_PREFIX + yymmddOrder + '_' + this.padZero_(nextDhSeq, 2);

        // Schema 01.DonHang (6 cột): [Ma_DH, MaKH, NgayDat, NgayGiaoDK, So_PO, File_PO]
        newOrderHeaderRows.push([
          maDH,
          maKH,
          poNgayDat,
          maxNgayGiaoDK,
          grp.soPO,
          filePoUrl
        ]);

        // Xử lý từng dòng mặt hàng trong PO
        for (let idx = 0; idx < grp.items.length; idx++) {
          const item = grp.items[idx];
          const prodNameClean = item.prodName.trim();
          const prodNameLower = prodNameClean.toLowerCase();

          // --- BƯỚC 3 TRONG SPEC: UPSERT SẢN PHẨM ---
          let matchedProd = prodLookupList.find(function(p) {
            return p.prodNameLower === prodNameLower && p.customerId === maKH;
          });
          if (!matchedProd) {
            matchedProd = prodLookupList.find(function(p) {
              return p.prodNameLower === prodNameLower;
            });
          }

          let maSP = '';
          if (!matchedProd) {
            maxProdNum++;
            maSP = CONFIG.SP_PREFIX + this.padZero_(maxProdNum, 4);
            const newProdRecord = [
              maSP,                             // 1. SP_ID
              maKH,                             // 2. customer_id
              item.mang || '',                  // 3. mang
              item.type || '1',                 // 4. type
              item.codeSP || '',                // 5. Code_Sanpham
              prodNameClean,                    // 6. name_napaco
              '',                               // 7. cavity
              0,                                // 8. tonkho
              item.unit || CONFIG.DEFAULT_UNIT, // 9. donvi
              '',                               // 10. dai
              '',                               // 11. rong
              '',                               // 12. cao
              '',                               // 13. imageURL
              item.donggoi || '',               // 14. donggoi
              Number(item.giaKhuon) || 0,       // 15. Gia Khuon
              Number(item.price) || 0           // 16. dongia
            ];
            newProductRows.push(newProdRecord);
            prodLookupList.push({
              rowIndex: -1,
              spId: maSP,
              customerId: maKH,
              prodName: prodNameClean,
              prodNameLower: prodNameLower,
              mang: item.mang || '',
              type: item.type || '1',
              codeSP: item.codeSP || '',
              unit: item.unit || CONFIG.DEFAULT_UNIT,
              donggoi: item.donggoi || '',
              giaKhuon: Number(item.giaKhuon) || 0,
              price: Number(item.price) || 0
            });
          } else {
            maSP = matchedProd.spId;
            // Cập nhật thông tin tham chiếu nếu sản phẩm cũ có bổ sung Code SP / Đơn giá / Thông số
            if (matchedProd.rowIndex > 1) {
              this.updateExistingProductIfChanged_(prodSheet, matchedProd, item);
            }
          }

          // Sinh mã dòng chi tiết (CT_yymmdd_xx)
          const itemOrderDate = this.normalizeDateStr_(item.ngayDat) || poNgayDat;
          const yymmddLine = this.dateStrToYYMMDD_(itemOrderDate);
          const nextCtSeq = (ctSeqByDateMap[yymmddLine] || 0) + 1;
          ctSeqByDateMap[yymmddLine] = nextCtSeq;
          const idChiTiet = CONFIG.CT_PREFIX + yymmddLine + '_' + this.padZero_(nextCtSeq, 2);

          const slDat = Number(item.qty) || 0;
          const donGia = Number(item.price) || 0;
          const itemDeliveryDate = this.normalizeDateStr_(item.ngayGiao || item.deliveryDate) || maxNgayGiaoDK;

          // Kết hợp ghi chú đơn hàng và ghi chú dòng nếu có
          let combinedNote = String(item.note || '').trim();
          if (grp.orderNote && String(grp.orderNote).trim()) {
            combinedNote = combinedNote
              ? (combinedNote + ' | PO: ' + String(grp.orderNote).trim())
              : String(grp.orderNote).trim();
          }

          // Schema 02.ChiTietDH (12 cột):
          // [ID_Chi_Tiet, ID_DonHang, Ma_SP, So_Luong_Dat, Don_Gia, So_Luong_Da_Chay, SL_Da_Giao, SL_Con_Lai, Ghi_Chu, Ngay_Giao, Ma_NV, Trang_Thai]
          newOrderDetailRows.push([
            idChiTiet,              // 1. ID_Chi_Tiet
            maDH,                   // 2. ID_DonHang
            maSP,                   // 3. Ma_SP
            slDat,                  // 4. So_Luong_Dat
            donGia,                 // 5. Don_Gia
            0,                      // 6. So_Luong_Da_Chay
            0,                      // 7. SL_Da_Giao
            slDat,                  // 8. SL_Con_Lai (= So_Luong_Dat - 0)
            combinedNote,           // 9. Ghi_Chu
            itemDeliveryDate,       // 10. Ngay_Giao
            salesCode,              // 11. Ma_NV
            CONFIG.LINE_STATUS.NEW  // 12. Trang_Thai ('Chờ sản xuất')
          ]);
        }

        createdOrdersSummary.push({
          maDH: maDH,
          soPO: grp.soPO,
          customerName: cusNameClean,
          maKH: maKH,
          itemCount: grp.items.length
        });
      }

      // --- BƯỚC 6 TRONG SPEC: GHI THEO LÔ ĐỒNG THỜI (BATCH WRITE) ---
      if (newCustomerRows.length > 0) {
        const startRow = custSheet.getLastRow() + 1;
        custSheet.getRange(startRow, 1, newCustomerRows.length, newCustomerRows[0].length).setValues(newCustomerRows);
        rollbackTracker.push({ sheet: custSheet, startRow: startRow, numRows: newCustomerRows.length });
      }

      if (newProductRows.length > 0) {
        const startRow = prodSheet.getLastRow() + 1;
        prodSheet.getRange(startRow, 1, newProductRows.length, newProductRows[0].length).setValues(newProductRows);
        rollbackTracker.push({ sheet: prodSheet, startRow: startRow, numRows: newProductRows.length });
      }

      if (newOrderHeaderRows.length > 0) {
        const startRow = orderSheet.getLastRow() + 1;
        orderSheet.getRange(startRow, 1, newOrderHeaderRows.length, newOrderHeaderRows[0].length).setValues(newOrderHeaderRows);
        rollbackTracker.push({ sheet: orderSheet, startRow: startRow, numRows: newOrderHeaderRows.length });
      }

      if (newOrderDetailRows.length > 0) {
        const startRow = detailSheet.getLastRow() + 1;
        detailSheet.getRange(startRow, 1, newOrderDetailRows.length, newOrderDetailRows[0].length).setValues(newOrderDetailRows);
        rollbackTracker.push({ sheet: detailSheet, startRow: startRow, numRows: newOrderDetailRows.length });
      }

      // Commit toàn bộ thay đổi xuống Google Sheets
      SpreadsheetApp.flush();

      // --- BƯỚC 7 TRONG SPEC: DỌN DẸP SHEET 'TAO DON' (A8:I) ---
      this.clearWorkspaceSheetAfterSave_(payload && payload.sourceSpreadsheetId);

      // Xóa cache để lần tra cứu tiếp theo lấy dữ liệu mới nhất
      this.clearSystemCache();

      const orderCodesText = createdOrdersSummary.map(function(o) { return o.maDH; }).join(', ');
      const message = 'Đã lưu thành công ' + createdOrdersSummary.length + ' đơn hàng (' + orderCodesText + ') gồm ' +
        newOrderDetailRows.length + ' chi tiết mặt hàng!';

      return {
        success: true,
        orderCount: createdOrdersSummary.length,
        summaryText: orderCodesText,
        createdOrders: createdOrdersSummary,
        newCustomerCount: newCustomerRows.length,
        newProductCount: newProductRows.length,
        message: message,
        data: {
          orderCode: orderCodesText,
          lineCount: newOrderDetailRows.length
        }
      };

    } catch (txError) {
      // Rollback các dòng vừa thêm nếu giao dịch thất bại giữa chừng
      for (let r = rollbackTracker.length - 1; r >= 0; r--) {
        try {
          const item = rollbackTracker[r];
          item.sheet.deleteRows(item.startRow, item.numRows);
        } catch (rbErr) {
          console.error('[Rollback Error]:', rbErr);
        }
      }
      throw txError;
    } finally {
      // --- BƯỚC 8 TRONG SPEC: LUÔN GIẢI PHÓNG LOCK TRONG FINALLY ---
      lock.releaseLock();
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
   * Đồng bộ danh mục Khách hàng & Sản phẩm từ DB MASTER sang _CACHE_DANH_MUC
   * Cú pháp giữ nguyên chuẩn Object Literal bên trong OrderService
   */
  syncCacheCatalogFromDatabase: function(optTargetSpreadsheetId) {
    const tStart = Date.now();
    console.log("⏱️ [BACKEND-SYNC] Bắt đầu đồng bộ danh mục chuẩn...");

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

    // 2. Đọc Sản Phẩm và ghép ĐÚNG Khách Hàng sở hữu qua customer_id
    const prodSheet = masterSS.getSheetByName(CONFIG.SHEET_NAMES.PRODUCTS);
    const prodData = prodSheet.getDataRange().getValues();
    const cacheRows = [];
    for (let j = 1; j < prodData.length; j++) {
      const tenSP = String(prodData[j][5] || '').trim(); // name_napaco
      if (!tenSP) continue;

      const cusId = String(prodData[j][1] || '').trim(); // customer_id
      const tenKH = khMap[cusId] || cusId;              // Khách hàng sở hữu chuẩn xác

      cacheRows.push([
        tenSP,                             // Cột A: Tên Sản Phẩm
        tenKH,                             // Cột B: Khách Hàng sở hữu
        String(prodData[j][4] || '').trim(), // Cột C: Code SP
        String(prodData[j][8] || 'Cái').trim() || 'Cái', // Cột D: Đơn Vị
        Number(prodData[j][15]) || 0       // Cột E: Đơn Giá
      ]);
    }

    // 3. Ghi đè vào sheet _CACHE_DANH_MUC
    let cacheSheet = targetSS.getSheetByName(CONFIG.SHEET_NAMES.CACHE_CATALOG || '_CACHE_DANH_MUC');
    if (!cacheSheet) cacheSheet = targetSS.insertSheet(CONFIG.SHEET_NAMES.CACHE_CATALOG || '_CACHE_DANH_MUC');
    cacheSheet.clearContents();

    const headers = [['Tên Sản Phẩm', 'Khách Hàng', 'Code SP', 'Đơn Vị', 'Đơn Giá Chuẩn (₫)']];
    cacheSheet.getRange(1, 1, 1, 5).setValues(headers).setFontWeight('bold');
    if (cacheRows.length > 0) {
      cacheSheet.getRange(2, 1, cacheRows.length, 5).setValues(cacheRows);
    }

    // Cột K (11): Danh sách tất cả khách hàng (Dropdown cột Khách Hàng)
    cacheSheet.getRange(1, 11).setValue('Danh Sách Tất Cả KH').setFontWeight('bold');
    if (allCustomerNames.length > 0) {
      cacheSheet.getRange(2, 11, allCustomerNames.length, 1).setValues(allCustomerNames.map(c => [c]));
    }

    // 4. Gán Data Validation Dropdown cho sheet TAO DON theo tên cột động từ CONFIG
    const wsSheet = targetSS.getSheetByName(CONFIG.SHEET_NAMES.ORDER_WORKSPACE || 'TAO DON');
    if (wsSheet) {
      const headerRow = wsSheet.getRange(CONFIG.FORM_COLUMNS.HEADER_ROW, 1, 1, wsSheet.getLastColumn()).getValues()[0];
      const findCol = (aliases) => {
        for (let c = 0; c < headerRow.length; c++) {
          const text = String(headerRow[c] || '').trim().toLowerCase();
          if (aliases.some(a => a.toLowerCase() === text)) return c + 1;
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
    return {
      success: true,
      message: `Đã đồng bộ chính xác ${allCustomerNames.length} khách hàng và ${cacheRows.length} sản phẩm!`
    };
  },
};