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

    // 2. [ĐÃ SỬA] Đọc bảng Sản Phẩm tìm header động theo CONFIG.COLUMNS.PRODUCTS
    const prodSheet = masterSS.getSheetByName(CONFIG.SHEET_NAMES.PRODUCTS);
    if (!prodSheet) throw new Error('Không tìm thấy bảng ' + CONFIG.SHEET_NAMES.PRODUCTS);
    const prodValues = prodSheet.getDataRange().getValues();
    const products = [];
    const productByIdMap = {};

    // Tìm dòng header thực tế chứa SP_ID
    let prodHdrIdx = -1;
    for (let r = 0; r < Math.min(prodValues.length, 5); r++) {
      const rowStr = prodValues[r].map(c => String(c || '').trim().toLowerCase());
      if (rowStr.indexOf(CONFIG.COLUMNS.PRODUCTS.ID.toLowerCase()) !== -1) {
        prodHdrIdx = r;
        break;
      }
    }
    if (prodHdrIdx === -1) prodHdrIdx = 0;

    const pHdrs = prodValues[prodHdrIdx].map(h => String(h || '').trim().toLowerCase());
    const cSpId = pHdrs.indexOf(CONFIG.COLUMNS.PRODUCTS.ID.toLowerCase());
    const cCusId = pHdrs.indexOf(CONFIG.COLUMNS.PRODUCTS.CUSTOMER_ID.toLowerCase());
    const cFilm = pHdrs.indexOf(CONFIG.COLUMNS.PRODUCTS.FILM.toLowerCase());
    const cType = pHdrs.indexOf(CONFIG.COLUMNS.PRODUCTS.TYPE.toLowerCase());
    const cCodeCust = pHdrs.indexOf(CONFIG.COLUMNS.PRODUCTS.CODE_CUST.toLowerCase());
    const cName = pHdrs.indexOf(CONFIG.COLUMNS.PRODUCTS.NAME.toLowerCase());
    const cCavity = pHdrs.indexOf(CONFIG.COLUMNS.PRODUCTS.CAVITY.toLowerCase());
    const cStock = pHdrs.indexOf(CONFIG.COLUMNS.PRODUCTS.STOCK.toLowerCase());
    const cUnit = pHdrs.indexOf(CONFIG.COLUMNS.PRODUCTS.UNIT.toLowerCase());
    const cImage = pHdrs.indexOf(CONFIG.COLUMNS.PRODUCTS.IMAGE.toLowerCase());
    const cPacking = pHdrs.indexOf(CONFIG.COLUMNS.PRODUCTS.PACKING.toLowerCase());
    const cMoldPrice = pHdrs.indexOf(CONFIG.COLUMNS.PRODUCTS.MOLD_PRICE.toLowerCase());
    const cPrice = pHdrs.indexOf(CONFIG.COLUMNS.PRODUCTS.PRICE.toLowerCase());

    console.log('[DEBUG_LOG_SANPHAM] Header row index:', prodHdrIdx, '| Cột SP_ID:', cSpId, '| Cột FILM (Màng):', cFilm);

    for (let j = prodHdrIdx + 1; j < prodValues.length; j++) {
      const row = prodValues[j];
      const spId = cSpId !== -1 ? String(row[cSpId] || '').trim() : '';
      const nameNapaco = cName !== -1 ? String(row[cName] || '').trim() : '';
      if (!spId || !nameNapaco) continue;

      const mangVal = cFilm !== -1 ? String(row[cFilm] || '').trim() : '';
      const prodObj = {
        spId: spId,
        customerId: cCusId !== -1 ? String(row[cCusId] || '').trim() : '',
        mang: mangVal,
        type: cType !== -1 ? String(row[cType] || '').trim() : '',
        codeSP: cCodeCust !== -1 ? String(row[cCodeCust] || '').trim() : '',
        prodName: nameNapaco,
        cavity: cCavity !== -1 ? (Number(row[cCavity]) || 0) : 0,
        tonkho: cStock !== -1 ? (Number(row[cStock]) || 0) : 0,
        unit: cUnit !== -1 ? (String(row[cUnit] || 'Cái').trim() || 'Cái') : 'Cái',
        imageURL: cImage !== -1 ? String(row[cImage] || '').trim() : '',
        donggoi: cPacking !== -1 ? String(row[cPacking] || '').trim() : '',
        giaKhuon: cMoldPrice !== -1 ? (Number(row[cMoldPrice]) || 0) : 0,
        price: cPrice !== -1 ? (Number(row[cPrice]) || 0) : 0
      };
      products.push(prodObj);
      productByIdMap[spId] = prodObj;
      productByIdMap[spId.toUpperCase()] = prodObj;
    }

    // 3. [ĐÃ SỬA] Đọc bảng Nhân Viên tìm header động theo CONFIG.COLUMNS.EMPLOYEES
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

    let empHdrIdx = -1;
    for (let r = 0; r < Math.min(empValues.length, 6); r++) {
      const rowStr = empValues[r].map(c => String(c || '').trim().toLowerCase());
      if (rowStr.indexOf(CONFIG.COLUMNS.EMPLOYEES.ID.toLowerCase()) !== -1) {
        empHdrIdx = r;
        break;
      }
    }
    if (empHdrIdx === -1) empHdrIdx = 0;

    const eHdrs = empValues[empHdrIdx].map(h => String(h || '').trim().toLowerCase());
    const cEmpId = eHdrs.indexOf(CONFIG.COLUMNS.EMPLOYEES.ID.toLowerCase());
    const cEmpName = eHdrs.indexOf(CONFIG.COLUMNS.EMPLOYEES.NAME.toLowerCase());
    const cEmpDept = eHdrs.indexOf(CONFIG.COLUMNS.EMPLOYEES.DEPARTMENT.toLowerCase());
    const cEmpRole = eHdrs.indexOf(CONFIG.COLUMNS.EMPLOYEES.ROLE.toLowerCase());
    const cEmpEmail = eHdrs.indexOf(CONFIG.COLUMNS.EMPLOYEES.EMAIL.toLowerCase());
    const cEmpPhone = eHdrs.indexOf(CONFIG.COLUMNS.EMPLOYEES.PHONE.toLowerCase());
    const cEmpAccess = eHdrs.indexOf(CONFIG.COLUMNS.EMPLOYEES.ACCESS.toLowerCase());
    const cEmpStatus = eHdrs.indexOf(CONFIG.COLUMNS.EMPLOYEES.STATUS.toLowerCase());

    console.log('[DEBUG_LOG_NHANVIEN] Header row index:', empHdrIdx, '| Cột Ma_NV:', cEmpId, '| Cột Name:', cEmpName);

    for (let k = empHdrIdx + 1; k < empValues.length; k++) {
      const row = empValues[k];
      const maNV = cEmpId !== -1 ? String(row[cEmpId] || '').trim() : '';
      const empName = cEmpName !== -1 ? String(row[cEmpName] || '').trim() : '';
      if (!maNV || !/^NV/i.test(maNV) || !empName) continue;

      const phongBan = cEmpDept !== -1 ? String(row[cEmpDept] || '').trim() : '';
      const role = cEmpRole !== -1 ? String(row[cEmpRole] || '').trim() : '';
      const email = cEmpEmail !== -1 ? String(row[cEmpEmail] || '').trim().toLowerCase() : '';
      const sdt = cEmpPhone !== -1 ? String(row[cEmpPhone] || '').trim() : '';
      const quyenTruyCap = cEmpAccess !== -1 ? String(row[cEmpAccess] || '').trim() : '';
      const trangThai = cEmpStatus !== -1 ? String(row[cEmpStatus] || '').trim() : '';

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

      // Ánh xạ đa quy tắc: hỗ trợ NV0001, NV-KD01, NV-KD1...
      employeeByIdMap[maNV] = empObj;
      employeeByIdMap[maNV.toUpperCase()] = empObj;
      const numMatch = maNV.match(/\d+/);
      if (numMatch) {
        const num = parseInt(numMatch[0], 10);
        employeeByIdMap['NV-KD' + String(num).padStart(2, '0')] = empObj;
        employeeByIdMap['NV-KD' + num] = empObj;
        employeeByIdMap['NV' + String(num).padStart(4, '0')] = empObj;
      }

      // [LOẠI BỎ HARDCODE] Dùng CONFIG.EMPLOYEE_STATUS.ACTIVE & CONFIG.DEPARTMENTS.SALES
      const isActive = trangThai.toLowerCase().includes(CONFIG.EMPLOYEE_STATUS.ACTIVE.toLowerCase()) || trangThai.toLowerCase().includes('hoạt động');
      if (isActive) {
        allActiveStaffList.push(empObj);
        const isSales = phongBan.toLowerCase().includes(CONFIG.DEPARTMENTS.SALES.toLowerCase());
        const isManager = quyenTruyCap.includes('Admin') || quyenTruyCap.includes('Manager');
        if (isSales || isManager) {
          salesStaffList.push(empObj);
        }
        if (sessionEmail && email && sessionEmail === email) {
          matchedEmployee = empObj;
        }
      }
    }
    console.log('[DEBUG_LOG_NHANVIEN] Số nhân viên KD tải được:', salesStaffList.length, '| Danh sách:', salesStaffList.map(s => s.display));

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
      const prodRowIndexMap = {};
      for (let j = 1; j < prodData.length; j++) {
        const spId = String(prodData[j][0] || '').trim();
        const cusId = String(prodData[j][1] || '').trim();
        const spName = String(prodData[j][5] || '').trim().toLowerCase();
        if (spId && spName) {
          prodMap[spName + '__' + cusId] = spId;
          if (!prodMap[spName]) prodMap[spName] = spId;
          prodRowIndexMap[spId] = j + 1; // j bắt đầu từ 1 (hàng 2 trên Sheet), nên dòng thực tế là j + 1
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
          } else {
            // Sản phẩm đã tồn tại -> Cập nhật lại thông số mới nhất vào bảng 02_SAN_PHAM
            const spRowIdx = prodRowIndexMap[targetSpId];
            if (spRowIdx) {
              // Cột 3 (mang), Cột 4 (type), Cột 5 (Code_Sanpham), Cột 9 (donvi), Cột 14 (donggoi), Cột 15 (Gia Khuon), Cột 16 (dongia)
              if (item.mang) prodSheet.getRange(spRowIdx, 3).setValue(item.mang);
              if (item.type) prodSheet.getRange(spRowIdx, 4).setValue(item.type);
              if (item.codeSP) prodSheet.getRange(spRowIdx, 5).setValue(item.codeSP);
              if (item.unit) prodSheet.getRange(spRowIdx, 9).setValue(item.unit);
              if (item.donggoi) prodSheet.getRange(spRowIdx, 14).setValue(item.donggoi);
              if (Number(item.giaKhuon) > 0) prodSheet.getRange(spRowIdx, 15).setValue(Number(item.giaKhuon));
              if (Number(item.price) > 0) prodSheet.getRange(spRowIdx, 16).setValue(Number(item.price));
            }
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
    /* [SỬA LỖI] Bỏ toLowerCase() để giữ nguyên hằng số nhóm ACTIVE_ALL / DONE_ALL */
    const statusFilter = String(params.status || 'ALL').trim();
    const fromDate = String(params.fromDate || '').trim();
    const toDate = String(params.toDate || '').trim();
    const page = Math.max(1, parseInt(params.page, 10) || 1);
    const pageSize = Math.max(1, parseInt(params.pageSize, 10) || 100);

    const catalog = this.getCatalogAndAuthData();
    const opSS = this.getOperationSpreadsheet_();
    const orderSheet = opSS.getSheetByName(CONFIG.SHEET_NAMES.ORDERS);
    const detailSheet = opSS.getSheetByName(CONFIG.SHEET_NAMES.ORDER_DETAILS);

        // =========================================================================
// [DEBUG SERVER - BƯỚC 1]: ĐẶT TẠI ĐẦU HÀM queryOrdersAndLineItems
// =========================================================================
console.log("=== [DEBUG SERVER] BẮT ĐẦU TRUY VẤN ĐƠN HÀNG ===");
console.log("1. Tham số nhận từ Client:", JSON.stringify({
  viewMode: viewMode,
  keyword: keyword,
  maKHFilter: maKHFilter,
  statusFilter: statusFilter,
  fromDate: fromDate,
  toDate: toDate
}));

    if (!orderSheet || !detailSheet) {
      throw new Error('Không tìm thấy bảng 01.DonHang hoặc 02.ChiTietDH!');
    }

    const orderRows = orderSheet.getDataRange().getValues();
    const detailRows = detailSheet.getDataRange().getValues();

    const moldMap = this.getMoldsMap_();

    // =========================================================================
    // 1. [ĐƯA LÊN TRƯỚC] ÁNH XẠ HEADER ĐƠN HÀNG (01.DonHang) TẠO orderMap TRƯỚC
    // =========================================================================
    const oHdrs = orderRows[0].map(h => String(h || '').trim().toLowerCase());
    const colOMaDH = oHdrs.indexOf(CONFIG.COLUMNS.ORDERS.ID.toLowerCase());
    const colOMaKH = oHdrs.indexOf(CONFIG.COLUMNS.ORDERS.CUSTOMER_ID.toLowerCase());
    const colONgayDat = oHdrs.indexOf(CONFIG.COLUMNS.ORDERS.ORDER_DATE.toLowerCase());
    const colONgayGiao = oHdrs.indexOf(CONFIG.COLUMNS.ORDERS.DELIVERY_DATE.toLowerCase());
    const colOSoPO = oHdrs.indexOf(CONFIG.COLUMNS.ORDERS.PO_NUMBER.toLowerCase());
    const colOFilePO = oHdrs.indexOf(CONFIG.COLUMNS.ORDERS.FILE_URL.toLowerCase());

    const orderMap = {};
    const orderHeadersList = [];

    for (let i = 1; i < orderRows.length; i++) {
      const oRow = orderRows[i];
      const maDH = colOMaDH !== -1 ? String(oRow[colOMaDH] || '').trim() : '';
      if (!maDH) continue;

      const maKH = colOMaKH !== -1 ? String(oRow[colOMaKH] || '').trim() : '';
      const ngayDat = colONgayDat !== -1 ? this.formatSheetDate_(oRow[colONgayDat]) : '';
      const ngayGiaoDK = colONgayGiao !== -1 ? this.formatSheetDate_(oRow[colONgayGiao]) : '';
      const soPO = colOSoPO !== -1 ? String(oRow[colOSoPO] || '').trim() : '';
      const filePO = colOFilePO !== -1 ? String(oRow[colOFilePO] || '').trim() : '';

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

    // =========================================================================
    // 2. ÁNH XẠ DÒNG CHI TIẾT MẶT HÀNG (02.ChiTietDH) - DÙNG orderMap AN TOÀN
    // =========================================================================
    const dHdrs = detailRows[0].map(h => String(h || '').trim().toLowerCase());
    const colDId = dHdrs.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.ID.toLowerCase());
    const colDOrderId = dHdrs.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.ORDER_ID.toLowerCase());
    const colDSpId = dHdrs.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.PRODUCT_ID.toLowerCase());
    const colDQty = dHdrs.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.QTY_ORDER.toLowerCase());
    const colDPrice = dHdrs.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.PRICE.toLowerCase());
    const colDQtyRun = dHdrs.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.QTY_RUN.toLowerCase());
    const colDQtyDel = dHdrs.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.QTY_DELIVERED.toLowerCase());
    const colDQtyRem = dHdrs.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.QTY_REMAINING.toLowerCase());
    const colDNote = dHdrs.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.NOTE.toLowerCase());
    const colDDate = dHdrs.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.DELIVERY_DATE.toLowerCase());
    const colDStaff = dHdrs.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.STAFF_ID.toLowerCase());
    const colDStatus = dHdrs.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.STATUS.toLowerCase());

    const allLineItems = [];
    for (let j = 1; j < detailRows.length; j++) {
      const row = detailRows[j];

      // =========================================================================
// [DEBUG SERVER - BƯỚC 2]: ĐẶT NGAY TRƯỚC VÀ SAU VÙNG LỌC BẢNG ĐƠN HÀNG (orderRows / detailRows)
// =========================================================================
console.log("2. Tổng số dòng đọc được từ Sheet:", {
  tongDonHangSheet: orderRows.length - 1,
  tongChiTietSheet: detailRows.length - 1
});


      const idChiTiet = colDId !== -1 ? String(row[colDId] || '').trim() : '';
      const idDonHang = colDOrderId !== -1 ? String(row[colDOrderId] || '').trim() : '';
      if (!idChiTiet || !idDonHang) continue;

      const maSP = colDSpId !== -1 ? String(row[colDSpId] || '').trim() : '';
      const slDat = colDQty !== -1 ? (Number(row[colDQty]) || 0) : 0;
      const donGia = colDPrice !== -1 ? (Number(row[colDPrice]) || 0) : 0;
      const slDaChay = colDQtyRun !== -1 ? (Number(row[colDQtyRun]) || 0) : 0;
      const slDaGiao = colDQtyDel !== -1 ? (Number(row[colDQtyDel]) || 0) : 0;
      const slConLaiRaw = colDQtyRem !== -1 ? row[colDQtyRem] : '';
      const slConLai = (slConLaiRaw !== '' && slConLaiRaw !== null && slConLaiRaw !== undefined)
        ? Number(slConLaiRaw)
        : Math.max(0, slDat - slDaGiao);
      const ghiChu = colDNote !== -1 ? String(row[colDNote] || '').trim() : '';
      const ngayGiao = colDDate !== -1 ? this.formatSheetDate_(row[colDDate]) : '';
      const maNV = colDStaff !== -1 ? String(row[colDStaff] || '').trim() : '';
      const trangThaiRaw = colDStatus !== -1 ? String(row[colDStatus] || CONFIG.LINE_STATUS.NEW).trim() : CONFIG.LINE_STATUS.NEW;

      const parentOrder = orderMap[idDonHang];
      const prodObj = catalog.productByIdMap[maSP] || catalog.productByIdMap[maSP.toUpperCase()];
      const tenSalesVal = this.resolveStaffName_(maNV, catalog.employeeByIdMap, catalog.salesStaffList);

      const sanPhamName = prodObj ? prodObj.prodName : maSP;
      const codeSP = prodObj ? prodObj.codeSP : '';
      const dvt = prodObj ? prodObj.unit : 'Cái';
      const mangVal = prodObj && prodObj.mang ? prodObj.mang : '---';
      const moldInfo = moldMap[maSP] || { dinhHinh: [], dao: [] };

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
        thanhTien: slDat * donGia,
        daChay: slDaChay,
        daGiao: slDaGiao,
        conLai: slConLai,
        ghiChu: ghiChu,
        ngayGiao: ngayGiao || (parentOrder ? parentOrder.ngayGiaoDK : ''),
        maNV: maNV,
        tenNV: tenSalesVal,
        trangThai: trangThaiRaw,
        mang: mangVal,
        moldInfo: moldInfo,
        imageURL: prodObj ? prodObj.imageURL : '',
        giaKhuon: prodObj ? prodObj.giaKhuon : 0
      };

      allLineItems.push(lineObj);

      if (parentOrder) {
        parentOrder.items.push(lineObj);
        parentOrder.itemCount++;
        parentOrder.tongSLDat += slDat;
        parentOrder.tongDaGiao += slDaGiao;
        parentOrder.tongConLai += slConLai;
        parentOrder.tongGiaTri += (slDat * donGia);
        if (!parentOrder.salesPerson && lineObj.tenNV && lineObj.tenNV !== '---') {
          parentOrder.salesPerson = lineObj.tenNV;
        }
      }
    }

    // Giả sử mảng kết quả sau khi duyệt và lọc là: orderHeadersList (hoặc lineItemsList)
console.log("3. Số bản ghi KHỚP điều kiện lọc chuẩn bị trả về Client:", {
  soLuongKetQua: (viewMode === 'HEADER' ? orderHeadersList.length : (typeof lineItemsList !== 'undefined' ? lineItemsList.length : 0))
});

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

    /* [MỚI - Ý 5 & RÀ SOÁT HARDCODE] Hàm helper kiểm tra khớp trạng thái theo phân nhóm CONFIG */
    /* [SỬA LỖI] Hàm helper kiểm tra khớp trạng thái bao quát case-sensitive */
    const isStatusMatch = function(itemStatus, filterVal) {
      const fValUpper = String(filterVal).trim().toUpperCase();
      if (!filterVal || fValUpper === 'ALL') return true;
      
      const sLow = String(itemStatus || '').toLowerCase().trim();

      // Nhóm 1: Tất cả đơn đang chạy (Loại trừ các trạng thái đã đóng/hủy)
      if (fValUpper === CONFIG.FILTER_STATUS_GROUPS.ACTIVE_ALL) {
        const isDone = sLow.includes(CONFIG.LINE_STATUS.COMPLETED.toLowerCase()) || 
                       sLow.includes(CONFIG.LINE_STATUS.CLOSED.toLowerCase()) ||
                       sLow.includes(CONFIG.ORDER_STATUS.CANCELLED.toLowerCase()) ||
                       sLow.includes('hoàn tất');
        return !isDone;
      }
      
      // Nhóm 2: Tất cả đơn đã xong/kết thúc
      if (fValUpper === CONFIG.FILTER_STATUS_GROUPS.DONE_ALL) {
        return sLow.includes(CONFIG.LINE_STATUS.COMPLETED.toLowerCase()) || 
               sLow.includes(CONFIG.LINE_STATUS.CLOSED.toLowerCase()) ||
               sLow.includes('hoàn tất');
      }
      
      // Trạng thái cụ thể: ép về lowercase để tìm kiếm tương đối (ví dụ tìm: "chờ sản xuất")
      return sLow.includes(String(filterVal).trim().toLowerCase());
    };

    // 4. Lọc theo chế độ xem (HEADER hoặc LINE_ITEM)
    if (viewMode === 'LINE_ITEM') {
      const filteredLines = allLineItems.filter(function(item) {
        if (maKHFilter && maKHFilter !== 'ALL' && item.maKH !== maKHFilter && item.khachHang !== maKHFilter) {
          return false;
        }
        /* [MỚI - Ý 5] Lọc trạng thái theo điều kiện HOẶC chuẩn xác */
        if (statusFilter && statusFilter !== 'ALL' && statusFilter !== 'all') {
          const sLow = item.trangThai.toLowerCase().trim();
          if (statusFilter === CONFIG.FILTER_STATUS_GROUPS.ACTIVE_ALL) {
            // Điều kiện HOẶC: Chờ sản xuất HOẶC Đang làm khuôn HOẶC Đang sản xuất
            const isChoSX = sLow.includes(CONFIG.LINE_STATUS.NEW.toLowerCase());
            const isKhuon = sLow.includes(CONFIG.LINE_STATUS.MOLDING.toLowerCase()) || sLow.includes('khuôn');
            const isDangSX = sLow.includes(CONFIG.LINE_STATUS.IN_PRODUCTION.toLowerCase()) && !sLow.includes('chờ');
            if (!isChoSX && !isKhuon && !isDangSX) return false;
          } else if (statusFilter === CONFIG.FILTER_STATUS_GROUPS.DONE_ALL) {
            // Điều kiện HOẶC: Xong HOẶC Kết thúc HOẶC Hoàn tất
            const isXong = sLow.includes(CONFIG.LINE_STATUS.COMPLETED.toLowerCase());
            const isKetThuc = sLow.includes(CONFIG.LINE_STATUS.CLOSED.toLowerCase());
            const isHoanTat = sLow.includes(CONFIG.ORDER_STATUS.COMPLETED.toLowerCase());
            if (!isXong && !isKetThuc && !isHoanTat) return false;
          } else {
            if (sLow.indexOf(statusFilter.toLowerCase()) === -1) return false;
          }
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

      filteredLines.reverse();
      const total = filteredLines.length;

      /* [MỚI - Ý 5] Kiểm tra giới hạn bản ghi tối đa theo CONFIG */
      const maxLimit = CONFIG.MAX_DISPLAY_RECORDS || 100;
      if (total > maxLimit) {
        return {
          success: false,
          exceededLimit: true,
          total: total,
          maxLimit: maxLimit,
          message: 'Dữ liệu tìm kiếm vượt quá số lượng cho phép (' + total + ' / ' + maxLimit + ' dòng). Vui lòng nhập thêm từ khóa hoặc điều kiện lọc để thu hẹp phạm vi tìm kiếm!'
        };
      }

      const startIndex = (page - 1) * pageSize;
      const pagedData = filteredLines.slice(startIndex, startIndex + pageSize);

      return {
        success: true,
        viewMode: 'LINE_ITEM',
        total: total,
        summary: this.computeQueryMetrics_(orderHeadersList, filteredLines),
        customers: catalog.customers,
        salesStaffList: catalog.salesStaffList,
        data: pagedData
      };
    } else {
      const filteredHeaders = orderHeadersList.filter(function(ord) {
        if (maKHFilter && maKHFilter !== 'ALL' && ord.maKH !== maKHFilter && ord.khachHang !== maKHFilter) {
          return false;
        }
        /* [MỚI - Ý 5] Lọc trạng thái theo phân nhóm cho Header */
        if (!isStatusMatch(ord.trangThai, statusFilter)) {
          return false;
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

      /* [MỚI - Ý 5] Kiểm tra giới hạn bản ghi tối đa theo CONFIG cho Header */
      const maxLimit = CONFIG.MAX_DISPLAY_RECORDS || 100;
      if (total > maxLimit) {
        return {
          success: false,
          exceededLimit: true,
          total: total,
          maxLimit: maxLimit,
          message: 'Dữ liệu tìm kiếm vượt quá số lượng cho phép (' + total + ' / ' + maxLimit + ' đơn). Vui lòng nhập thêm từ khóa hoặc điều kiện lọc để thu hẹp phạm vi tìm kiếm!'
        };
      }

      const startIndex = (page - 1) * pageSize;
      const pagedData = filteredHeaders.slice(startIndex, startIndex + pageSize);

      return {
        success: true,
        viewMode: 'HEADER',
        total: total,
        summary: this.computeQueryMetrics_(filteredHeaders, allLineItems),
        customers: catalog.customers,
        salesStaffList: catalog.salesStaffList,
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
      if (maNV && /^NV/i.test(maNV) && rowEmail === email && trangThai.toLowerCase().includes(CONFIG.EMPLOYEE_STATUS.ACTIVE.toLowerCase())) {
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

        // Thay thế đoạn:
        // if (maNV.toUpperCase().startsWith('NV') && tt.includes('hoạt động') && pb.includes('kinh doanh'))
        // Thành:
        const isAct = tt.includes(CONFIG.EMPLOYEE_STATUS.ACTIVE.toLowerCase());
        const isSal = pb.includes(CONFIG.DEPARTMENTS.SALES.toLowerCase());
        if (maNV.toUpperCase().startsWith('NV') && isAct && isSal) {
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

  /**
   * [MỚI - SPEC-OMS-LINEITEM-202611-01] Đọc bảng 05.Cong_Cu và tạo map Khuôn định hình theo SP_ID
   * @private
   */
  getMoldsMap_: function() {
    try {
      const masterSS = this.getMasterSpreadsheet_();
      const sheetName = (CONFIG.SHEET_NAMES && CONFIG.SHEET_NAMES.TOOLS) ? CONFIG.SHEET_NAMES.TOOLS : '05.Cong_Cu';
      const toolSheet = masterSS.getSheetByName(sheetName);
      if (!toolSheet) return {};

      const data = toolSheet.getDataRange().getValues();
      if (data.length <= 1) return {};

      const headers = data[0].map(h => String(h || '').trim().toLowerCase());
      const colSpId = headers.indexOf(CONFIG.COLUMNS.TOOLS.SP_ID.toLowerCase());
      const colType = headers.indexOf(CONFIG.COLUMNS.TOOLS.TOOL_TYPE.toLowerCase());
      const colMold = headers.indexOf(CONFIG.COLUMNS.TOOLS.MOLD_CODE.toLowerCase());
      const colStatus = headers.indexOf(CONFIG.COLUMNS.TOOLS.STATUS.toLowerCase());

      if (colSpId === -1 || colMold === -1) return {};

      const moldMap = {};
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const spId = String(row[colSpId] || '').trim();
        const typeRaw = colType !== -1 ? String(row[colType] || '').trim().toLowerCase() : '';
        const moldCode = String(row[colMold] || '').trim();
        const statusRaw = colStatus !== -1 ? String(row[colStatus] || '').trim() : '';

        if (!spId || !moldCode) continue;
        if (!moldMap[spId]) {
          moldMap[spId] = { dinhHinh: [], dao: [] };
        }

        // [LOẠI BỎ HARDCODE] Phân loại khuôn định hình và khuôn dao qua CONFIG.MOLD_TYPES
        const isForming = typeRaw.includes(CONFIG.MOLD_TYPES.FORMING.toLowerCase()) || typeRaw.includes('dinhhinh') || typeRaw.includes('định hình') || !typeRaw;
        const isCutting = typeRaw.includes(CONFIG.MOLD_TYPES.CUTTING.toLowerCase()) || typeRaw.includes('dao');

        if (isForming) {
          moldMap[spId].dinhHinh.push({ code: moldCode, status: statusRaw });
        } else if (isCutting) {
          moldMap[spId].dao.push({ code: moldCode, status: statusRaw });
        }
      }
      return moldMap;
    } catch (e) {
      console.warn('[OrderService.getMoldsMap_] Cảnh báo đọc khuôn:', e.message);
      return {};
    }
  },

  /**
   * [MỚI - SPEC-OMS-LINEITEM-202611-01] Chỉnh sửa hàng loạt (Bulk Edit) với LockService và Batch Read/Write
   * @param {Object} payload { targetIds: string[], fieldsToUpdate: Object }
   */
  updateOrderLineItemsBulk: function(payload) {
    /* [MỚI - VẤN ĐỀ 5] Hỗ trợ đồng thời 2 cấu trúc payload: mảng itemsToUpdate (riêng từng SP) hoặc targetIds chung */
    let updateMap = {}; // idChiTiet -> fieldsToUpdate
    if (payload && Array.isArray(payload.itemsToUpdate) && payload.itemsToUpdate.length > 0) {
      payload.itemsToUpdate.forEach(item => {
        if (item.idChiTiet) updateMap[item.idChiTiet] = item.fields || {};
      });
    } else if (payload && Array.isArray(payload.targetIds) && payload.targetIds.length > 0) {
      const commonFields = payload.fieldsToUpdate || {};
      payload.targetIds.forEach(id => {
        updateMap[id] = commonFields;
      });
    } else {
      return { success: false, message: 'Danh sách dòng mặt hàng cần cập nhật không hợp lệ.' };
    }

    const targetIds = Object.keys(updateMap);
    if (targetIds.length === 0) {
      return { success: false, message: 'Không có dữ liệu mặt hàng nào được chọn.' };
    }

    // 1. KIỂM TRA PHÂN QUYỀN SERVER-SIDE
    const catalog = this.getCatalogAndAuthData();
    const currentUser = catalog.currentUser;
    const allowedRoles = CONFIG.ALLOWED_ROLES_BULK_EDIT || ['Kinh doanh', 'QLSX', 'QC', 'Kho'];

    let hasPermission = false;
    if (currentUser && currentUser.isAuthenticated) {
      const uRole = String(currentUser.role || '').toLowerCase();
      const uDept = String(currentUser.phongBan || '').toLowerCase();
      const uAccess = String(currentUser.quyenTruyCap || '').toLowerCase();
      if (uAccess.includes('admin') || uAccess.includes('manager')) {
        hasPermission = true;
      } else {
        hasPermission = allowedRoles.some(r => {
          const rLow = r.toLowerCase();
          return uRole.includes(rLow) || uDept.includes(rLow);
        });
      }
    } else {
      try {
        if (Session.getActiveUser().getEmail()) hasPermission = true;
      } catch (e) {
        hasPermission = false;
      }
    }

    if (!hasPermission) {
      return { success: false, message: 'Bạn không có quyền thực hiện thao tác này.' };
    }

    // 2. KHÓA HỆ THỐNG CHỐNG RACE CONDITION (TIMEOUT 30s)
    const lock = LockService.getScriptLock();
    let hasLock = false;
    try {
      hasLock = lock.tryLock(30000);
    } catch (e) {
      hasLock = false;
    }

    if (!hasLock) {
      return { success: false, message: 'Hệ thống đang bận xử lý giao dịch khác. Vui lòng thử lại sau 30 giây.' };
    }

    try {
      const operSS = this.getOperationSpreadsheet_();
      const masterSS = this.getMasterSpreadsheet_();

      const detailSheet = operSS.getSheetByName(CONFIG.SHEET_NAMES.ORDER_DETAILS);
      const orderSheet = operSS.getSheetByName(CONFIG.SHEET_NAMES.ORDERS);
      const prodSheet = masterSS.getSheetByName(CONFIG.SHEET_NAMES.PRODUCTS);

      if (!detailSheet || !orderSheet || !prodSheet) {
        throw new Error('Không tìm thấy các sheet cơ sở dữ liệu cần cập nhật.');
      }

      // Đọc toàn bộ bảng 02.ChiTietDH (Batch Read)
      const detailData = detailSheet.getDataRange().getValues();
      const dHeaders = detailData[0].map(h => String(h || '').trim());
      
      const colIdChiTiet = dHeaders.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.ID);
      const colIdDH = dHeaders.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.ORDER_ID);
      const colMaSP = dHeaders.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.PRODUCT_ID);
      const colSlDat = dHeaders.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.QTY_ORDER);
      const colDonGia = dHeaders.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.PRICE);
      const colDaGiao = dHeaders.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.QTY_DELIVERED);
      const colConLai = dHeaders.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.QTY_REMAINING);
      const colNgayGiao = dHeaders.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.DELIVERY_DATE);

      const targetIdSet = new Set(targetIds);
      const affectedOrdersMap = {}; // maDH -> ngayDat mới
      const affectedProductsMap = {}; // maSP -> { mang, giaKhuon, donvi }
      let updatedCount = 0;

      // 3. CẬP NHẬT TRONG BỘ NHỚ BẢNG 02.ChiTietDH
      for (let r = 1; r < detailData.length; r++) {
        const rowId = String(detailData[r][colIdChiTiet] || '').trim();
        if (targetIdSet.has(rowId)) {
          const fields = updateMap[rowId] || {};
          updatedCount++;

          const currentDH = String(detailData[r][colIdDH] || '').trim();
          const currentSP = String(detailData[r][colMaSP] || '').trim();

          if (currentDH && fields.ngayDat) {
            affectedOrdersMap[currentDH] = fields.ngayDat;
          }
          if (currentSP) {
            if (!affectedProductsMap[currentSP]) affectedProductsMap[currentSP] = {};
            if (fields.mang !== undefined && fields.mang !== '') affectedProductsMap[currentSP].mang = fields.mang;
            if (fields.giaKhuon !== undefined && fields.giaKhuon !== '') affectedProductsMap[currentSP].giaKhuon = fields.giaKhuon;
            if (fields.dvt !== undefined && fields.dvt !== '') affectedProductsMap[currentSP].donvi = fields.dvt;
            /* [MỚI - Ý 1] Cập nhật luôn Đơn giá mới vào Master 01_02_SAN_PHAM */
            if (fields.donGia !== undefined && fields.donGia !== null && fields.donGia !== '') {
              affectedProductsMap[currentSP].dongia = Number(fields.donGia) || 0;
            }
          }

          // Cập nhật Số lượng đặt & Tính lại SL_Con_Lai
          if (fields.soLuongDat !== undefined && fields.soLuongDat !== null && fields.soLuongDat !== '') {
            const newSlDat = Number(fields.soLuongDat) || 0;
            detailData[r][colSlDat] = newSlDat;
            const currentDaGiao = Number(detailData[r][colDaGiao]) || 0;
            detailData[r][colConLai] = Math.max(0, newSlDat - currentDaGiao);
          }

          // Cập nhật Đơn giá
          if (fields.donGia !== undefined && fields.donGia !== null && fields.donGia !== '') {
            detailData[r][colDonGia] = Number(fields.donGia) || 0;
          }

          // Cập nhật Ngày giao
          if (fields.ngayGiao !== undefined && fields.ngayGiao !== null && fields.ngayGiao !== '') {
            detailData[r][colNgayGiao] = this.formatSheetDate_(fields.ngayGiao);
          }

          /* [MỚI - Ý 3] Cập nhật Sales (Mã NV) */
          const colStaff = dHeaders.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.STAFF_ID);
          if (colStaff !== -1 && fields.maNV !== undefined && fields.maNV !== null && fields.maNV !== '') {
            detailData[r][colStaff] = String(fields.maNV).trim();
          }

          /* [MỚI - Ý 3] Cập nhật Trạng Thái dòng mặt hàng */
          const colStatus = dHeaders.indexOf(CONFIG.COLUMNS.ORDER_DETAILS.STATUS);
          if (colStatus !== -1 && fields.trangThai !== undefined && fields.trangThai !== null && fields.trangThai !== '') {
            detailData[r][colStatus] = String(fields.trangThai).trim();
          }
        }
      }

      // Batch Write bảng 02.ChiTietDH
      if (updatedCount > 0) {
        detailSheet.getRange(1, 1, detailData.length, detailData[0].length).setValues(detailData);
      }

      // 4. CẬP NHẬT HEADER BẢNG 01.DonHang
      const affectedDHKeys = Object.keys(affectedOrdersMap);
      if (affectedDHKeys.length > 0) {
        const orderData = orderSheet.getDataRange().getValues();
        const oHeaders = orderData[0].map(h => String(h || '').trim());
        const colMaDH = oHeaders.indexOf(CONFIG.COLUMNS.ORDERS.ID);
        const colNgayDat = oHeaders.indexOf(CONFIG.COLUMNS.ORDERS.ORDER_DATE);

        if (colMaDH !== -1 && colNgayDat !== -1) {
          let orderChanged = false;
          for (let ro = 1; ro < orderData.length; ro++) {
            const oId = String(orderData[ro][colMaDH] || '').trim();
            if (affectedOrdersMap[oId]) {
              orderData[ro][colNgayDat] = this.formatSheetDate_(affectedOrdersMap[oId]);
              orderChanged = true;
            }
          }
          if (orderChanged) {
            orderSheet.getRange(1, 1, orderData.length, orderData[0].length).setValues(orderData);
          }
        }
      }

      // 5. CẬP NHẬT MASTER 01_02_SAN_PHAM (Màng, Giá khuôn, Đơn vị tính)
      const affectedSPKeys = Object.keys(affectedProductsMap);
      if (affectedSPKeys.length > 0) {
        const prodData = prodSheet.getDataRange().getValues();
        const pHeaders = prodData[0].map(h => String(h || '').trim());
        const colSpId = pHeaders.indexOf(CONFIG.COLUMNS.PRODUCTS.ID);
        const colFilm = pHeaders.indexOf(CONFIG.COLUMNS.PRODUCTS.FILM);
        const colMoldPrice = pHeaders.indexOf(CONFIG.COLUMNS.PRODUCTS.MOLD_PRICE);
        const colUnit = pHeaders.indexOf(CONFIG.COLUMNS.PRODUCTS.UNIT);
        /* [MỚI - Ý 1] Tìm cột đơn giá trong Master Sản Phẩm */
        const colPrice = pHeaders.indexOf(CONFIG.COLUMNS.PRODUCTS.PRICE || 'dongia');

        if (colSpId !== -1) {
          let prodChanged = false;
          for (let rp = 1; rp < prodData.length; rp++) {
            const pId = String(prodData[rp][colSpId] || '').trim();
            if (affectedProductsMap[pId]) {
              const spUpdates = affectedProductsMap[pId];
              if (spUpdates.mang !== undefined && colFilm !== -1) {
                prodData[rp][colFilm] = String(spUpdates.mang).trim();
                prodChanged = true;
              }
              if (spUpdates.giaKhuon !== undefined && colMoldPrice !== -1) {
                prodData[rp][colMoldPrice] = Number(spUpdates.giaKhuon) || 0;
                prodChanged = true;
              }
              if (spUpdates.donvi !== undefined && colUnit !== -1) {
                prodData[rp][colUnit] = String(spUpdates.donvi).trim();
                prodChanged = true;
              }
              /* [MỚI - Ý 1] Ghi đè đơn giá mới vào bảng 01_02_SAN_PHAM */
              if (spUpdates.dongia !== undefined && colPrice !== -1) {
                prodData[rp][colPrice] = Number(spUpdates.dongia) || 0;
                prodChanged = true;
              }
            }
          }
          if (prodChanged) {
            prodSheet.getRange(1, 1, prodData.length, prodData[0].length).setValues(prodData);
          }
        }
      }

      SpreadsheetApp.flush();
      return {
        success: true,
        updatedCount: updatedCount,
        message: 'Đã lưu thành công thông tin cho ' + updatedCount + ' dòng mặt hàng!'
      };

    } catch (err) {
      console.error('[OrderService.updateOrderLineItemsBulk] Lỗi:', err);
      return {
        success: false,
        message: 'Lỗi máy chủ trong quá trình cập nhật: ' + (err.message || String(err))
      };
    } finally {
      lock.releaseLock();
    }
  },

  /* [MỚI - Ý 2] Chuẩn hóa triệt để tra cứu tên nhân viên */
  resolveStaffName_: function(rawStaff, employeeByIdMap, allEmployeesList) {
    if (!rawStaff) return '---';
    const s = String(rawStaff).trim();
    if (!s || s === '---') return '---';

    // 1. Nếu chuỗi đã có sẵn tên dạng "NV0001 - Hoàng Thành" hoặc "NV-KD01 - Hoàng Thành"
    if (s.indexOf(' - ') !== -1) {
      const parts = s.split(' - ');
      if (parts[1] && parts[1].trim()) return parts[1].trim();
    }

    const sClean = s.toUpperCase().replace(/\s+/g, '');

    // 2. Tìm trực tiếp trong employeeByIdMap (case-insensitive & whitespace-stripped)
    if (employeeByIdMap) {
      if (employeeByIdMap[s]) return employeeByIdMap[s].hoTen;
      if (employeeByIdMap[sClean]) return employeeByIdMap[sClean].hoTen;

      // Quét toàn bộ keys của map để tránh lệch định dạng
      const mapKeys = Object.keys(employeeByIdMap);
      for (let k = 0; k < mapKeys.length; k++) {
        const key = mapKeys[k];
        if (key.toUpperCase().replace(/\s+/g, '') === sClean) {
          return employeeByIdMap[key].hoTen;
        }
      }
    }

    // 3. Chuẩn hóa mã số (ví dụ: NV-KD01 -> 1 -> NV0001)
    const matchNum = s.match(/\d+/);
    if (matchNum) {
      const numVal = parseInt(matchNum[0], 10);
      const targetStaffId = 'NV' + String(numVal).padStart(4, '0');
      if (employeeByIdMap && employeeByIdMap[targetStaffId]) {
        return employeeByIdMap[targetStaffId].hoTen;
      }
      if (allEmployeesList && Array.isArray(allEmployeesList)) {
        const found = allEmployeesList.find(e => {
          const eClean = String(e.maNV || '').toUpperCase().replace(/\s+/g, '');
          return eClean === targetStaffId || eClean.endsWith(String(numVal));
        });
        if (found && found.hoTen) return found.hoTen;
      }
    }

    return s;
  },
}