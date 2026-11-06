/**
 * FILE: Config.gs - NƠI DUY NHẤT ĐỊNH NGHĨA CẤU HÌNH HỆ THỐNG & TÊN CỘT GIAO DIỆN
 */
var CONFIG = {

  // ==========================================
  // 1. CỜ BẬT / TẮT LOG KIỂM TRA DỮ LIỆU
  // true: Xuất log chi tiết (Development / Staging)
  // false: Tắt toàn bộ log để tối ưu hiệu năng (Production)
  // ==========================================
  DEBUG_MODE: true,

  // BỔ SUNG THUỘC TÍNH MÚI GIỜ CHUẨN (Khắc phục dứt điểm lỗi Invalid argument: timeZone)
  TIMEZONE: 'Asia/Ho_Chi_Minh',

  // 1. CẤU HÌNH ID TRỰC TIẾP
  ID_MASTER: '1v4gg-9FcB2bxPa1w6gWVYkjlIFumcwacadUiuc77I1Y',
  ID_OPERATION: '1GLmBtMQ7uqZPgYuyMhHo1AYx1GWkraC3ebt0WgKn44c',
  ID_STORAGES: '1fiPuD9FkNE_cYVvtfiUsg3KHBz5zX-L6',

  // Khóa lưu trong ScriptProperties (khắc phục dứt điểm lỗi Cannot read properties of undefined reading ID_MASTER)
  PROPERTY_KEYS: {
    ID_MASTER: 'NAPACO_ID_MASTER',
    ID_OPERATION: 'NAPACO_ID_OPERATION',
    ID_STORAGES: 'NAPACO_ID_STORAGES'
  },

  // 2. TÊN CÁC SHEET
  SHEET_NAMES: {
    CUSTOMERS: '01_KHACH_HANG',
    PRODUCTS: '02_SAN_PHAM',
    EMPLOYEES: '07.NhanVien',
    ORDERS: '01.DonHang',
    ORDER_DETAILS: '02.ChiTietDH',
    ORDER_WORKSPACE: 'TAO DON',
    CACHE_CATALOG: '_CACHE_DANH_MUC'
  },

  // BỔ SUNG KHỐI TRẠNG THÁI VÒNG ĐỜI (Khắc phục dứt điểm lỗi reading 'NEW')
  ORDER_STATUS: {
    NEW: 'Mới',
    IN_PRODUCTION: 'Đang sản xuất',
    COMPLETED: 'Hoàn tất',
    CANCELLED: 'Hủy'
  },

  LINE_STATUS: {
    NEW: 'Chờ sản xuất',
    MOLDING: 'Đang làm khuôn',
    IN_PRODUCTION: 'Đang sản xuất',
    COMPLETED: 'Xong',
    CLOSED: 'Kết thúc'
  },

  // 3. ĐỊNH NGHĨA TẬP TRUNG TÊN CỘT GIAO DIỆN 'TAO DON' (DÒNG 7)
  FORM_COLUMNS: {
    HEADER_ROW: 7,
    DATA_START_ROW: 8,
    ALIASES: {
      CUSTOMER: ['1. Khách Hàng (Bắt buộc)', '1. Khách Hàng', 'Khách Hàng', 'Tên Khách Hàng', 'Khach Hang'],
      PRODUCT:  ['2. Sản Phẩm (Bắt buộc)', '2. Sản Phẩm', 'Sản Phẩm', 'Tên Sản Phẩm', 'San Pham'],
      CODE:     ['3. Code SP', 'Code SP', 'Code Sản Phẩm', 'Mã SP'],
      UNIT:     ['4. Đơn Vị', 'Đơn Vị', 'ĐVT', 'DVT'],
      PRICE:    ['5. Đơn Giá (₫)', 'Đơn Giá (₫)', 'Đơn Giá', 'Giá Bán', 'Giá'],
      QTY:      ['6. Số Lượng', 'Số Lượng', 'SL'],
      PO:       ['7. Số PO', 'Số PO', 'PO', 'So_PO'],
      NOTE:     ['8. Ghi Chú Đơn Hàng', '8. Ghi Chú', 'Ghi Chú Đơn Hàng', 'Ghi Chú', 'Ghi chu'],
      LINK_PO:  ['9. Link File PO (Drive)', '9. Link PO', 'Link File PO', 'Link PO', 'Link']
    }
  },

// ========================================================
  // NƠI DUY NHẤT ĐỊNH NGHĨA TÊN CỘT DATABASE (SINGLE SOURCE OF TRUTH)
  // Toàn bộ logic Backend/Frontend chỉ gọi qua CONFIG.COLUMNS
  // ========================================================
  COLUMNS: {
    // Bảng 01_KHACH_HANG
    CUSTOMERS: {
      ID: 'Ma_KH',
      NAME: 'name',
      ADDRESS: 'address',
      CONTACTER: 'contacter',
      PHONE: 'phone'
    },
    // Bảng 02_SAN_PHAM
    PRODUCTS: {
      ID: 'SP_ID',
      CUSTOMER_ID: 'customer_id',
      FILM: 'mang',
      TYPE: 'type',
      CODE_CUST: 'Code_Sanpham',
      NAME: 'name_napaco',
      CAVITY: 'cavity',
      STOCK: 'tonkho',
      UNIT: 'donvi',
      LENGTH: 'dai',
      WIDTH: 'rong',
      HEIGHT: 'cao',
      IMAGE: 'imageURL',
      PACKING: 'donggoi',
      MOLD_PRICE: 'Gia Khuon',
      PRICE: 'dongia'
    },
    // Bảng 07.NhanVien
    EMPLOYEES: {
      ID: 'Ma_NV',
      NAME: 'name',
      DEPARTMENT: 'phong_ban',
      ROLE: 'role',
      EMAIL: 'email',
      PHONE: 'sdt',
      ACCESS: 'quyen_truy_cap',
      STATUS: 'trang_thai'
    },
    // Bảng 01.DonHang
    ORDERS: {
      ID: 'Ma_DH',
      CUSTOMER_ID: 'MaKH',
      ORDER_DATE: 'NgayDat',
      DELIVERY_DATE: 'NgayGiaoDK',
      PO_NUMBER: 'So_PO',
      FILE_URL: 'File_PO'
    },
    // Bảng 02.ChiTietDH
    ORDER_DETAILS: {
      ID: 'ID_Chi_Tiet',
      ORDER_ID: 'ID_DonHang',
      PRODUCT_ID: 'Ma_SP',
      QTY_ORDER: 'So_Luong_Dat',
      PRICE: 'Don_Gia',
      QTY_RUN: 'So_Luong_Da_Chay',
      QTY_DELIVERED: 'SL_Da_Giao',
      QTY_REMAINING: 'SL_Con_Lai',
      NOTE: 'Ghi_Chu',
      DELIVERY_DATE: 'Ngay_Giao',
      STAFF_ID: 'Ma_NV',
      STATUS: 'Trang_Thai'
    }
  },

  DEFAULT_DELIVERY_DAYS: 10,
  DEFAULT_STATUS_DETAIL: 'Chờ sản xuất',
  ORDER_CODE_PREFIX: 'PO-'
};

function setupSystemProperties() {
  const scriptProps = PropertiesService.getScriptProperties();
  const currentProps = scriptProps.getProperties();

  const defaultSecrets = {};
  if (!currentProps[CONFIG.PROPERTY_KEYS.ID_MASTER]) {
    defaultSecrets[CONFIG.PROPERTY_KEYS.ID_MASTER] = CONFIG.ID_MASTER;
  }
  if (!currentProps[CONFIG.PROPERTY_KEYS.ID_OPERATION]) {
    defaultSecrets[CONFIG.PROPERTY_KEYS.ID_OPERATION] = CONFIG.ID_OPERATION;
  }
  if (!currentProps[CONFIG.PROPERTY_KEYS.ID_STORAGES]) {
    defaultSecrets[CONFIG.PROPERTY_KEYS.ID_STORAGES] = CONFIG.ID_STORAGES;
  }

  if (Object.keys(defaultSecrets).length > 0) {
    scriptProps.setProperties(defaultSecrets, false);
  }

  return {
    success: true,
    message: 'Đã thiết lập cấu hình bảo mật ScriptProperties thành công!'
  };
}

function getSecureEnv_() {
  const scriptProps = PropertiesService.getScriptProperties();
  let idMaster = scriptProps.getProperty(CONFIG.PROPERTY_KEYS.ID_MASTER);
  let idOperation = scriptProps.getProperty(CONFIG.PROPERTY_KEYS.ID_OPERATION);
  let idStorages = scriptProps.getProperty(CONFIG.PROPERTY_KEYS.ID_STORAGES);

  if (!idMaster || !idOperation || !idStorages) {
    setupSystemProperties();
    idMaster = scriptProps.getProperty(CONFIG.PROPERTY_KEYS.ID_MASTER) || CONFIG.ID_MASTER;
    idOperation = scriptProps.getProperty(CONFIG.PROPERTY_KEYS.ID_OPERATION) || CONFIG.ID_OPERATION;
    idStorages = scriptProps.getProperty(CONFIG.PROPERTY_KEYS.ID_STORAGES) || CONFIG.ID_STORAGES;
  }

  return {
    ID_MASTER: idMaster,
    ID_OPERATION: idOperation,
    ID_STORAGES: idStorages
  };
}