/**
 * ============================================================================
 * NAPACO ENTERPRISE OMS - SYSTEM CONFIGURATION & SECURITY PROPERTIES
 * File: Config.gs
 * Tầng kiến trúc: Config/Utils (Dùng chung cho Backend Service & Controller)
 * Nhiệm vụ:
 *  - Quản lý tập trung tên các bảng dữ liệu (Master & Operation) theo chuẩn Schema.
 *  - Bảo mật tuyệt đối ID_MASTER, ID_OPERATION, ID_STORAGES thông qua ScriptProperties.
 *  - Không để lộ bất kỳ Spreadsheet ID hay Folder ID nào xuống HTML/Client JS.
 * ============================================================================
 */

/**
 * BỔ SUNG VÀO CONFIG TRONG FILE Config.gs
 */
CONFIG.FORM_COLUMNS = {
  HEADER_ROW: 7,       // Dòng tiêu đề trên sheet TAO DON
  DATA_START_ROW: 8,   // Dòng bắt đầu nhập liệu
  ALIASES: {
    CUSTOMER: ['1. Khách Hàng (Bắt buộc)', '1. Khách Hàng', 'Khách Hàng', 'Tên Khách Hàng', 'Khach Hang'],
    PRODUCT:  ['2. Sản Phẩm (Bắt buộc)', '2. Sản Phẩm', 'Sản Phẩm', 'Tên Sản Phẩm', 'San Pham'],
    CODE:     ['3. Code SP', 'Code SP', 'Code Sản Phẩm', 'Mã SP'],
    UNIT:     ['4. Đơn Vị', 'Đơn Vị', 'ĐVT'],
    PRICE:    ['5. Đơn Giá (₫)', 'Đơn Giá (₫)', 'Đơn Giá', 'Giá Bán', 'Giá'],
    QTY:      ['6. Số Lượng', 'Số Lượng', 'SL'],
    PO:       ['7. Số PO', 'Số PO', 'PO', 'So_PO'],
    NOTE:     ['8. Ghi Chú Đơn Hàng', '8. Ghi Chú', 'Ghi Chú Đơn Hàng', 'Ghi Chú', 'Ghi chu'],
    LINK_PO:  ['9. Link File PO (Drive)', '9. Link PO', 'Link File PO', 'Link PO', 'Link']
  }
};

const CONFIG = {
  SYSTEM_NAME: 'NAPACO ERP',
  VERSION: '1.0',
  TIMEZONE: 'Asia/Ho_Chi_Minh',

  // Tên các khóa lưu trong ScriptProperties
  PROPERTY_KEYS: {
    ID_MASTER: 'NAPACO_ID_MASTER',
    ID_OPERATION: 'NAPACO_ID_OPERATION',
    ID_STORAGES: 'NAPACO_ID_STORAGES'
  },

  // Tên chuẩn của các Sheet trong Database (Tuyệt đối không tự ý đổi tên cột/sheet)
  SHEET_NAMES: {
    // Trong File 01. DB MASTER
    CUSTOMERS: '01_KHACH_HANG',
    PRODUCTS: '02_SAN_PHAM',
    SUPPLIERS: '03.DM_NhaCungCap',
    MATERIALS: '04.ThongSoVatLieu',
    MACHINES: '06.DM_May',
    EMPLOYEES: '07.NhanVien',

    // Trong File 02. Operation
    ORDERS: '01.DonHang',
    ORDER_DETAILS: '02.ChiTietDH',

    // Trong File Client Giao diện
    ORDER_WORKSPACE: 'TAO DON',
    CACHE_CATALOG: '_CACHE_DANH_MUC'
  },

  // Cấu hình vùng nhập liệu trên Sheet 'TAO DON'
  WORKSPACE_CONFIG: {
    HEADER_ROW: 7,
    DATA_START_ROW: 8,
    DATA_END_COL: 9, // Cột I (A8:I)
    CLEAR_RANGE_A1: 'A8:I'
  },

  // Cấu hình kích thước UI
  UI: {
    MODAL_WIDTH: 1150,
    MODAL_HEIGHT: 700
  },

  // Cấu hình Concurrency & LockService
  LOCK_TIMEOUT_MS: 30000, // 30 giây theo đặc tả
  CACHE_TTL_SECONDS: 300, // Cache danh mục 5 phút

  // Quy tắc nghiệp vụ mặc định
  DEFAULT_DELIVERY_DAYS: 10, // Ngày giao = TODAY() + 10 ngày
  DEFAULT_UNIT: 'Cái',
  ORDER_CODE_PREFIX: 'PO-',
  DH_PREFIX: 'DH_',
  CT_PREFIX: 'CT_',
  KH_PREFIX: 'KH',
  SP_PREFIX: 'SP',
  MAX_FILE_SIZE_BYTES: 15 * 1024 * 1024, // 15MB

  // Trạng thái vòng đời dòng chi tiết & đơn hàng
  LINE_STATUS: {
    NEW: 'Chờ sản xuất',
    MOLDING: 'Đang làm khuôn',
    IN_PRODUCTION: 'Đang sản xuất',
    COMPLETED: 'Xong',
    CLOSED: 'Kết thúc'
  },

  ORDER_STATUS: {
    NEW: 'Mới',
    IN_PRODUCTION: 'Đang sản xuất',
    COMPLETED: 'Hoàn tất',
    CANCELLED: 'Hủy'
  }
};

/**
 * Khởi tạo / cập nhật ScriptProperties cho hệ thống.
 * Chạy hàm này 1 lần trong Apps Script Editor (hoặc qua Menu Admin) để nạp ID vào vùng bảo mật.
 */
function setupSystemProperties() {
  const scriptProps = PropertiesService.getScriptProperties();
  const currentProps = scriptProps.getProperties();

  // Chỉ thiết lập giá trị mặc định nếu trong ScriptProperties chưa được cấu hình
  const defaultSecrets = {};
  if (!currentProps[CONFIG.PROPERTY_KEYS.ID_MASTER]) {
    defaultSecrets[CONFIG.PROPERTY_KEYS.ID_MASTER] = '1v4gg-9FcB2bxPa1w6gWVYkjlIFumcwacadUiuc77I1Y';
  }
  if (!currentProps[CONFIG.PROPERTY_KEYS.ID_OPERATION]) {
    defaultSecrets[CONFIG.PROPERTY_KEYS.ID_OPERATION] = '1GLmBtMQ7uqZPgYuyMhHo1AYx1GWkraC3ebt0WgKn44c';
  }
  if (!currentProps[CONFIG.PROPERTY_KEYS.ID_STORAGES]) {
    defaultSecrets[CONFIG.PROPERTY_KEYS.ID_STORAGES] = '1fiPuD9FkNE_cYVvtfiUsg3KHBz5zX-L6';
  }

  if (Object.keys(defaultSecrets).length > 0) {
    scriptProps.setProperties(defaultSecrets, false);
  }

  return {
    success: true,
    message: 'Đã thiết lập cấu hình bảo mật ScriptProperties thành công!'
  };
}

/**
 * Lấy đối tượng chứa các ID bảo mật từ ScriptProperties (Chỉ gọi ở tầng Backend)
 * Tự động nạp giá trị khởi tạo nếu lần đầu chạy chưa có trong PropertiesService
 * @return {{ID_MASTER: string, ID_OPERATION: string, ID_STORAGES: string}}
 */
function getSecureEnv_() {
  const scriptProps = PropertiesService.getScriptProperties();
  let idMaster = scriptProps.getProperty(CONFIG.PROPERTY_KEYS.ID_MASTER);
  let idOperation = scriptProps.getProperty(CONFIG.PROPERTY_KEYS.ID_OPERATION);
  let idStorages = scriptProps.getProperty(CONFIG.PROPERTY_KEYS.ID_STORAGES);

  if (!idMaster || !idOperation || !idStorages) {
    setupSystemProperties();
    idMaster = scriptProps.getProperty(CONFIG.PROPERTY_KEYS.ID_MASTER);
    idOperation = scriptProps.getProperty(CONFIG.PROPERTY_KEYS.ID_OPERATION);
    idStorages = scriptProps.getProperty(CONFIG.PROPERTY_KEYS.ID_STORAGES);
  }

  return {
    ID_MASTER: idMaster,
    ID_OPERATION: idOperation,
    ID_STORAGES: idStorages
  };
}