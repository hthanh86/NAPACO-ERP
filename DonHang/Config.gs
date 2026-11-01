/**
 * FILE: Config.gs
 * Cấu hình ID trực tiếp, tên sheet và Bảng định nghĩa ánh xạ cột (Schema Mapping)
 */

const CONFIG = {
  // ID trực tiếp của các file & thư mục
  ID_MASTER: '1v4gg-9FcB2bxPa1w6gWVYkjlIFumcwacadUiuc77I1Y', // File 01. DB MASTER
  ID_OPERATION: '1GLmBtMQ7uqZPgYuyMhHo1AYx1GWkraC3ebt0WgKn44c', // File 02. Operation
  ID_STORAGES: '1fiPuD9FkNE_cYVvtfiUsg3KHBz5zX-L6', // Thư mục lưu file PO

  // Tên các Sheet giao diện
  SHEET_TAO_DON: 'TAO DON',
  SHEET_CACHE: '_CACHE_DANH_MUC',
  HEADER_ROW_TAO_DON: 7, // Tiêu đề sheet TAO DON nằm ở dòng 7
  MAX_ROWS_TAO_DON: 50,  // Số dòng tối đa của form tạo đơn

  // Tên các Sheet Database
  SHEET_KHACH_HANG: '01_KHACH_HANG',
  SHEET_SAN_PHAM: '02_SAN_PHAM',
  SHEET_NHAN_VIEN: '07.NhanVien',
  SHEET_DON_HANG: '01.DonHang',
  SHEET_CHI_TIET_DH: '02.ChiTietDH',

  DEFAULT_DELIVERY_DAYS: 10,
  DEFAULT_STATUS_DETAIL: 'Chờ sản xuất'
};

/**
 * 1. BẢNG ĐỊNH NGHĨA ÁNH XẠ CỘT (TABLE MAPPING SCHEMA)
 * Dễ dàng bổ sung, sửa đổi tên cột tương ứng giữa Form và Database
 */
const TABLE_MAPPING_SCHEMA = {
  // --- BẢNG ĐƠN HÀNG (01.DonHang) ---
  DON_HANG: {
    targetFileId: CONFIG.ID_OPERATION,
    targetSheetName: CONFIG.SHEET_DON_HANG,
    headerRow: 1,
    // [Key trong dữ liệu Form] -> [Danh sách tên cột có thể có trên Sheet Database]
    fields: [
      { key: 'maDH',         formHeaders: [],               targetHeaders: ['Mã Đơn Hàng', 'Ma_DH', 'Mã ĐH'] },
      { key: 'maKH',         formHeaders: ['Ma_KH', 'Khách Hàng'],        targetHeaders: ['MaKH'] },
      { key: 'soPO',         formHeaders: ['Số PO', 'So_PO', 'PO'],        targetHeaders: ['Số PO', 'So_PO', 'PO'] },
      { key: 'ngayDat',      formHeaders: ['Ngày Đặt'],        targetHeaders: ['Ngày Đặt', 'Ngay_Dat', 'NgayDat'] },
      { key: 'ngayGiaoDK',   formHeaders: ['Ngày Giao'],       targetHeaders: ['Ngày Giao', 'Ngay_Giao_DK', 'Ngày Giao Dự Kiến', 'NgayGiaoDK'] },
      { key: 'fileUrl',      formHeaders: ['Link PO'],      targetHeaders: ['File Đính Kèm', 'Link PO', 'File_PO', 'Đính Kèm'] },
    ]
  },

  // --- BẢNG CHI TIẾT ĐƠN HÀNG (02.ChiTietDH) ---
  CHI_TIET_DH: {
    targetFileId: CONFIG.ID_OPERATION,
    targetSheetName: CONFIG.SHEET_CHI_TIET_DH,
    headerRow: 1,
    fields: [
      { key: 'maCT',         formHeaders: [],       targetHeaders: ['Mã Chi Tiết', 'ID_Chi_Tiet', 'Mã CT'] },
      { key: 'maDH',         formHeaders: ['Mã Đơn Hàng', 'Mã Đơn'],   targetHeaders: ['Mã Đơn Hàng', 'ID_DonHang', 'Mã ĐH'] },
      { key: 'masp',         formHeaders: ['Mã Sản Phẩm','Sản Phẩm'],     targetHeaders: ['Mã Sản Phẩm', 'Ma_SP'] },
      { key: 'qty',          formHeaders: ['Số Lượng'],     targetHeaders: ['Số Lượng', 'So_Luong_Dat', 'SL'] },
      { key: 'price',        formHeaders: ['Đơn Giá', 'Price'],   targetHeaders: ['Đơn Giá', 'Don_Gia'] },
      { key: 'sl_dachay',    targetHeaders: ['So_Luong_Da_Chay'], defaultValue:0 },
      { key: 'sl_dagiao',    targetHeaders: ['SL_Da_Giao'], defaultValue:0 },
      { key: 'sl_conlai',    targetHeaders: ['SL_Con_Lai'] },
      { key: 'note',         formHeaders: ['Ghi chú'],  targetHeaders: ['Ghi Chú'] },
      { key: 'deliveryDate', formHeaders: ['Ngày Giao'],  targetHeaders: ['Ngày Giao', 'Ngay_Giao'] },
      { key: 'manv',         formHeaders: ['Nhân Viên'],  targetHeaders: ['Ma_NV'] },
      { key: 'status',       targetHeaders: ['Trạng Thái', 'Trang_Thai'], defaultValue: CONFIG.DEFAULT_STATUS_DETAIL }
    ]
  },

  // --- BẢNG SẢN PHẨM MASTER (02_SAN_PHAM) - Khi phát sinh sản phẩm mới ---
  SAN_PHAM_MASTER: {
    targetFileId: CONFIG.ID_MASTER,
    targetSheetName: CONFIG.SHEET_SAN_PHAM,
    headerRow: 1,
    fields: [
      { key: 'spId',         formHeaders: [], targetHeaders: ['SP_ID'] },
      { key: 'customerId',   formHeaders: ['MaKH', 'Ma_KH', 'Mã Khách Hàng'], targetHeaders: ['customer_id'] },
      { key: 'mang', formHeaders: ['Màng'], targetHeaders: ['mang'] },
      { key: 'type', targetHeaders: ['type'] },
      { key: 'codeSanpham',formHeaders: ['Code Sản Phẩm'], targetHeaders: ['Code_Sanpham'] },
      { key: 'nameNapaco', formHeaders: ['Tên Sản Phẩm', 'Sản Phẩm'], targetHeaders: ['name_napaco'] },
      { key: 'cavity', formHeaders: ['cavity'], targetHeaders: ['cavity'] },
      { key: 'tonkho', formHeaders: ['Tồn Kho'], targetHeaders: ['tonkho'] },
      { key: 'donvi', formHeaders: ['Đơn Vị'], targetHeaders: ['donvi'] },
      { key: 'dai', formHeaders: ['Chiều Dài'],  targetHeaders: ['dai'] },
      { key: 'rong', formHeaders: ['Chiều Rộng'], targetHeaders: ['rong'] },
      { key: 'cao', formHeaders: ['Chiều Cao'],  targetHeaders: ['cao'] },
      { key: 'imageUrl', formHeaders: ['Link File'], targetHeaders: ['imageURL'] },
      { key: 'donggoi', formHeaders: ['Đóng Gói'],  targetHeaders: ['donggoi'] },
      { key: 'dongia', formHeaders: ['Đơn Giá'], targetHeaders: ['dongia'] }
          ]
    },

  // --- BẢNG SẢN PHẨM MASTER (02_SAN_PHAM) - Khi phát sinh sản phẩm mới ---
  KHACH_HANG_MASTER: {
    targetFileId: CONFIG.ID_MASTER,
    targetSheetName: CONFIG.SHEET_KHACH_HANG,
    headerRow: 1,
    fields: [
      { key: 'Ma_KH', formHeaders: [],  targetHeaders: ['Ma_KH'] },
      { key: 'name', formHeaders: ['Tên Khách Hàng', 'Tên Khách', 'Name'], targetHeaders: ['name'] },
      { key: 'address', formHeaders: ['Địa Chỉ'],  targetHeaders: ['address'] },
      { key: 'contacter', formHeaders: ['Người liên hệ'],  targetHeaders: ['contacter'] },
      { key: 'phone', formHeaders: ['SĐT'],  targetHeaders: ['phone'] }
    ]
  }
};

/**
 * BẢNG ÁNH XẠ TRỰC TIẾP CỘT TỪ FORM "TAO DON" SANG DATABASE
 * - formHeaders: Tiêu đề cột trên sheet TAO DON (hỗ trợ nhiều alias)
 * - targetTable: 'DON_HANG' hoặc 'CHI_TIET_DH'
 * - targetHeaders: Tiêu đề cột đích trên database tương ứng
 * - type: Kiểu dữ liệu để chuẩn hóa trước khi ghi (string, number, date, currency)

const FORM_TO_DB_MAPPING = {
  // ================= 1. CÁC CỘT ĐƯỢC ÁNH XẠ VÀO 01.DonHang =================
  soPO: {
    targetTable: 'DON_HANG',
    formHeaders: ['Số PO', 'PO', 'So_PO'],
    targetHeaders: ['Số PO', 'So_PO', 'PO'],
    type: 'string',
    required: true
  },
  customerName: {
    targetTable: 'DON_HANG',
    formHeaders: ['Khách Hàng', 'Tên Khách Hàng', 'Khach Hang'],
    targetHeaders: ['Khách Hàng', 'Tên Khách Hàng', 'Ma_KH'],
    type: 'string',
    required: true
  },
  salesPerson: {
    targetTable: 'DON_HANG',
    formHeaders: ['Nhân Viên KD', 'Nhân Viên', 'NVKD'],
    targetHeaders: ['Nhân Viên KD', 'Nhân Viên', 'Ma_NV'],
    type: 'string'
  },
  ngayDat: {
    targetTable: 'DON_HANG',
    formHeaders: ['Ngày Đặt', 'Ngay Dat', 'Ngay_Dat'],
    targetHeaders: ['Ngày Đặt', 'Ngay_Dat', 'Ngay Dat'],
    type: 'date'
  },
  ngayGiaoDK: {
    targetTable: 'DON_HANG',
    formHeaders: ['Ngày Giao', 'Ngay Giao', 'Ngay_Giao'],
    targetHeaders: ['Ngày Giao', 'Ngay_Giao_DK', 'Ngày Giao Dự Kiến'],
    type: 'date'
  },
  orderNote: {
    targetTable: 'DON_HANG',
    formHeaders: ['Ghi chú', 'Ghi chu', 'Ghi Chú Đơn Hàng'],
    targetHeaders: ['Ghi Chú', 'Ghi_Chu', 'Ghi chu'],
    type: 'string'
  },

  // ================= 2. CÁC CỘT ĐƯỢC ÁNH XẠ VÀO 02.ChiTietDH =================
  codeSP: {
    targetTable: 'CHI_TIET_DH',
    formHeaders: ['Code Sản Phẩm', 'Code SP', 'Code_Sanpham'],
    targetHeaders: ['Code Sản Phẩm', 'Code_Sanpham', 'Code SP'],
    type: 'string'
  },
  prodName: {
    targetTable: 'CHI_TIET_DH',
    formHeaders: ['Sản Phẩm', 'Tên Sản Phẩm', 'San Pham'],
    targetHeaders: ['Tên Sản Phẩm', 'name_napaco', 'Sản Phẩm'],
    type: 'string',
    required: true
  },
  unit: {
    targetTable: 'CHI_TIET_DH',
    formHeaders: ['Đơn Vị', 'ĐVT', 'DVT', 'Đơn vị tính'],
    targetHeaders: ['Đơn Vị', 'donvi', 'ĐVT'],
    type: 'string'
  },
  qty: {
    targetTable: 'CHI_TIET_DH',
    formHeaders: ['Số Lượng', 'SL', 'So Luong'],
    targetHeaders: ['Số Lượng', 'soluong', 'SL'],
    type: 'number'
  },
  price: {
    targetTable: 'CHI_TIET_DH',
    formHeaders: ['Đơn Giá', 'Giá', 'Don Gia'],
    targetHeaders: ['Đơn Giá', 'dongia', 'Giá'],
    type: 'number'
  },
  lineTotal: {
    targetTable: 'CHI_TIET_DH',
    formHeaders: ['Thành Tiền', 'Thanh Tien'],
    targetHeaders: ['Thành Tiền', 'thanhtien'],
    type: 'currency'
  },
  lineDeliveryDate: {
    targetTable: 'CHI_TIET_DH',
    formHeaders: ['Ngày Giao', 'Ngay Giao'],
    targetHeaders: ['Ngày Giao', 'Ngay_Giao'],
    type: 'date'
  }
};
 */
/**
 * Tự động lấy thư mục gốc chứa file Đơn Hàng hiện tại
 */
function getProjectRootFolder() {
  const currentFileId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const file = DriveApp.getFileById(currentFileId);
  const parents = file.getParents();
  if (parents.hasNext()) {
    return parents.next();
  }
  throw new Error('Không tìm thấy thư mục gốc chứa file Đơn Hàng!');
}

/**
 * Mở Spreadsheet theo tên file trong cùng thư mục gốc (có lưu cache để tăng tốc)
 */
function getSpreadsheetByNameInRoot(fileName, propKey) {
  const scriptProps = PropertiesService.getScriptProperties();
  let cachedId = scriptProps.getProperty(propKey);

  if (cachedId) {
    try {
      return SpreadsheetApp.openById(cachedId);
    } catch (e) {
      scriptProps.deleteProperty(propKey);
    }
  }

  const rootFolder = getProjectRootFolder();
  const files = rootFolder.getFilesByName(fileName);
  if (files.hasNext()) {
    const targetFile = files.next();
    scriptProps.setProperty(propKey, targetFile.getId());
    return SpreadsheetApp.openById(targetFile.getId());
  }
  throw new Error('Không tìm thấy file "' + fileName + '" trong thư mục gốc!');
}

/**
 * Lấy thư mục Storages để lưu file PO
 */
function getStorageFolder() {
  const scriptProps = PropertiesService.getScriptProperties();
  let cachedFolderId = scriptProps.getProperty('STORAGE_FOLDER_ID');

  if (cachedFolderId) {
    try {
      return DriveApp.getFolderById(cachedFolderId);
    } catch (e) {
      scriptProps.deleteProperty('STORAGE_FOLDER_ID');
    }
  }

  const rootFolder = getProjectRootFolder();
  const folders = rootFolder.getFoldersByName(CONFIG.FOLDER_NAME_STORAGES);
  if (folders.hasNext()) {
    const folder = folders.next();
    scriptProps.setProperty('STORAGE_FOLDER_ID', folder.getId());
    return folder;
  }
  throw new Error('Không tìm thấy thư mục "' + CONFIG.FOLDER_NAME_STORAGES + '" trong thư mục gốc!');
}