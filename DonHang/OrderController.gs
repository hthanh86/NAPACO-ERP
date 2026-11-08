/**
 * ============================================================================
 * NAPACO ENTERPRISE OMS - ORDER CONTROLLER (WEB APP & RPC ROUTER)
 * File: OrderController.gs
 * Tầng kiến trúc: Tầng 2 - Controller Layer
 * Nhiệm vụ:
 *  - Cung cấp hàm doGet(e) phục vụ Standalone Web App (routing giữa màn hình
 *    Tra cứu đơn hàng Dual-View 'OrderList_View' và Tạo đơn hàng 'OrderModal').
 *  - Cung cấp các hàm Global RPC để Client HTML gọi qua google.script.run.
 *  - Kiểm soát thông báo lỗi chuẩn hóa, không làm lộ thông tin nhạy cảm.
 * ============================================================================
 */

/**
 * Entry point cho Standalone Web App
 * Hỗ trợ tham số URL:
 *  - ?view=list (Mặc định): Mở giao diện Tra cứu đơn hàng kép (OrderList_View.html)
 *  - ?view=create: Mở giao diện Tạo đơn hàng (OrderModal.html)
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const view = String(params.view || 'list').toLowerCase();

    if (view === 'create' || view === 'modal') {
      const sourceSheetId = params.sourceId || '';
      const initRes = OrderController.getOrderModalInitData(sourceSheetId);
      const template = HtmlService.createTemplateFromFile('OrderModal');
      const jsonStr = JSON.stringify(initRes.success ? initRes.data : {});
      template.initialDataJson = Utilities.base64Encode(Utilities.newBlob(jsonStr, 'application/json').getBytes());

      return template.evaluate()
        .setTitle('NAPACO OMS - Tạo Phiếu Đơn Đặt Hàng Mới')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }

    // Mặc định mở giao diện Tra cứu đơn hàng Dual-View
    const listTemplate = HtmlService.createTemplateFromFile('OrderList_View');
    return listTemplate.evaluate()
      .setTitle('NAPACO OMS - Hệ Thống Tra Cứu & Quản Lý Đơn Hàng')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:24px;color:#991b1b;background:#fef2f2;border:1px solid #f87171;border-radius:8px;">' +
      '<h3>⚠️ Lỗi khởi chạy ứng dụng NAPACO OMS</h3>' +
      '<p>' + String(err.message || err) + '</p></div>'
    );
  }
}

/**
 * Hàm hỗ trợ nhúng file HTML con (nếu dùng <?!= include(...) ?>)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Đối tượng Controller điều phối nghiệp vụ Module 1
 */
const OrderController = {

  /**
   * 1. Lấy dữ liệu khởi tạo cho OrderModal.html (bao gồm xác thực kép, danh mục & các dòng từ TAO DON)
   * @param {string=} optSourceSpreadsheetId ID bảng tính nguồn (nếu gọi từ Sheet TAO DON)
   */
  getOrderModalInitData: function(optSourceSpreadsheetId) {
    try {
      const data = OrderService.buildModalInitialData(optSourceSpreadsheetId);
      return {
        success: true,
        data: data
      };
    } catch (err) {
      console.error('[OrderController.getOrderModalInitData] Error:', err);
      return {
        success: false,
        error: {
          code: 'ERR_INIT_MODAL',
          message: err.message || 'Không thể tải dữ liệu khởi tạo form đơn hàng.'
        }
      };
    }
  },

  /**
   * 2. Lấy dữ liệu danh mục Master + thông tin xác thực nhân sự cho Web App độc lập
   */
  getOrderFormData: function() {
    try {  
        const data = OrderService.getCatalogAndAuthData();  
        return {  
          success: true,  
          customers: data.customers,  
          products: data.products,  
          salesStaffList: data.salesStaffList,  
          currentUser: data.currentUser,  
          serverDate: data.serverDate,  
          defaultDeliveryDate: data.defaultDeliveryDate,
          systemConfig: {
            filterGroups: CONFIG.FILTER_STATUS_GROUPS,
            statusGroupsList: CONFIG.FILTER_STATUS_GROUPS_LIST || {
              ACTIVE_ALL: [CONFIG.LINE_STATUS.NEW, CONFIG.LINE_STATUS.MOLDING, CONFIG.LINE_STATUS.IN_PRODUCTION],
              DONE_ALL: [CONFIG.LINE_STATUS.COMPLETED, CONFIG.LINE_STATUS.CLOSED, CONFIG.ORDER_STATUS.CANCELLED]
            }
          }
        };
    } catch (err) {
      console.error('[OrderController.getOrderFormData] Error:', err);
      return {
        success: false,
        error: {
          code: 'ERR_INIT_FORM',
          message: err.message || 'Không thể tải danh mục hệ thống.'
        }
      };
    }
  },

  /**
   * 3. Lưu giao dịch tạo đơn hàng liên tầng (Flow-based Upsert) có LockService
   * Đáp ứng chuẩn API Contract: saveOrderTransaction(payload)
   */
  
  saveOrderTransaction: function(payload) {
    try {
      if (!payload || !payload.poGroups || payload.poGroups.length === 0) {
        throw new Error("Dữ liệu đơn hàng gửi lên không hợp lệ hoặc danh sách rỗng!");
      }
      // Chuyển toàn bộ việc quản lý Lock và ghi DB cho OrderService
      return OrderService.saveOrderTransaction(payload);
    } catch (err) {
      console.error('[OrderController.saveOrderTransaction] Error:', err);
      return {
        success: false,
        error: {
          code: err.code || 'ERR_SAVE_ORDER',
          message: err.message || 'Lỗi trong quá trình lưu đơn hàng.'
        },
        message: err.message || 'Lỗi trong quá trình lưu đơn hàng.'
      };
    }
  },

  /**
   * 4. Tra cứu danh sách đơn hàng và chi tiết mặt hàng theo 2 chế độ xem kép (HEADER | LINE_ITEM)
   * Đáp ứng chuẩn API Contract: queryOrdersAndLineItems(searchParams)
   */
  queryOrdersAndLineItems: function(searchParams) {
    try {
      return OrderService.queryOrdersAndLineItems(searchParams || {});
    } catch (err) {
      console.error('[OrderController.queryOrdersAndLineItems] Error:', err);
      return {
        success: false,
        viewMode: (searchParams && searchParams.viewMode) ? searchParams.viewMode : 'HEADER',
        total: 0,
        data: [],
        error: {
          code: 'ERR_QUERY_ORDERS',
          message: err.message || 'Lỗi khi truy vấn dữ liệu đơn hàng.'
        }
      };
    }
  },

  /* [MỚI - SPEC-OMS-LINEITEM-202611-01] Controller tiếp nhận lệnh Bulk Edit */
  updateOrderLineItemsBulk: function(payload) {
    try {
      return OrderService.updateOrderLineItemsBulk(payload);
    } catch (err) {
      console.error('[OrderController.updateOrderLineItemsBulk] Error:', err);
      return {
        success: false,
        message: err.message || 'Lỗi xử lý chỉnh sửa hàng loạt.'
      };
    }
  },

  /**
   * 5. Hàm phục vụ hiển thị nhanh trên Sheet (nếu gọi từ ClientLauncher)
   */
  getOrdersForSheet: function(viewMode) {
    try {
      const res = OrderService.queryOrdersAndLineItems({
        viewMode: viewMode === 'LINE_ITEM' ? 'LINE_ITEM' : 'HEADER',
        page: 1,
        pageSize: 500
      });
      return {
        success: true,
        rows: res.data || []
      };
    } catch (err) {
      return {
        success: false,
        error: { message: err.message }
      };
    }
  }
};

// ============================================================================
// CÁC HÀM GLOBAL RPC EXPOSED CHO GOOGLE.SCRIPT.RUN TỪ CLIENT HTML
// ============================================================================

/**
 * RPC: Lấy dữ liệu khởi tạo Modal tạo đơn hàng
 */
function getOrderModalInitData(sourceSpreadsheetId) {
  return OrderController.getOrderModalInitData(sourceSpreadsheetId);
}

/**
 * RPC: Lấy danh mục Khách hàng, Sản phẩm, Nhân viên và User Auth
 */
function getOrderFormData() {
  return OrderController.getOrderFormData();
}

/**
 * RPC chuẩn theo tài liệu đặc tả: Lưu giao dịch đơn hàng liên tầng
 */
function saveOrderTransaction(payload) {
  return OrderController.saveOrderTransaction(payload);
}

/**
 * RPC tương thích ngược với code cũ: saveOrderFromModal(payload)
 * Tự động ánh xạ về cùng luồng xử lý saveOrderTransaction
 */
function saveOrderFromModal(payload) {
  const result = OrderController.saveOrderTransaction(payload);
  if (!result.success) {
    throw new Error(result.message || (result.error && result.error.message) || 'Lỗi lưu đơn hàng');
  }
  return result;
}

/**
 * RPC tương thích với mã nguồn mẫu: submitOrderTransaction(payload)
 */
function submitOrderTransaction(payload) {
  return OrderController.saveOrderTransaction(payload);
}

/**
 * RPC chuẩn theo tài liệu đặc tả: Truy vấn đơn hàng & dòng mặt hàng (Dual-View)
 */
function queryOrdersAndLineItems(searchParams) {
  return OrderController.queryOrdersAndLineItems(searchParams);
}

/**
 * [MỚI - SPEC-OMS-LINEITEM-202611-01] Global RPC gọi Chỉnh sửa hàng loạt dòng mặt hàng
 */
function updateOrderLineItemsBulk(payload) {
  return OrderController.updateOrderLineItemsBulk(payload);
}

/**
 * Router đón request POST từ ClientLauncher.gs của Google Sheets
 */
function doPost(e) {
  try {
    // Kiểm tra đầu vào an toàn
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Payload request trống hoặc không hợp lệ.");
    }
    
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var payload = data.payload || {};
    var result = null;

    if (action === 'SYNC_CATALOG') {
      result = OrderService.syncCacheCatalogFromDatabase();
    } else if (action === 'GET_MODAL_INIT') {
      result = OrderService.buildModalInitialData(payload.sourceSpreadsheetId);
    } else if (action === 'SAVE_ORDER') {
      result = OrderService.saveOrderTransaction(payload);
    } else {
      throw new Error("Action không tồn tại: " + action);
    }

    // Trả về JSON chuẩn 100%
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: result
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    // Bắt toàn bộ lỗi bên trong service/controller và ép gói về JSON, tuyệt đối không cho văng HTML
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: {
        message: err.message || String(err),
        stack: err.stack || ""
      }
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * CẦU NỐI CÔNG KHAI ĐỂ CLIENT LAUNCHER GỌI QUA THƯ VIỆN BACKEND
 */

function refreshOrderData(targetSpreadsheetId) {
  return OrderService.syncCacheCatalogFromDatabase(targetSpreadsheetId);
}

/**
 * ĐỌC DỮ LIỆU TỪ SHEET 'TAO DON' THEO TÊN CỘT ĐỘNG VÀ ĐÓNG GÓI CHO MODAL
 */
/**
 * ĐỌC DỮ LIỆU TỪ SHEET 'TAO DON' THEO TÊN CỘT ĐỘNG VÀ ĐÓNG GÓI CHO MODAL
 */
function getOrderFormDataForModal(optSourceSpreadsheetId) {
  try {
    const ss = optSourceSpreadsheetId 
      ? SpreadsheetApp.openById(optSourceSpreadsheetId) 
      : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return JSON.stringify({ poGroups: [], salesStaffList: [] });

    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ORDER_WORKSPACE || 'TAO DON');
    if (!sheet) return JSON.stringify({ poGroups: [], salesStaffList: [] });

    const headerRowIdx = CONFIG.FORM_COLUMNS.HEADER_ROW;
    const startDataRow = CONFIG.FORM_COLUMNS.DATA_START_ROW;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < startDataRow || lastCol < 1) {
      return JSON.stringify({ poGroups: [], salesStaffList: [] });
    }

    // 1. Quét tiêu đề dòng 7 và tìm vị trí cột dựa trên CONFIG
    const headers = sheet.getRange(headerRowIdx, 1, 1, lastCol).getValues()[0];
    const findColIdx = (aliases) => {
      for (let c = 0; c < headers.length; c++) {
        const text = String(headers[c] || '').trim().toLowerCase();
        if (aliases.some(a => text.includes(a.toLowerCase()))) return c;
      }
      return -1;
    };

    const c_kh    = findColIdx(CONFIG.FORM_COLUMNS.ALIASES.CUSTOMER);
    const c_sp    = findColIdx(CONFIG.FORM_COLUMNS.ALIASES.PRODUCT);
    const c_code  = findColIdx(CONFIG.FORM_COLUMNS.ALIASES.CODE);
    const c_unit  = findColIdx(CONFIG.FORM_COLUMNS.ALIASES.UNIT);
    const c_price = findColIdx(CONFIG.FORM_COLUMNS.ALIASES.PRICE);
    const c_qty   = findColIdx(CONFIG.FORM_COLUMNS.ALIASES.QTY);
    const c_po    = findColIdx(CONFIG.FORM_COLUMNS.ALIASES.PO);
    const c_note  = findColIdx(CONFIG.FORM_COLUMNS.ALIASES.NOTE);
    const c_link  = findColIdx(CONFIG.FORM_COLUMNS.ALIASES.LINK_PO);

    // 2. Đọc Cache Sản phẩm và Danh sách Nhân Viên KD từ _CACHE_DANH_MUC
    const cacheSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CACHE_CATALOG || '_CACHE_DANH_MUC');
    const cacheProductMap = {};
    const salesStaffList = [];

    if (cacheSheet && cacheSheet.getLastRow() >= 2) {
      const cacheRowsCount = cacheSheet.getLastRow() - 1;

      // Đọc Sản phẩm từ Cột A:E
      const cacheData = cacheSheet.getRange(2, 1, cacheRowsCount, 5).getValues();
      for (let i = 0; i < cacheData.length; i++) {
        const pName = String(cacheData[i][0] || '').trim().toLowerCase();
        if (pName) {
          cacheProductMap[pName] = {
            unit: cacheData[i][3] || 'Cái',
            price: Number(cacheData[i][4]) || 0
          };
        }
      }

      // Đọc Danh sách Nhân Viên KD từ Cột G (Cột 7)
      if (cacheSheet.getLastColumn() >= 7) {
        const staffVals = cacheSheet.getRange(2, 7, cacheRowsCount, 1).getValues();
        for (let s = 0; s < staffVals.length; s++) {
          const val = String(staffVals[s][0] || '').trim();
          if (val) {
            const parts = val.split(' - ');
            salesStaffList.push({
              display: val,
              maNV: parts[0] ? parts[0].trim() : val,
              hoTen: parts[1] ? parts[1].trim() : val
            });
          }
        }
      }
    }

    // Fallback: nếu Cache chưa có NV thì lấy qua Service
    if (salesStaffList.length === 0 && typeof OrderService.getCatalogAndAuthData === 'function') {
      const authData = OrderService.getCatalogAndAuthData();
      if (authData && authData.salesStaffList) {
        authData.salesStaffList.forEach(st => salesStaffList.push(st));
      }
    }

    // 3. Tự động nhận diện tài khoản Google đăng nhập để chọn sẵn Sales
    let defaultSalesPerson = '';
    try {
      const userEmail = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
      if (userEmail && typeof OrderService.getCatalogAndAuthData === 'function') {
        const authData = OrderService.getCatalogAndAuthData();
        if (authData.currentUser && authData.currentUser.display) {
          defaultSalesPerson = authData.currentUser.display;
        }
      }
    } catch (e) {
      defaultSalesPerson = '';
    }

    // 4. Đọc dữ liệu các dòng trên TAO DON và gom nhóm theo PO
    const rawData = sheet.getRange(startDataRow, 1, lastRow - startDataRow + 1, lastCol).getValues();
    const poGroupsMap = {};
    const today = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "yyyy-MM-dd");
    const defaultDelivery = Utilities.formatDate(new Date(Date.now() + 10 * 86400000), "Asia/Ho_Chi_Minh", "yyyy-MM-dd");

    for (let r = 0; r < rawData.length; r++) {
      const row = rawData[r];
      const prodName = c_sp >= 0 ? String(row[c_sp] || '').trim() : '';
      if (!prodName) continue;

      const custName = c_kh >= 0 ? String(row[c_kh] || '').trim() : '';
      let poNum = c_po >= 0 ? String(row[c_po] || '').trim() : '';
      if (!poNum) {
        const yymmdd = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "yyMMdd");
        const cleanCus = custName.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/Đ/g, "D").toUpperCase().replace(/[^A-Z0-9]+/g, "-").replace(/^-+|-+$/g, "");
        poNum = 'PO-' + yymmdd + '-' + (cleanCus || 'KH');
      }

      const existingLink = c_link >= 0 ? String(row[c_link] || '').trim() : '';
      const orderNote = c_note >= 0 ? String(row[c_note] || '').trim() : '';

      if (!poGroupsMap[poNum]) {
        poGroupsMap[poNum] = {
          customerName: custName,
          soPO: poNum,
          autoPONumber: poNum,
          salesPerson: defaultSalesPerson,
          ngayDat: today,
          ngayGiaoDK: defaultDelivery,
          orderNote: orderNote,
          existingFileUrl: existingLink,
          items: []
        };
      }

      const codeSP = c_code >= 0 ? String(row[c_code] || '').trim() : '';
      const unit = c_unit >= 0 ? (String(row[c_unit] || 'Cái').trim() || 'Cái') : 'Cái';
      const qty = c_qty >= 0 ? (Number(row[c_qty]) || 0) : 0;
      const price = c_price >= 0 ? (Number(row[c_price]) || 0) : 0;
      const cached = cacheProductMap[prodName.toLowerCase()] || {};

      poGroupsMap[poNum].items.push({
        prodName: prodName,
        codeSP: codeSP,
        unit: unit,
        qty: qty > 0 ? qty : 5000,
        price: price > 0 ? price : (cached.price || 0),
        giaKhuon: 0,
        mang: '',
        type: '1',
        donggoi: '',
        ngayDat: today,
        ngayGiao: defaultDelivery,
        deliveryDate: defaultDelivery,
        note: orderNote,
        isNewProduct: !Boolean(cacheProductMap[prodName.toLowerCase()])
      });
    }

    return JSON.stringify({
      poGroups: Object.values(poGroupsMap),
      salesStaffList: salesStaffList
    });

  } catch (err) {
    console.error("❌ [BACKEND] Lỗi getOrderFormDataForModal:", err);
    return JSON.stringify({ poGroups: [], salesStaffList: [] });
  }
}

/**
 * HÀM CẦU NỐI MỞ MODAL ĐƯỢC CLIENT GỌI QUA THƯ VIỆN BACKEND
 */
function openOrderModal(optSourceSpreadsheetId) {
  const targetId = optSourceSpreadsheetId || (SpreadsheetApp.getActiveSpreadsheet() ? SpreadsheetApp.getActiveSpreadsheet().getId() : null);
  
  // 1. Chuẩn bị dữ liệu từ sheet TAO DON và mã hóa Base64
  const jsonDataStr = getOrderFormDataForModal(targetId);
  const base64Data = Utilities.base64Encode(Utilities.newBlob(jsonDataStr, 'application/json').getBytes());

  // 2. Nạp dữ liệu thẳng vào Template HTML (KHÔNG CẦN CHỜ RPC FALLBACK)
  const template = HtmlService.createTemplateFromFile('OrderModal');
  template.initialDataJson = base64Data;

  const htmlOutput = template.evaluate()
    .setWidth(1150)
    .setHeight(700)
    .setTitle('NAPACO OMS - Tạo Đơn Hàng');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'NAPACO OMS - Phiếu Tạo Đơn Hàng');
}

// 3. Hàm xử lý onEdit tự động điền sản phẩm
function onEditClientHandler(e) {
  onEditAutoFillProduct(e);
}

/**
 * Hàm nhận giao dịch lưu đơn hàng
 */
function saveOrderFromModal(payload) {
  console.log('[BACKEND] >>> Bắt đầu xử lý saveOrderFromModal:');
  console.log('[BACKEND] Payload nhận được:', JSON.stringify(payload).substring(0, 300) + '...');

  if (typeof OrderService !== 'undefined' && typeof OrderService.saveOrderTransaction === 'function') {
    const res = OrderService.saveOrderTransaction(payload);
    console.log('[BACKEND] <<< Lưu thành công, kết quả:', JSON.stringify(res));
    return res;
  } else if (typeof executeSaveOrdersWithSchema === 'function') {
    const res = executeSaveOrdersWithSchema(payload);
    console.log('[BACKEND] <<< Lưu thành công qua executeSaveOrdersWithSchema:', JSON.stringify(res));
    return res;
  } else {
    throw new Error('Không tìm thấy hàm xử lý ghi đơn hàng (saveOrderTransaction / executeSaveOrdersWithSchema)!');
  }
}

function test_refreshOrderData_Backend() {
  const masterSS = SpreadsheetApp.openById(CONFIG.ID_MASTER);
  
  // IN RA TẤT CẢ TÊN SHEET ĐANG CÓ TRONG 01. DB MASTER ĐỂ ĐỐI CHIẾU
  const allSheetNames = masterSS.getSheets().map(s => s.getName());
  console.log("🔍 [DEBUG-MASTER] Danh sách tên Sheet trong DB MASTER:", JSON.stringify(allSheetNames));

  // Thử mở sheet nhân viên theo cấu hình
  const empSheet = masterSS.getSheetByName(CONFIG.SHEET_NAMES.EMPLOYEES);
  if (!empSheet) {
    console.error("❌ [DEBUG-MASTER] KHÔNG TÌM THẤY SHEET:", CONFIG.SHEET_NAMES.EMPLOYEES);
  } else {
    console.log("✅ [DEBUG-MASTER] Tìm thấy sheet:", empSheet.getName(), "- Số dòng:", empSheet.getLastRow());
    console.log("🔍 [DEBUG-MASTER] Dòng tiêu đề (Row 1):", JSON.stringify(empSheet.getRange(1, 1, 1, empSheet.getLastColumn()).getValues()[0]));
    if (empSheet.getLastRow() >= 2) {
      console.log("🔍 [DEBUG-MASTER] Dòng dữ liệu đầu tiên (Row 2):", JSON.stringify(empSheet.getRange(2, 1, 1, empSheet.getLastColumn()).getValues()[0]));
    }
    // In ra các dòng từ 4 đến 7 để xác định vị trí tiêu đề và dòng dữ liệu đầu tiên
  for (let r = 4; r <= 8; r++) {
    const rowVals = empSheet.getRange(r, 1, 1, empSheet.getLastColumn()).getValues()[0];
    console.log(`🔍 [ROW-${r}]:`, JSON.stringify(rowVals));
  }
  }

  // Chạy đồng bộ như cũ
  const TEST_CLIENT_SPREADSHEET_ID = '1iOJq_bxA6c8Oky1hsH-7ukiyT8YyGCLLgJ8AJ1Om5bE';
  const result = OrderService.syncCacheCatalogFromDatabase(TEST_CLIENT_SPREADSHEET_ID);
  console.log("✅ [TEST-BACKEND] Kết quả:", JSON.stringify(result));
}

/**
 * RPC: Tìm kiếm tệp trên Google Drive của người dùng theo từ khóa
 * @param {string} keyword Từ khóa tìm kiếm (tên file, số PO, tên khách...)
 * @return {{success: boolean, files: Array<{id: string, name: string, url: string, mimeType: string, size: number, updated: string}>}}
 */
function searchUserDriveFiles(keyword) {
  try {
    const kw = String(keyword || '').trim();
    if (!kw) return { success: true, files: [] };
    
    // Tìm kiếm tệp không nằm trong thùng rác
    const safeKw = kw.replace(/'/g, "\\'");
    const query = "title contains '" + safeKw + "' and trashed = false";
    const fileIterator = DriveApp.searchFiles(query);
    
    const results = [];
    let count = 0;
    while (fileIterator.hasNext() && count < 20) {
      const file = fileIterator.next();
      results.push({
        id: file.getId(),
        name: file.getName(),
        url: file.getUrl(),
        mimeType: file.getMimeType(),
        size: file.getSize(),
        updated: Utilities.formatDate(file.getLastUpdated(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy")
      });
      count++;
    }
    return { success: true, files: results };
  } catch (err) {
    console.error("❌ Lỗi searchUserDriveFiles:", err);
    return { success: false, error: err.message, files: [] };
  }
}