# Tồn Kho NVL

Mã đã được cập nhật trực tiếp trong dự án GSHEET APP. Chưa push hoặc triển khai Apps Script.

## Tệp thay đổi

- src/Config.js: cấu hình bảng KHO-FM-01B_CAPTRA và cột giao dịch.
- src/DonHang/OrderController.js: route film-stock, menu Tồn Kho ngay dưới Nhận Màng.
- src/DonHang/OrderMenu_View.html: thông báo đỏ phiếu trả chờ xác nhận trên các trang dùng menu chung; kiểm tra khi mở trang, trở lại trang và mỗi 60 giây khi trang đang hiển thị.
- src/KHO_NVL/FilmStockService.js: quyền NVL_STOCK_VIEW, cộng giao dịch theo N_NVL_ID, nối DN_ID → VL_ID và sản phẩm từ chi tiết đơn hàng.
- src/KHO_NVL/FilmReturnService.js: cộng dồn kg nhật ký, tạo/cập nhật phiếu trả chờ và xác nhận nhận lại.
- src/KHO_NVL/FilmStock_View.html: bộ lọc độ dày/khổ min–max, NCC và Còn Tồn/Hết Tồn; bảng nhóm thông tin.

## Quy tắc

Tồn = So_KG_Nhan − tổng CẤP NVL + tổng TRẢ NVL có Nguoi_Nhan.
Phiếu trả chưa có Nguoi_Nhan chỉ xuất hiện trong thông báo, chưa cộng vào tồn.
Mỗi dòng là một lần nhập. Hết Tồn hiển thị tối đa 1.000 kết quả mới nhất sau khi lọc, theo ngày nhập giảm dần; cùng ngày theo thứ tự dòng nhận mới hơn. Số âm được đưa vào nhóm Hết Tồn và cảnh báo để đối chiếu, không bị ép thành 0.
Type lấy từ sản phẩm gắn với DN_ID qua 02.ChiTietDH: 1 = Định hình, 2 = Bế. Chưa gắn sản phẩm sẽ hiển thị rõ, không tự suy đoán.
Đọc bảng theo lô và tổng hợp bằng Map trong từng yêu cầu; không đọc từng phiếu riêng lẻ. Giới hạn 1.000 là giới hạn kết quả trả về, vẫn cần đọc giao dịch để xác định tồn chính xác.
Thủ kho cần quyền NVL_STOCK_VIEW trong 08.MaTranQuyen để xem tồn và nhận thông báo. Tất cả người có quyền này sẽ nhận thông báo.

## Trả tự động và xác nhận

Đã ánh xạ đầy đủ 10 cột KHO-FM-01B_CAPTRA theo tiêu đề cung cấp; không thêm cột.
Khối lượng nhật ký = Trong_Luong_bia × (SL_Bia_Thuc_Te_Dat + SL_NG_Trong_Ca) / 1.000 kg.
Hai số lượng phải là số nguyên không âm; nhập rõ 0 khi không phát sinh NG. Nhật ký thiếu trường sẽ chặn đồng bộ/xác nhận để tránh trả sai lượng.
Cộng nhật ký theo mã cấp gốc và các mã trả tự động liên quan. Ví dụ cấp 100 kg, dùng 30 kg → chờ trả 70 kg; dùng tiếp 20 kg qua mã trả chờ → cập nhật còn 50 kg. Cập nhật một phiếu chờ hiện có, không thêm phiếu mỗi dòng nhật ký. Khi dùng hết, phiếu chờ giữ lại với 0 kg và không hiện cảnh báo.
Mã phiếu trả tự động dùng tiền tố MV_TRA_AUTO_ kèm UUID. Liên kết lần cấp gốc lưu trong note của ô MV_NVL_ID với tiền tố NAPACO_AUTO_RETURN_V1:. Đây là ghi chú ô, không phải cột Ghi_Chu. Phải giữ note khi di chuyển/xuất nhập dữ liệu; không sao chép note sang giao dịch thủ công. Phiếu trả cũ chưa có liên kết không được tự suy đoán theo N_NVL_ID hoặc ID_Lenh_SX.
Màn Tồn Kho có danh sách toàn bộ phiếu trả chờ, độc lập bộ lọc tồn. Nút “Đã nhận đủ … kg” cần cả NVL_RECEIVE và NVL_STOCK_VIEW. Hệ thống tự ghi mã người đang đăng nhập vào Nguoi_Nhan; không nhận người xác nhận từ trình duyệt. Phải kiểm tra hàng thực tế trước khi bấm. Chưa hỗ trợ nhận một phần hoặc ghi nhận chênh lệch kg.
Trước khi xác nhận phiếu tự động, hệ thống tính lại nhật ký; nếu kg khác số đang hiển thị sẽ yêu cầu tải lại. Sau xác nhận, sử dụng vượt số kg còn lại hoặc sửa nhật ký làm sai lệch lượng đã trả sẽ bị chặn để đối chiếu.
Thông báo dùng đường đọc riêng: chỉ quyền nhân viên và bảng cấp/trả, không tải toàn bộ dữ liệu tồn mỗi 60 giây.

## Điểm tích hợp Nhật ký Hút còn thiếu

Dự án nguồn chưa có trình lưu hoặc màn hình 05.NHAT_KY_HUT. Vì vậy chưa tự chạy sau khi nhập nhật ký trong ứng dụng thực tế.
Đã có hàm xử lý sẵn `syncFilmReturnFromJournal(movementId)` với quyền PRODUCTION_EDIT, có thể gọi lại để đồng bộ sau một lần lưu thành công.
Khi bổ sung trình lưu nhật ký, phải dùng cùng ScriptLock: lấy khóa, kiểm tra dữ liệu, lưu nhật ký, gọi `filmSyncReturnLocked_(operation, movementId, userId)`, rồi nhả khóa trong finally. Cần xử lý trường hợp nhật ký đã lưu nhưng đồng bộ trả lỗi bằng thông báo và cho thử lại, không lưu trùng nhật ký. Nếu sửa mã MV_NVL_ID trên dòng đã lưu phải đồng bộ cả mã cũ và mới; phiên bản ban đầu nên khóa trường này sau lưu. Các đường ghi cấp/trả khác cũng cần dùng cùng khóa. Không dùng trigger onEdit để thay thế vì cập nhật bằng Apps Script không kích hoạt trigger này.
Tên cột nhật ký đang dùng đúng theo trao đổi: MV_NVL_ID, Trong_Luong_bia, SL_Bia_Thuc_Te_Dat, SL_NG_Trong_Ca.

## Kiểm tra

node tests/stock.test.cjs
node tests/returns.test.cjs

Đã kiểm tra công thức nhận/cấp/trả, nhiều giao dịch cùng lần nhập, loại trừ trả chờ, cộng trả sau xác nhận, lọc min/max, dữ liệu số sai, Type và giới hạn 1.000 dòng. Đã kiểm tra thêm cộng dồn nhật ký, đổi gram sang kg, tái sử dụng phiếu trả chờ, chạy lại không tạo trùng, lưu liên kết note, chặn xác nhận khi kg thay đổi, chạy lại xác nhận không ghi trùng và nhả khóa khi lỗi. Kiểm tra cú pháp toàn bộ JavaScript và script trong HTML thành công. Chưa chạy với dữ liệu Google Sheets thật hoặc xác minh giao diện trên web app.
