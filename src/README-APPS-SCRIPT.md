# Pull, push và triển khai Apps Script

Chạy các lệnh trong thư mục chứa `.clasp.json`. Tệp này nằm trong `src/`, còn thư mục dự án nằm ngay phía trên `src/`.

```bash
cd "$HOME/Library/CloudStorage/GoogleDrive-hthanhnpc@gmail.com/My Drive/GSHEET APP"
```

## Kiểm tra mã trên máy đã đồng bộ chưa

```bash
bash "$HOME/Documents/ChatGPT/Bế Dán - Giải Quyết Lỗi/kiem-tra-apps-script.sh"
```

Lệnh so sánh `src/` với mã hiện tại trong trình soạn thảo Apps Script. Nếu có chênh lệch, xem danh sách tệp để quyết định cần pull hay push.

## Pull mã từ Apps Script về máy

```bash
git status --short -- src
npx --yes @google/clasp@3 pull
```

Nếu có thay đổi trên máy cần giữ, hãy commit hoặc sao lưu trước khi pull vì lệnh có thể ghi đè các tệp tương ứng trong `src/`.

## Push mã từ máy lên Apps Script

```bash
npx --yes @google/clasp@3 show-file-status
npx --yes @google/clasp@3 push
```

Trước khi push, kiểm tra danh sách có đủ các tệp cần giữ, gồm `appsscript.json` và các tệp trong `DonHang/`. Push cập nhật toàn bộ mã trong trình soạn thảo Apps Script theo tập tệp gửi đi.

## Tạo phiên bản và cập nhật URL web app

Sau khi push thành công, xem danh sách deployment:

```bash
npx --yes @google/clasp@3 list-deployments
```

Thay `DEPLOYMENT_ID_CUA_WEB_APP` bằng ID của deployment đang dùng:

```bash
npx --yes @google/clasp@3 update-deployment "AKfycbycZrEyqRYs3SShQyJkXMkvIJsyE7j-NV6cmJRqpbS2lVfvGeohyiS91rty42lapn-n" \
  --description "Cap nhat ma nguon"
```

Lệnh tạo phiên bản mới và cập nhật deployment đã chọn. Mở lại cùng URL `/exec` để kiểm tra.

## Ghi nhớ

- `pull` tải mã từ Apps Script về máy.
- `push` đưa mã từ máy lên trình soạn thảo Apps Script.
- `update-deployment` đưa phiên bản mới lên URL `/exec` hiện tại.
- Giữ `appsscript.json` ngay dưới `src/`. Tệp hướng dẫn Markdown được bỏ qua theo cấu hình mặc định của clasp.
