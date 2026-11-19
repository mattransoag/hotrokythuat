# Hệ thống hỗ trợ kỹ thuật MTTQ tỉnh An Giang

Website GitHub Pages: https://mattransoag.github.io/hotrokythuat/

## Phiên bản 2.0 (giao diện)

- Giao diện responsive xanh đậm
- Điều hướng Tổng quan / Gửi yêu cầu / Tra cứu / Quản trị
- Biểu mẫu giao diện có kiểm tra đầu vào, **vô hiệu hóa khi chưa có backend**
- Khu vực thống kê hiển thị dấu — thay vì số liệu giả
- Cấu hình đường dẫn Google Apps Script tại `config.js`

## Kết nối dữ liệu

1. Tạo Apps Script và Google Sheets **riêng của đơn vị**.
2. Triển khai Web App, kiểm thử phân quyền, truy cập và bảo mật.
3. Điền URL Web App dạng `https://script.google.com/macros/s/.../exec` vào `config.js`.
4. Khi cấu hình, mục Gửi yêu cầu hiển thị ứng dụng Apps Script qua iframe, mục Tra cứu và Quản trị mở Apps Script. Nếu Google chặn iframe, sử dụng liên kết trực tiếp.
5. Chỉ triển khai hệ thống chính thức sau khi hoàn thiện cơ chế xác thực, mã phiếu, Gmail, quản lý và bảo vệ dữ liệu.

**Chưa có backend:** biểu mẫu không gửi dữ liệu, thống kê không có số liệu thật. Không sao chép URL Apps Script của website mẫu vì sẽ dùng nhầm dữ liệu của hệ thống khác.

Phát triển phần mềm: Vũ Duy Khánh.
