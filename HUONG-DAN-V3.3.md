# Hướng dẫn cài đặt V3.3 – Một Admin (Vũ Duy Khánh)

**Trạng thái:** Mã nguồn đã cập nhật trên GitHub. Chưa tự triển khai vào Google Apps Script hay Google Sheets của anh.

## Kiến trúc và quyền truy cập
- Website GitHub Pages công khai: https://mattransoag.github.io/hotrokythuat/
- Web App Apps Script: người gửi không cần tài khoản; chỉ gửi và tra cứu bằng mã phiếu + mã xác thực.
- **Google Sheets riêng tư là giao diện quản trị duy nhất.** Không chia sẻ Sheet/Drive cho người gửi. Không có dashboard hay danh sách hồ sơ công khai.
- Chỉ tài khoản Google của anh có quyền chỉnh sửa Sheet; không tạo tài khoản admin hay mật khẩu lưu trong GitHub.

## Cài đặt
1. Tạo Google Sheets mới, đặt tên `QUẢN LÝ HỖ TRỢ KỸ THUẬT MTTQ AN GIANG`.
2. Trong Sheet chọn **Tiện ích mở rộng → Apps Script**.
3. Dán toàn bộ `gas/Code.gs` vào `Code.gs`; tạo tệp HTML tên `App` và dán `gas/App.html`.
4. Lưu; chọn hàm **setup** → **Run** → cấp quyền Google Sheets, Gmail, Drive và trình kích hoạt.
5. Sau khi chạy, kiểm tra các tab **YeuCau**, **LichSu**, **CauHinh**, **TongQuan**. Một thư mục Drive riêng lưu ảnh minh chứng được tạo tự động.
6. Apps Script → **Deploy → New deployment → Web app**. Chọn **Execute as: Me** và **Who has access: Anyone** để cơ sở gửi yêu cầu không cần đăng nhập. Lưu ý đây là Web App công khai; không đặt hàm đọc toàn bộ hồ sơ trong giao diện công khai.
7. Sao chép URL kết thúc `/exec` vào `config.js` trong GitHub, trường `appsScriptUrl`. Không dùng URL của website mẫu.
8. Thử gửi yêu cầu với email của chính anh, kiểm tra mã phiếu/mã xác thực, email, ảnh Drive và tra cứu. Chỉ công bố rộng rãi sau khi kiểm thử.

## Thao tác mỗi ngày
- Mở tab **TongQuan** để xem tổng hồ sơ, chờ tiếp nhận, đang xử lý, hoàn thành.
- Mở tab **YeuCau**, lọc cột **L (TrangThai)** để xem hồ sơ chưa xử lý.
- Cột **M (KetQua)**: nhập kết quả; cột **S (HoanThanh)**: tích checkbox. Hệ thống đổi trạng thái, ghi lịch sử và gửi Gmail.
- Menu **MTTQ - Quản lý hồ sơ** có mẫu trả lời nhanh và lệnh hoàn thành dòng đang chọn.
- Cột **P (AnhMinhChung)**: liên kết ảnh do người gửi tải lên (riêng tư).
- Cột **T (AnhKetQua)**: anh có thể dán liên kết ảnh kết quả lưu trong Drive riêng; chưa có chức năng upload ảnh kết quả trực tiếp từ giao diện.
- Cột **R (GhiChuNoiBo)**: ghi chú của anh, không xuất hiện trong tra cứu công khai.

## Giới hạn và lưu ý an toàn
- Chưa triển khai hệ thống trên tài khoản Google của anh; GitHub không thể tự tạo Sheets hoặc cấp quyền Gmail.
- Chưa có giao diện admin web riêng; Sheet là bảng quản trị riêng tư, tránh mở API quản trị công khai.
- Apps Script Web App công khai có thể bị gửi yêu cầu rác. Trước khi sử dụng quy mô lớn nên bổ sung giới hạn tốc độ, chống spam, kiểm thử quyền truy cập và quota MailApp/Drive.
- Google Drive không tự công khai ảnh. Người tra cứu phiếu không xem được tệp ảnh trực tiếp trừ khi được chia sẻ riêng.
- Khi gửi thành công nhưng Gmail lỗi, mã phiếu và mã xác thực vẫn được hiển thị trên màn hình để người gửi lưu.
