# FitRowXL - Add-in tự động hóa giãn dòng cột cho excel
[![Lượt tải](https://img.shields.io/github/downloads/SanbiVN/FitRowXL/total.svg)](https://github.com/SanbiVN/FitRowXL/releases/download/fit_row/FitRowXL_v2.46.xlsm) 

### Phiên bảng mới 2026
- Sử dụng Ribbon để thiết lập và giãn dòng nhanh chóng
<img width="1491" height="189" alt="image" src="https://github.com/user-attachments/assets/a74bf7cb-062d-4648-9d1f-03662d4255a5" />

# DANH MỤC
- [Tính năng mới](#tính-năng-mới)
- [TẢI XUỐNG](#tải-xuống)
- [Chức năng](#chức-năng)
- [Thiết lập giãn dòng nhanh](#thiết-lập-giãn-dòng-nhanh)
- [Thiết lập giãn dòng nhiều vùng ô](#thiết-lập-giãn-dòng-nhiều-vùng-ô)
- [Lưu ý](#lưu-ý)
  
# Tính năng mới
- Thêm chế độ tự động cập nhật ứng dụng lên phiên bản mới, hoặc phục hồi phiên bản.
- Hỗ trợ giãn dòng trước khi in ấn tại mục thiết lập **Giãn nhiều vùng ô**:
> Khi tạo mới sẽ có nhập Macro gọi trước và sau khi giãn.\
Cũng có thể gọi Macro trong mã VBA của bạn với **Application.Run** "**FitRowAreas**", "**TenDaThietLap**"\
Để giãn dòng nhanh cho vùng ô không có ô Gộp, trong **Danh sách giãn dòng** hãy chọn đối tượng là Table

- Thêm thiết lập tự động giãn dòng và cột tức thời khi giá trị ô không gộp thay đổi (không làm mất chế độ Undo).

![fit rows and columns instantly](https://github.com/user-attachments/assets/3c661ef6-26cb-4511-b646-eaf9764ac9ec)

# TẢI XUỐNG
<!-- items that need to be updated release to release -->
[ptUserAddin]: https://github.com/SanbiVN/FitRowXL/releases/download/v1.3/FitRowXL_v1.3.zip
[ptUserXlsm]: https://github.com/SanbiVN/FitRowXL/releases/download/fit_row/FitRowXL_v2.46.xlsm

|  Thông tin   | Tải xuống | Ghi chú |
|--------------|-----------|----------|
| FixRowXL Add-in | [FitRowXL_v1.3.zip][ptUserAddin] | Bản mới 2026 sử dụng Add-in Ribbon thiết lập giãn dòng nhanh chóng |
| FixRowXL gọi hàm | [FitRowXL_v2.46.xlsm][ptUserXlsm] | Bản dùng cho nhúng code trực tiếp vào tệp để gọi hàm   |

***Mật khẩu VBA là 1

# Chức năng
- Co giãn dòng hoàn toàn tự động.
- Co giãn dòng kể cả các ô đã được gộp.
- Co giãn dòng với các giá trị nhiều ô gộp cùng dòng.
- Co giãn dòng kể cả chiều cao vượt giới hạn của Excel là 412.5
- Hoạt động cả ở chế độ Xem In Ấn vùng in đã Scale.


# Thiết lập giãn dòng nhanh

<img width="559" height="112" alt="image" src="https://github.com/user-attachments/assets/2d534ebd-e0e3-4844-9bbd-c03630556047" />\
​​
> Phím tắt giãn dòng nhanh mặc định (có thể đổi): **CTRL+SHIFT+ALT+R**

Thiết lập một tên mới để lưu thiết lập để tái sử dụng về sau.
Thiết lập các chỉ số giãn dòng như sau:

Giá trị	| Kiểu	giá trị | Chức năng
----------------------|------|----------
Đệm chiều cao |	Số |	Tăng chiều cao thêm một số
Chiều cao mặc định	| Số	| Chiều cao mặc định nếu giá trị rỗng, dễ hiểu, nếu co giãn vùng ô A1:C20, mà cả vùng đó rỗng, thì chỉnh về chiều cao mặc định.
Chiều cao dòng trống |	Số	| Đặt chiều cao mặc định cho cả dòng rỗng (giãn vùng A1:Z20, dòng A2:Z2 rỗng)
Tỉ lệ chiều rộng |	Số	| Đặt tỉ lệ giãn chiều rộng, Tăng giảm chiều rộng trước khi tính toán giãn dòng
Chiều cao vùng trống	| Số	| Nếu vùng dữ liệu là Table hãy nhập vào hàm này, để tăng tốc giãn dòng
Kiểu giãn dòng	| Tên	| Đặt kiểu giãn dòng cho các cột gộp ô

# Thiết lập giãn dòng nhiều vùng ô

<img width="180" height="111" alt="image" src="https://github.com/user-attachments/assets/9a07d321-1e4b-4b45-940d-233f4722e7ee"/>

Tính năng giãn nhiều vòng ô cho phép thiết lập nhiều vùng ô với các tùy chọn chỉ sổ riêng biệt, đồng thời hỗ trợ in ấn.
Cho phép tạo và lưu thiết lập để tái sử dụng, cho phép gọi trong dự án chứa mã VBA của bạn để thực hiện giãn dòng trước khi in ấn hoặc công việc khác.

### Các nút chức năng
- **Tạo mới**: Mở form tạo thiết lập và lưu thiết lập để tái sử dụng, cho phép gọi trong dự án chứa mã VBA của bạn để thực hiện giãn dòng trước khi in ấn hoặc công việc khác. Các thiết lập sẽ được lưu vào trong chính dự án Excel, dựa vào Name. Form gồm có các thiết lập:
  - Đặt thủ tục gọi trước và sau khi gian dòng
  - Đặt các vùng ô cần dịch chuyển vừa khít nằm trong trang sau khi giãn dòng cho vùng in.
  - Danh sách tạo các vùng ô và chỉ số.
  - Các nút nhấn giãn thử, xóa mục, tích chọn.
- **Sửa thiết lập**: Sửa thiết lập từ tên trong hộp chọn
- **Xóa thiết lập**: Xóa nếu không còn sử dụng lại thiết lập.
- **Xóa thiết lập**: Xóa nếu không còn sử dụng lại thiết lập.
- **Hộp tên thiết lập**: Để chọn mục đã thiết lập.
- **Giãn dòng**: Giãn dòng dựa vào tên thiết lập.

Các chỉ số thiết lập|
----------------------|
Đệm chiều cao 
Chiều cao mặc định
Chiều cao dòng trống
Tỉ lệ chiều rộng 
Chiều cao vùng trống	
Kiểu giãn dòng	

<!-- 
# Các hàm Bổ trợ:
1. Gõ hàm ```FITROW_OFF```: nếu đang chỉnh sửa trang tính hãy tắt chế độ co giãn dòng hoặc bật chế độ Design Mode trong Tab Developer.​
2. Gõ hàm ```FITROW_ON```: Bật chế độ co giãn dòng tự động.​
3. Thủ tục FITROW_Toggle + Check box có tên là chxAutoFitRow dùng để bật tắt chế độ co giãn dòng nếu muốn (Ví dụ nằm ở Sheet1 trong tập tin đính kèm bên dưới).​
Bước 3 này là một thủ thuật để ngăn chặn code tính toán lúc ứng dụng vừa khởi động, vì có thể sẽ gặp phải tình trạng code sẽ làm chậm quá trình khởi động.​ \
-->​
# Lưu ý
Code sẽ tạo trang tính ẩn có tên ```__CELLFIXING__``` để giãn dòng. \
Sau khi giãn dòng tự động chế độ Undo và Redo của trang tính sẽ bị mất trạng thái. 
