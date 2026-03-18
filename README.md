# FitRowXL - Add-in tự động hóa giãn dòng cột cho excel
[![Lượt tải](https://img.shields.io/github/downloads/SanbiVN/FitRowXL/total.svg)](https://github.com/SanbiVN/FitRowXL/releases/download/fit_row/FitRowXL_v2.46.xlsm) 

### Phiên bảng mới 2026
- Sử dụng Ribbon để thiết lập và giãn dòng nhanh chóng
<img width="1480" height="180" alt="image" src="https://github.com/user-attachments/assets/7c86dda2-3ab8-4e24-87ad-3c4b76a6eeb9" />

# DANH MỤC
- [Tính năng mới](#tính-năng-mới)
- [TẢI XUỐNG](#tải-xuống)
- [HƯỚNG DẪN CÀI ĐẶT](#hướng-dẫn-cài-đặt)
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

# TẢI XUỐNG
<!-- items that need to be updated release to release -->
[ptUserAddin]: https://github.com/SanbiVN/FitRowXL/releases/download/v1.6/FitRowXL_v1.6.zip
[ptUserXlsm]: https://github.com/SanbiVN/FitRowXL/releases/download/fit_row/FitRowXL_v2.46.xlsm

|  Thông tin   | Tải xuống | Ghi chú |
|--------------|-----------|----------|
| FixRowXL Add-in | [FitRowXL_v1.6.zip][ptUserAddin] | Bản mới 2026 sử dụng Add-in Ribbon thiết lập giãn dòng nhanh chóng |
| FixRowXL gọi hàm | [FitRowXL_v2.46.xlsm][ptUserXlsm] | Bản dùng cho nhúng code trực tiếp vào tệp để gọi hàm   |

***Mật khẩu VBA là 1


# HƯỚNG DẪN CÀI ĐẶT

Tệp Add-in xlam để cài đặt vào Excel, sau khi cài đặt thì giao diện sử dụng hiển thị trên thanh Ribbon với tên ```TaxTCT```. Ứng dụng chỉ cần cài đặt một lần duy nhất, còn lại tự động kiểm tra và tải cài đặt phiên bản mới. \
Giải nén vào một thư mục được đặt tên phù hợp, sau khi giải nén, vào thông tin tệp ngoài thư mục bỏ unblock tệp trước khi cài đặt nếu có.
> <img width="377" height="389" alt="image" src="https://github.com/user-attachments/assets/e8cf3b18-41ab-433f-a873-b32b76e079de" /><img width="363" height="478" alt="image" src="https://github.com/user-attachments/assets/359bee94-f4b7-4fa2-bc48-23ab7723fd7b" />

**Cách 1:** 
- Mở trực tiếp Add-in hoặc nhấn chuột vào tệp để mở, trong Excel cần **```Enabled Macro```** để chương trình hoạt động. 
- Nếu chương trình chưa cài đặt khởi động cùng Excel, khi nhấn **BẮT ĐẦU** chương trình sẽ hỏi có cài đặt khởi động vào Excel không?

**Cách 2:** Thực hiện cài đặt Add-in bằng tay: 
  - Nếu chưa có tab Deverloper hiển thị trên thanh Ribbon (Thanh công cụ): nhấn chuột phải vào thanh Ribbon, chọn **```Customize the Ribbon```**.
  - Trong thẻ Deverloper chọn **```Excel Add-ins```**, sau đó chọn nút **```Browse...```** vào thư mục chứa tệp Add-in, đánh dấu Add-in vừa thêm và chọn nút OK 
  - Nếu đã cài đặt vào Excel, nhưng mỗi khi mở ứng dụng không thấy trên thanh Ribbon, thì vào **```Task Manager```** cần End Task ứng dụng Excel chạy ngầm.

 Nếu ứng dụng bị chặn không cho chạy macro thì hãy vào Cài đặt Excel, vào Trust Center, vào tạo đường dẫn thư mục an toàn cho thư mục chứa add-in tải về.
 


# Chức năng
- Co giãn dòng hoàn toàn tự động.
- Co giãn dòng kể cả các ô đã được gộp, bao gồm các vùng gộp xen kẻ bất đối xứng.
- Co giãn dòng với các giá trị nhiều ô gộp cùng dòng.
- Co giãn dòng kể cả chiều cao vượt giới hạn của Excel là 412.5
- Hoạt động cả ở chế độ Xem In Ấn vùng in đã Scale.

<img width="1031" height="355" alt="image" src="https://github.com/user-attachments/assets/124e4697-1c31-44d1-aae6-1001b375420a" />


# Thiết lập giãn dòng nhanh

<img width="563" height="116" alt="image" src="https://github.com/user-attachments/assets/190ebd7d-e68b-49f1-8792-afb0e2e0e783" />\

<img width="192" height="114" alt="image" src="https://github.com/user-attachments/assets/c6b64bb8-96a1-42f6-8860-fb9f4953c920" />

​​
> Phím tắt giãn dòng nhanh mặc định (có thể đổi): **CTRL+SHIFT+ALT+R**

Thiết lập một tên mới để lưu thiết lập để tái sử dụng về sau. Sau khi giãn dòng có thể nhấn hoàn tác nếu giãn dòng không như ý muốn.
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

<img width="185" height="114" alt="image" src="https://github.com/user-attachments/assets/67574e8a-d7e0-4b64-8713-d673968aa423" />

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

# Tự động hóa giãn dòng

<img width="204" height="112" alt="image" src="https://github.com/user-attachments/assets/fb8122cb-e5f9-47a7-84b1-3ff733561ff3" />

![fit rows and columns instantly](https://github.com/user-attachments/assets/3c661ef6-26cb-4511-b646-eaf9764ac9ec)

Mục đặt tự động hóa giãn cột và dòng tự động ngay tức thời, hành vi giãn dòng này không làm mất trạng thái chế độ Undo và Redo.
(Sử dụng lệnh gửi nhấn chuột vào tiêu đề để giãn dòng như thao tác tay)

### Các nút chức năng
- **Đặt cho vùng ô**: Kiểm tra và cập nhật add-in
- **Đặt cho cả trang tính**
- **Đặt từ hợp chọn**
- **Bật/tắt tự động hóa**


<!-- 
# Các hàm Bổ trợ:
1. Gõ hàm ```FITROW_OFF```: nếu đang chỉnh sửa trang tính hãy tắt chế độ co giãn dòng hoặc bật chế độ Design Mode trong Tab Developer.​
2. Gõ hàm ```FITROW_ON```: Bật chế độ co giãn dòng tự động.​
3. Thủ tục FITROW_Toggle + Check box có tên là chxAutoFitRow dùng để bật tắt chế độ co giãn dòng nếu muốn (Ví dụ nằm ở Sheet1 trong tập tin đính kèm bên dưới).​
Bước 3 này là một thủ thuật để ngăn chặn code tính toán lúc ứng dụng vừa khởi động, vì có thể sẽ gặp phải tình trạng code sẽ làm chậm quá trình khởi động.​ \
-->


# Các nút chức năng của add-in
<img width="305" height="114" alt="image" src="https://github.com/user-attachments/assets/33b9cfa4-4876-4396-8996-f67fa6c4f6c8" />

- **Cập nhật**: Kiểm tra và cập nhật add-in
- **Đặt lại cài đặt**
- **Thoát và gỡ cài đặt**
- **Hướng dẫn và nguồn**
- **Liên hệ Zalo**
​
# Lưu ý
Code sẽ tạo trang tính ẩn có tên ```__CELLFIXING__``` để giãn dòng. \
Sau khi giãn dòng tự động chế độ Undo và Redo của trang tính sẽ bị mất trạng thái. 
