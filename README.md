# FitRowXL HÀM TỰ ĐỘNG GIÃN DÒNG

[Click vào đây để tải xuống](https://github.com/SanbiVN/FitRowXL/releases/download/fit_row/FitRowXL_v2.31.xlsm)

[![Lượt tải](https://img.shields.io/github/downloads/SanbiVN/FitRowXL/total.svg)](https://github.com/SanbiVN/FitRowXL/releases/download/fit_row/FitRowXL_v2.31.xlsm) 


### Chức năng ưu việt:
- Co giãn dòng hoàn toàn tự động.
- Co giãn dòng kể cả các ô đã được gộp.
- Co giãn dòng với các giá trị nhiều ô gộp cùng dòng.
- Co giãn dòng kể cả chiều cao vượt giới hạn của Excel là 409.5
- Hoạt động cả ở chế độ Xem In Ấn.
- Cách gõ hàm cài đặt đối số tùy chỉnh ưu việt:
 1. Thêm chiều cao nhất định cho dòng đã giãn.
 2. Đặt chiều cao mặc định cho vùng trống.
 3. Đặt chiều cao mặc định cho dòng trống.
 4. Tự đặt tỉ lệ giãn chiều rộng, chiều cao và thụt đầu dòng, khi chiều cao dòng vượt giới hạn.

Vì dùng hàm UDF nên rất tối ưu, tiết kiệm CPU.
Chỉ cần gõ một biểu thức FITROW cho cả vùng cần co giãn.

## Hướng dẫn sử dụng hàm:

Hàm FITROW được viết theo phương pháp mới nên cách nhập đối số là gõ hàm như dưới đây:

Hàm cài đặt và bổ trợ	| Kiểu	| Chức năng
----------------------|------|----------
ff_Padding(Height) |	Số |	Tăng chiều cao thêm một số
ff_defaultHeight(Height)	| Số	| Chiều cao mặc định nếu giá trị rỗng, dễ hiểu, nếu co giãn vùng ô A1:C20, mà cả vùng đó rỗng, thì chỉnh về chiều cao mặc định.
ff_HeightOfRowNull(Height) |	Số	| Đặt chiều cao mặc định cho cả dòng rỗng (giãn vùng A1:Z20, dòng A2:Z2 rỗng)
ff_AllSheets() |	Có	| Giãn dòng kể cả vùng ở trang tính không hiện hành.
ff_AutoFit()	| Có	| Bật tự động Fit khi ô tham chiếu thay đổi giá trị
ff_Indexes(cell1,cell2,...)	| Vùng | chứa nhóm văn bản	Căn chỉnh biên bản ở chế độ PrintView, khi giãn dòng, chiều cao trang in có thể cao hơn hoặc thấp hơn, làm cho trang in bị xê dịch, nên cần điều chỉnh để phù hợp.
ff_Scale(scaleWidth,scaleHeight,indentWidth)		| | Đặt tỉ lệ giãn chiều rộng, chiều cao và thụt đầu dòng, khi chiều cao dòng vượt giới hạn
​
Ví dụ: giãn dòng A1 và đối số, gõ =FITROW(A1,ff_Padding(5)) ​
Các hàm với các ký tự đầu là ff_... Chính là các hàm cài đặt và bổ trợ cho hàm chính FITROW​
Ví dụ: gõ =FITROW(A1,B4,C5), sẽ co giãn các ô A1, B4, C5, các cài đặt là mặc định​

CÁC HÀM LỆNH TẠO NÚT VÀ BIỂU THỨC NHANH:

HÀM	| Chức năng
----------------------|----------------
=FITROW_AddFX()​ | Tạo nhanh biểu thức FITROW vào ô
=FITROW_AddFXPrintArea()​ | Tạo nhanh biểu thức FITROW vùng in vào ô
=FITROW_AddButton()​ | Tạo nút nhấn để giãn dòng
=FITROW_AddButtonPrintArea()​ | Tạo nút nhấn để giãn dòng vùng in
=FitRow_Off()​ | Tắt chế độ tự động giãn dòng
=FitRow_On()​ | Bật chế độ tự động giãn dòng


Viết hàm nhanh: =FITROW(A2:F1000)
Viết hàm có cài đặt đối số: =FITROW(A2:F1000,ff_defaultHeight(40),ff_Padding(5))
Cách nhập nhiều vùng cần co giãn dòng:
=FITROW(A1:C9,D2:F3,E5:E6)​
​
Phím tắt giãn dòng: CTRL+SHIFT+ALT+R

Các hàm Bổ trợ:
1. Gõ hàm FITROW_OFF: nếu đang chỉnh sửa trang tính hãy tắt chế độ co giãn dòng hoặc bật chế độ Design Mode trong Tab Developer.​
2. Gõ hàm FITROW_ON: Bật chế độ co giãn dòng tự động.​
3. Thủ tục FITROW_Toggle + Check box có tên là chxAutoFitRow dùng để bật tắt chế độ co giãn dòng nếu muốn (Ví dụ nằm ở Sheet1 trong tập tin đính kèm bên dưới).​
Bước 3 này là một thủ thuật để ngăn chặn code tính toán lúc ứng dụng vừa khởi động, vì có thể sẽ gặp phải tình trạng code sẽ làm chậm quá trình khởi động.​
​
Hãy để dòng code sau vào sự kiện Workbook_Open: Call FITROW_Off​
Hãy mở lại bằng bước 2 hoặc bước 3.​


****Lưu ý:
Code sẽ tạo trang tính ẩn có tên __CELLFIXING__ để giãn dòng.
Khi giãn dòng tự động chế độ Undo và Redo của trang tính sẽ không hoạt động.
Nếu trong trang tính có hàm giãn dòng, không nên sử dụng hàm RandBetween, và các hàm random.

***Mã có thể chưa được tối ưu nhất, nên có thể cập nhật lại nhiều lần, nên nếu các bạn có sử dụng code thì nên thường xuyên xem lại bài viết, sẽ có thông báo cập nhật nếu có ở đầu bài viết.
