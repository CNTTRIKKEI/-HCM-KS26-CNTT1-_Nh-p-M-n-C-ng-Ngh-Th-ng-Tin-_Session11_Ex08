# PHẦN 1: PHÂN TÍCH LỖI VĂN BẢN & CƠ CHẾ AUTOFILTER
## 1. Danh sách các lỗi vi phạm Quy tắc Vàng Database
|  Ô/Hàng vi phạm | Quy tắc vi phạm | Bản chất lỗi và hậu quả đối với Autofilter |
| --- | --- | ---|
| Hàng A1:E1 | Quy tắc 1: Hàng đầu tiên = tên cột, không Merge/chèn tiêu đề trang trí. | Chèn tiêu đề lớn "DANH SÁCH BÁO CÁO VIÊN..." ở ngay trên bảng làm Excel/Sheets nhầm đây là hàng Header, dẫn đến AutoFilter không nhận diện đúng các tên cột ở hàng 2 (STT, Họ và Tên,...). |
| Hàng nằm giữa (sau hàng 4) | Quy tắc 5: Tổng/Average đặt ngoài hoặc tách biệt khỏi vùng dữ liệu. | Chèn dòng tính tổng ngắt quãng TỔNG SỐ CNTT: 45 ngay giữa bảng dữ liệu. Hàng này biến thành "vách ngăn", làm AutoFilter ngắt vùng lọc (chỉ lọc được dữ liệu phía trên dòng tổng, bỏ qua các dòng phía dưới như của Lê Minh Châu). |
| Ô B3(Trần Thị Bình) | Quy tắc (Độ sạch & Chuẩn hóa dữ liệu) | Chứa khoảng trắng vô hình ở cuối chuỗi. Dù mắt thường thấy giống nhau, Filter/Hàm đếm sẽ tách thành 2 đối tượng khác nhau. |
| Toàn bộ cấu trúc | Quy tắc 4: Không được có hàng/cột trống ngắt quãng giữa Data. | Khi chèn dòng tổng ngắt quãng, dữ liệu bị đứt đoạn, khiến phím tắt Ctrl + Shift + L không quét hết toàn bộ 120 báo cáo viên. |
## 2. Giải thích cơ chế so sánh chuỗi ký tự của Filter (Lỗi khoảng trắng)
 * Cơ chế so sánh theo mã ASCII/Unicode: Bảng tính (Excel/Google Sheets) kiểm tra tính tương đồng của hai chuỗi ký tự bằng cách so sánh từng mã ký tự (ASCII/Unicode) theo thứ tự từ trái sang phải, bao gồm cả ký tự khoảng trắng (Spacebar - Mã ASCII 32).
 * Vì sao Filter không khớp:
   - Chuỗi A: "Trần Thị Bình" --> Độ dài = 14 ký tự.
   - Chuỗi B: "Trần Thị Bình  " --> Độ dài = 16 ký tự (bao gồm 2 khoảng trắng phía sau).
   Kết quả: Hệ thống nhận diện hai chuỗi này có độ dài và chuỗi mã hóa khác hẳn nhau. Do đó, khi bạn chọn lọc theo từ khóa "Trần Thị Bình", bộ lọc sẽ tự động bỏ qua chuỗi "Trần Thị Bình  ".
# PHẦN 2: HƯỚNG DẪN SỬA LỖI & NGHIỆM THU
## 3. Cấu trúc bảng dữ liệu chuẩn hóa (Clean Database Structure)
Để AutoFilter hoạt động chuẩn xác trên toàn bộ 120 báo cáo viên, bảng tính cần được tái cấu trúc theo nguyên tắc: Tách tiêu đề lớn ra ngoài, Header nằm ở ngay hàng 1, dữ liệu liên tục từ hàng 2 đến hàng 121, dòng tổng đặt hoàn toàn ở cuối/bên ngoài bảng.
Sơ đồ cấu trúc bảng chuẩn :
| [Hàng 1] | STT | Họ và Tên | Khoa / Viện | Điểm xét duyệt | Trạng thái | 
| --- | --- | --- | --- | --- | --- |
| [Hàng 2] | 1 | Nguyễn Văn An | CNTT | 8.5 | đã duyệt |
| [Hàng 3] | 2 | Trần Thị Bình | CNTT | 7.9 | Chờ duyệt |  
| [Hàng 4] | 3 | Lê Minh Châu | Điện Tử | 8.2 | đã duyệt |
| ... | ... | ... | ... | ... |

| [Hàng 123] | Tổng số CNTT: | =COUNTIF(C2:C121, "CNTT" |
| --- | --- | --- | 
Xử lý các yếu tố sai lệch:

  1. Tiêu đề báo cáo: Chuyển lên các hàng phía trên bảng (nếu cần) hoặc đặt tên Sheet là DS_BaoCaoVien.

  2. Dòng tổng ngắt quãng: Xóa bỏ triệt để khỏi giữa bảng, chuyển xuống ô A123 nằm độc lập bên dưới bảng dữ liệu.

  3. Tên cột (Header): Đặt cố định ở ngay Hàng 1 (A1:E1), tuyệt đối không Merge ô.
## 4. Công thức hàm TRIM() & Thao tác lọc đa điều kiện
Bước 1: Làm sạch cột "Họ và Tên" bằng hàm TRIM()
 1. Chèn thêm một cột phụ bên cạnh cột B (Cột C mới).

 2. Tại ô C2, nhập công thức: =TRiM(B2)
 3. Sao chép (Flash Fill / Drag) công thức từ C2 xuống hết C121.
 4. Bôi đen toàn bộ vùng dữ liệu cột C $\rightarrow$ Nhấn Ctrl + C --> Nhấp chuột phải vào cột B --> Chọn Paste Special --> Values Only (Dán giá trị thuần để bỏ khoảng trắng thừa).
 5. Xóa cột phụ vừa tạo.
Bước 2: Thao tác lọc danh sách báo cáo viên CNTT có điểm ≥ 8.0
 1. Kích hoạt AutoFilter trên toàn bảng:

* Chọn một ô bất kỳ bên trong bảng dữ liệu đã chuẩn hóa (ví dụ ô A1).

* Nhấn tổ hợp phím Ctrl + Shift + L (hoặc vào menu Data 
→ Create a Filter). Thanh Filter sẽ tự động gắn vào hàng Header A1:E1.

 2. Thiết lập Điều kiện 1 (Khoa = CNTT):

* Nhấp vào biểu tượng Filter ở cột Khoa / Viện (Cột C).

* Bỏ chọn (Select All) → Tích chọn duy nhất giá trị CNTT → Nhấn OK.

 3. Thiết lập Điều kiện 2 (Điểm xét duyệt ≥ 8.0):

* Nhấp vào biểu tượng Filter ở cột Điểm Xét Duyệt (Cột D).

* Chọn Filter by condition (Lọc theo điều kiện) → Chọn Greater than or equal to (Lớn hơn hoặc bằng).

Nhập giá trị: 8.0 → Nhấn OK.
> Kết quả nghiệm thu: Bảng tính thực hiện phép lọc giao (AND) giữa hai điều kiện và trả về đúng danh sách dự kiến (~35 báo cáo viên thuộc khoa CNTT đạt điểm từ 8.0 trở lên).

