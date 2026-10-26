# BÀI 6: HOÀN THIỆN BỘ HỒ SƠ TRIỂN KHAI NỀN TẢNG SMART-HOME

## 1. Giới thiệu dự án

Smart-Home là dự án nền tảng bất động sản trực tuyến, hỗ trợ tìm kiếm nhà đất, quản lý giá thuê, tra cứu hoa hồng môi giới và tham gia đấu giá trực tuyến.

Bộ hồ sơ dự án gồm 3 file chính: Word, Excel và PowerPoint.

## 2. Phần 1 – Phân tích bảng Excel Xấu và Đúng

### 2.1. Đặc điểm bảng Excel Xấu

- Sử dụng Merge Cell trong vùng dữ liệu.
- Có dòng trống hoặc dòng tổng kết xen giữa các bản ghi.
- Tiêu đề cột không rõ ràng hoặc không thống nhất.
- Một cột chứa nhiều loại thông tin khác nhau.
- Dữ liệu thiếu mã định danh hoặc không cùng định dạng.

Những lỗi này khiến việc lọc, sắp xếp, tính toán và tra cứu dữ liệu trở nên khó khăn.

### 2.2. Đặc điểm bảng Excel Đúng

- Mỗi cột là một trường dữ liệu (Field).
- Mỗi dòng là một bản ghi (Record).
- Mỗi cột có một tiêu đề duy nhất.
- Không Merge Cell trong vùng dữ liệu.
- Không chèn dòng trống hoặc dòng tổng kết vào giữa bảng.
- Mã nhà đất và mã môi giới được sử dụng thống nhất.

Bảng dữ liệu chuẩn giúp Excel xử lý thông tin chính xác, thuận tiện cho việc lọc và tra cứu bằng công thức.

## 3. Phần 2 – Triển khai bộ hồ sơ Smart-Home

### 3.1. File Word – SRS_SmartHome.docx

Tài liệu SRS mô tả yêu cầu phần mềm của nền tảng Smart-Home.

Nội dung chính:
- Giới thiệu mục tiêu và phạm vi dự án.
- Xác định các nhóm người dùng.
- Mô tả chức năng đăng tin và tìm kiếm bất động sản.
- Mô tả chức năng tính giá thuê sau chiết khấu.
- Mô tả chức năng quản lý môi giới và đấu giá trực tuyến.
- Đưa ra yêu cầu phi chức năng về bảo mật, hiệu năng và dữ liệu.
- Xác định tiêu chí nghiệm thu hệ thống.

Các tiêu đề được định dạng bằng Heading 1, Heading 2 và Heading 3 để tạo cấu trúc tài liệu rõ ràng và hỗ trợ tạo mục lục tự động.

### 3.2. File Excel – QuanLyNhaDat.xlsx

File Excel dùng để quản lý danh sách nhà đất và tính toán hoa hồng môi giới.

**Sheet NhaDat**

Quản lý các trường: mã nhà đất, tên bất động sản, loại hình, giá thuê, chiết khấu, giá sau chiết khấu, mã môi giới và hoa hồng.

Công thức tính giá thuê sau chiết khấu:

    =D2*(1-E2)

Trong đó:
- D2 là giá thuê ban đầu.
- E2 là tỷ lệ chiết khấu.
- Kết quả là giá thuê sau khi áp dụng chiết khấu.

**Sheet CongThuc**

Lưu danh sách mã môi giới, tên môi giới và tỷ lệ hoa hồng để phục vụ tra cứu.

Công thức tính hoa hồng tại ô H2 của sheet NhaDat:

    =F2*VLOOKUP(G2,CongThuc!$A$2:$C$4,3,FALSE)

Trong đó:
- F2 là giá thuê sau chiết khấu.
- G2 là mã môi giới.
- CongThuc!$A$2:$C$4 là vùng dữ liệu tra cứu.
- 3 là số thứ tự cột lấy tỷ lệ hoa hồng.
- FALSE yêu cầu tìm kiếm chính xác mã môi giới.

Tỷ lệ hoa hồng trong bài tập là dữ liệu giả định phục vụ thực hành Excel.

### 3.3. File PowerPoint – SmartHome.pptx

Bộ Pitch Deck gồm 4 slide để giới thiệu dự án và thuyết phục nhà đầu tư.

1. Giới thiệu dự án Smart-Home.
2. Những vấn đề trong giao dịch bất động sản.
3. Giải pháp của Smart-Home.
4. Khách hàng mục tiêu.

Các slide áp dụng nguyên tắc Rule 10/20/30: không quá 10 slide, thời lượng thuyết trình khoảng 20 phút và cỡ chữ lớn, dễ đọc theo nguyên tắc 30 pt.

Mỗi slide tập trung vào một thông điệp chính, hạn chế chữ dài và sử dụng White Space để làm nổi bật nội dung quan trọng.

## 4. Phần 3 – Kiểm thử lỗi VLOOKUP

### 4.1. Nguyên nhân lỗi

Khi kéo công thức VLOOKUP xuống các dòng khác, vùng tra cứu có thể bị dịch chuyển nếu sử dụng tham chiếu tương đối.

Ví dụ, vùng `A2:C4` có thể thay đổi thành `A3:C5`, khiến công thức tra cứu sai hoặc không tìm thấy mã môi giới.

### 4.2. Cách khắc phục

Sử dụng dấu `$` để cố định vùng dữ liệu tra cứu:

    CongThuc!$A$2:$C$4

Công thức hoàn chỉnh:

    =F2*VLOOKUP(G2,CongThuc!$A$2:$C$4,3,FALSE)

Dấu `$` cố định cả cột và hàng của vùng tra cứu. Trong khi đó, các ô F2 và G2 vẫn tự thay đổi theo từng dòng khi kéo công thức xuống.

Tham số FALSE giúp tìm đúng mã môi giới, tránh trả về tỷ lệ hoa hồng không chính xác.

## 5. Tổng kết

Qua bài tập, em đã vận dụng kiến thức về chuẩn hóa dữ liệu Excel, sử dụng Field và Record đúng cách, tính giá thuê sau chiết khấu, tra cứu hoa hồng bằng VLOOKUP và cố định vùng dữ liệu bằng dấu `$`.

Bên cạnh đó, em đã xây dựng tài liệu SRS có cấu trúc bằng Heading Styles và thiết kế Pitch Deck theo nguyên tắc một slide – một thông điệp.

Ba file Word, Excel và PowerPoint phối hợp thành bộ hồ sơ cơ bản phục vụ việc chuẩn bị triển khai dự án Smart-Home.
