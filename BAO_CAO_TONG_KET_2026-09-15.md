# Báo cáo tổng kết đợt chỉnh sửa phần mềm tính Tổng mức đầu tư

**Ngày:** 15/09/2026
**Dành cho:** người sử dụng phần mềm (không cần hiểu về lập trình)

---

## 1. Về bản chất: công trình thủy lợi thuộc loại "Nông nghiệp & PTNT"

Trước đây, khi nhập các công trình **thủy lợi** (kênh mương, đê kè, hồ đập, trạm bơm...),
phần mềm không nhận diện được chúng thuộc nhóm "Nông nghiệp & PTNT", nên nhiều khoản chi phí
bị tính ra **0 đồng** hoặc sai tỷ lệ, gồm:

- Chi phí quản lý dự án (QLDA)
- Chi phí tư vấn (thiết kế, giám sát, thẩm tra, thẩm định thiết kế và dự toán)
- Chi phí chung CPC/TT/TNCTTT và chi phí nhà tạm khu vực công trình
- Phí thẩm định thiết kế, phí thẩm định dự toán

**Đã sửa:** toàn bộ phần mềm giờ tự hiểu "công trình thủy lợi" = "Nông nghiệp & PTNT".
Dù người dùng nhập loại công trình theo kiểu nào (thủy lợi, kênh mương, trạm bơm, nông nghiệp...),
phần mềm đều quy về đúng nhóm để tra đúng tỷ lệ và cho kết quả đúng, khác 0.

Kết quả kiểm tra mẫu (công trình thủy lợi cấp III, chi phí xây dựng 10 tỷ đồng):

| Khoản mục | Tỷ lệ trước | Tỷ lệ sau khi sửa |
|---|---|---|
| Quản lý dự án (QLDA) | 0% | 3,117% |
| Chi phí tư vấn thiết kế | 0% | 2,80% |
| Chi phí chung CPC | 0% | 6,1% |
| Thu nhập chịu thuế tính trước (TNCTTT) | 0% | 5,5% |
| Chi phí nhà tạm khu vực công trình | 0% | 5,5% |
| Phí thẩm định thiết kế | 0% | 0,121% |
| Phí thẩm định dự toán | 0% | 0,117% |

## 2. Phí thẩm định dự án: sửa cách tính theo đúng quy định

- **Trước đây:** phần mềm tính phí thẩm định dự án dựa trên (GXD + GTB) là sai.
- **Đã sửa:** phí thẩm định dự án giờ được tính dựa trên **Tổng mức đầu tư của dự án**
  (gồm cả bồi thường GPMB, quản lý dự án, tư vấn...), tra mức phí theo đúng bảng
  **"32. Tham dinh du an"** của Bộ Tài chính mà khách hàng cung cấp.
- Phí thẩm định thiết kế và phí thẩm định dự toán vẫn tra theo đúng bảng 33 và 34.

## 3. Công thức chi phí nhà tạm: giữ nguyên, không sửa

Công thức nhà tạm `LT = GXDTT × Tỷ lệ × (1 + TGTGT)` là **đúng** theo văn bản đính chính
của Thông tư 36/2026. Vì vậy chúng tôi **không thay đổi** công thức này.

## 4. Số liệu lập dự toán: bỏ heuristics gây sai

- **Trước đây:** khi đọc file Excel, phần mềm phát hiện khối lượng = 1 thì tự gán giá trị
  bằng 1 (bỏ qua dữ liệu thực), gây sai lệch ở các cột thành tiền.
- **Đã sửa:** bỏ hoàn toàn quy tắc "khối lượng = 1 = thiếu dữ liệu". Số liệu nhập bao nhiêu
  thì phần mềm tính bấy nhiêu.

## 5. Định dạng phần trăm trong file Excel

- **Trước đây:** các cột tỷ lệ CPC/TT/TNCTTT khi xuất ra file Excel bị nhân thêm 100,
  gây sai lệch (ví dụ: 6,1% xuất thành 610%).
- **Đã sửa:** bỏ phép nhân 100 thừa, tỷ lệ xuất chính xác (6,1% = 0,061 trong ô Excel,
  hiển thị đúng là 6,10% do định dạng ô Excel là "%").

## 6. Báo cáo Thẩm định: tỷ lệ QLDA/CPC/TNCTTT lấy đúng từ CSDL Thông tư 38

- **Trước đây:** tỷ lệ QLDA/CPC/TNCTTT trong báo cáo bị hardcode cứng, không đúng theo loại
  công trình thực tế.
- **Đã sửa:** giờ tỷ lệ báo cáo lấy trực tiếp từ CSDL Thông tư 38 (bảng 1.1, 3.2, 3.6),
  khớp chính xác theo loại công trình (Dân dụng, Giao thông, Công nghiệp, NN&PTNT, Hạ tầng KT).

| Loại CT | QLDA đúng | CPC đúng | TNCTTT đúng |
|---|---|---|---|
| Dân dụng | 3,283% | 7,3% | 5,5% |
| Công nghiệp | 3,450% | 6,2% | 6,0% |
| Giao thông | 2,951% | 6,2% | 6,0% |
| Nông nghiệp & PTNT | 3,117% | 6,1% | 5,5% |
| Hạ tầng kỹ thuật | 2,787% | 5,5% | 5,5% |

## 7. Tổng cộng: sửa cách nhận diện dòng cuối

- **Trước đây:** phần mềm tìm dòng "TỔNG CỘNG" bằng cách so khớp chính xác, dẫn đến bỏ sót
  khi file Excel có thêm khoảng trắng hoặc chữ hoa khác.
- **Đã sửa:** nhận diện bằng `StartsWith` (bắt đầu bằng), chữ hoa không phân biệt.

## 8. Tốc độ: giảm lag khi thao tác trên lưới

### 8a. Tab Tổng hợp kinh phí (tỷ lệ CPC/TT/TNCTTT/GTGT/Nhà tạm)

- **Trước đây:** mỗi lần gõ số, phần mềm tính toán ngay lập tức gây lag.
- **Đã sửa:** thêm bộ đếm thời gian (debounce 400ms), chỉ tính khi ngừng gõ 400ms.

### 8b. Tab Đơn giá nhân công/máy thi công (giá nhiên liệu)

- **Trước đây:** mỗi lần thay đổi giá xăng/diezel/điện, phần mềm mở database tra cứu lại
  toàn bộ danh sách nhân công và máy thi công → rất chậm.
- **Đã sửa:** thêm debounce 500ms + lưu cache dữ liệu database trong suốt phiên làm việc,
  không mở database nhiều lần.

### 8c. Tab Gọi đơn giá: giảm thời gian ghi dữ liệu vào Excel

- **Trước đây:** mỗi dòng một lần gọi COM Object, gây chậm trên file lớn.
- **Đã sửa:** đọc batch (một lần) các cột A/B/C/D từ sheet, ghi batch (một lần) các cột
  A/C/D. Đồng thời giữ lại dữ liệu gốc các dòng không thay đổi để không bị xóa.

## 9. Sửa lỗi crash (NullReferenceException)

- **Tab Tổng hợp kinh phí:** sửa lỗi crash khi click vào cột bị ẩn hoặc header cột.
- **Tab Thẩm định đơn giá chi tiết:** sửa lỗi crash khi `DanhSachCongTac` rỗng,
  hoặc `DanhSachHaoPhi` rỗng.
- **Tab Gọi đơn giá (Ribbon):** sửa lỗi crash khi kết quả thẩm định rỗng.

## 10. Giá nhiên liệu mặc định

- **Trước đây:** khi không tìm thấy file cấu hình giá nhiên liệu, các ô nhập giá nhiên liệu
  bị bỏ trống → toàn bộ chi phí máy thi công = 0.
- **Đã sửa:** giá mặc định: Xăng 22.150 đ/lít, Diezel 28.090 đ/lít, Điện 2.204 đ/kWh.

## 11. Kiểm tra chất lượng

- Phần mềm biên dịch thành công, không có lỗi kỹ thuật (0 lỗi, cảnh báo sẵn có không ảnh hưởng chức năng).
- Đã bổ sung và chạy **50 bài kiểm tra tự động**, tất cả đều đạt:
  - Kiểm tra loại công trình thủy lợi cho kết quả giống hệt loại "Nông nghiệp & PTNT".
  - Kiểm tra tỷ lệ và mốc phí thẩm định (dự án, thiết kế, dự toán) khớp chính xác 3 file Excel của Bộ Tài chính.
  - Kiểm tra các trường hợp giới hạn (phí tối thiểu 500.000 đ, tối đa 150.000.000 đ) vẫn hoạt động đúng.
  - Kiểm tra hồi quy: tỷ lệ QLDA/CPC/TNCTTT mặc định khớp chính xác CSDL Thông tư 38/2026.
  - Kiểm tra hồi quy: K_TD_DA được nội suy lại đúng khi sửa lưới, không làm thay đổi các mục khác.
  - Kiểm tra Bảng 3.7 nhà tạm: tại mốc quy mô ≤15 tỷ trả đúng 1,1% (công trình còn lại) và
    2,2% (công trình theo tuyến).

## 12. Hộp chọn "Hạng mục tính toán": bỏ tên giả gây nhầm lẫn

- **Trước đây:** khi dự toán không có tiêu đề hạng mục rõ ràng, phần mềm tự đặt tên hạng mục
  tạm là **"Hạng mục chung (Dân dụng)"** → hộp chọn hiển thị cái tên giả này, khiến người dùng
  tưởng công trình là dân dụng trong khi thực tế có thể là giao thông, cấp thoát nước...
- **Đã sửa:** hạng mục tạm sẽ hiển thị **"(Chưa phân loại)"** và **không gán loại công trình**,
  nói rõ "đây là hạng mục chưa được xác định loại", tránh gán sai định mức.

## 13. Tỷ lệ chi phí nhà tạm (Bảng 3.7) tính đúng theo quy mô thực tế

- **Trước đây:** quy mô chi phí xây dựng dùng để tra Bảng 3.7 lấy từ ô nhập tay; ô này có thể
  còn giữ giá trị cũ (như 40 tỷ) từ lần thao tác trước → dự án có chi phí xây dựng trước thuế
  ≤15 tỷ vẫn bị tính tỷ lệ **1,07%** (thay vì **1,1%**) với công trình còn lại, hoặc
  **2,14%** (thay vì **2,2%**) với công trình theo tuyến.
- **Đã sửa:** quy mô giờ được **tự động lấy từ chi phí xây dựng trước thuế thực tế** của dự toán
  (tỷ đồng), hiển thị sẵn vào ô nhập và dùng để tra chính xác Bảng 3.7:
  - Chi phí xây dựng trước thuế ≤15 tỷ → 1,1% (công trình còn lại) / 2,2% (theo tuyến).
  - Các mốc lớn hơn vẫn nội suy tuyến tính đúng theo Thông tư 36/2026.
- **Bổ sung:** lỗi tương tự vẫn còn ở giai đoạn **"Lập Báo cáo kinh tế - kỹ thuật"** — ở giai đoạn
  này ô quy mô bị khóa hiển thị "≤ 40" nên phần mềm lại dùng mốc 40 tỷ để tra Bảng 3.7, cho ra
  **1,07%** (thay vì 1,1%) và **2,14%** (thay vì 2,2%).
  Đã sửa: kể cả ở giai đoạn Báo cáo KT-KT, tỷ lệ nhà tạm vẫn được tra theo **quy mô thực tế**
  của dự toán (không dùng mốc "≤ 40").

---

*Ghi chú kỹ thuật (nội bộ): các tệp đã sửa gồm*
`DinhMucTT38Database.cs`, `DinhMucTT38Engine.cs` (thêm `CapNhatTiLeThamDinhDuAn`),
`InterpolationHelper.cs`, `ChiPhiKinhPhiItem.cs`,
`TongHopKinhPhiForm.cs`, `XuatBaoCaoThamDinhForm.cs`, `XuatBangBieuService.cs`,
`DuToanExcelReader.cs`, `ThamDinhExcelWriter.cs`,
`AieRibbon.cs`, `ThamDinhDonGiaForm.cs`, `ThamDinhDuToanChiTietForm.cs`,
`LapDuToanExcelService.cs` (tên hạng mục "(Chưa phân loại)").
*Bản sửa chưa được commit vào kho mã nguồn theo yêu cầu của người dùng.*
