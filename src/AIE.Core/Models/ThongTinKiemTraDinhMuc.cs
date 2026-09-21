using System;
using System.Collections.Generic;

namespace AIE.Core.Models
{
    /// <summary>
    /// Kết quả tra cứu định mức chi phí (TT38/2026) của một khoản mục cụ thể trên Bảng Tổng hợp kinh phí.
    /// </summary>
    public class ThongTinKiemTraDinhMuc
    {
        /// <summary>Mã khoản mục (VD: "TV_FS_KTKT", "G_QLDA")</summary>
        public string MaChiPhi { get; set; } = "";

        /// <summary>Tên khoản mục chi phí</summary>
        public string TenChiPhi { get; set; } = "";

        /// <summary>Tên bảng định mức áp dụng (VD: "Bảng 1.1 - QLDA", "Bảng 2.3 - Lập FS", "Bảng 2.19 - Thẩm tra TK")</summary>
        public string TenBang { get; set; } = "";

        /// <summary>Số trang trong PDF gốc (38.2026.TT-BXD-PL8.pdf) chứa bảng định mức</summary>
        public int SoTrangPdf { get; set; }

        /// <summary>Mảng mốc quy mô (tỷ đồng) của bảng định mức áp dụng</summary>
        public decimal[] MocQuyMo { get; set; } = Array.Empty<decimal>();

        /// <summary>Mảng tỷ lệ % định mức tại từng mốc (theo loại CT)</summary>
        public decimal[] TiLeDinhMuc { get; set; } = Array.Empty<decimal>();

        /// <summary>Quy mô hiện tại (tỷ đồng) dùng làm đầu vào nội suy</summary>
        public decimal QuyMoTy { get; set; }

        /// <summary>Tỷ lệ % nội suy theo quy mô (đã nội suy từ mốc)</summary>
        public decimal TyLeNoiSuy { get; set; }

        /// <summary>Hệ số điều chỉnh (k) đang áp dụng (VD: 1.0, 1.15, 1.55, 0.8, 1.2...)</summary>
        public decimal HeSo { get; set; } = 1.0m;

        /// <summary>Tỷ lệ % hiện tại đang dùng trên dòng chi phí (item.TyLePhanTram)</summary>
        public decimal TyLeHienTai { get; set; }

        /// <summary>Diễn giải cơ sở tính (VD: "G_XD + G_TB (trước thuế)", "G_XD", "G_TB")</summary>
        public string CoSoTinh { get; set; } = "";

        /// <summary>Loại công trình đã chuẩn hóa</summary>
        public string LoaiCongTrinh { get; set; } = "";

        /// <summary>Cấp công trình (cho bảng 2.7-2.16 thiết kế theo cấp)</summary>
        public string CapCongTrinh { get; set; } = "";

        /// <summary>Ghi chú thêm (VD: "Thiết kế 3 bước × 1.55", "TB ≥ 25% → k=1.2")</summary>
        public string GhiChu { get; set; } = "";
    }

    /// <summary>
    /// Ánh xạ tên/STT bảng định mức TT38 sang số trang trong PDF gốc.
    /// Nguồn: ánh xạ từ file pl8_text.txt (48 trang), bảng 1.1 - 2.26 + DD1 + HTKT.
    /// </summary>
    public static class BangToPdfPageMap
    {
        private static readonly Dictionary<string, int> _map = new(StringComparer.OrdinalIgnoreCase)
        {
            // --- CHƯƠNG 1: Chi phí tư vấn ---
            ["1.1"]  = 7,    // QLDA
            ["2.1"]  = 11,   // BC đề xuất chủ trương
            ["2.2"]  = 12,   // BCNC tiền khả thi
            ["2.3"]  = 13,   // BCNC khả thi (FS)
            ["2.4"]  = 14,   // KT-KT (≤ 40 tỷ)
            ["2.5"]  = 15,   // KT-KT sửa chữa
            ["2.6"]  = 16,   // KT-KT nạo vét
            ["2.7"]  = 20,   // TKKT Dân dụng
            ["2.8"]  = 20,   // TKBVTC Dân dụng
            ["2.9"]  = 22,   // TKKT Công nghiệp
            ["2.10"] = 22,   // TKBVTC Công nghiệp
            ["2.11"] = 26,   // TKKT Giao thông
            ["2.12"] = 26,   // TKBVTC Giao thông
            ["2.13"] = 28,   // TKKT Nông nghiệp & Môi trường
            ["2.14"] = 28,   // TKBVTC Nông nghiệp & Môi trường
            ["2.15"] = 30,   // TKKT Hạ tầng kỹ thuật
            ["2.16"] = 30,   // TKBVTC Hạ tầng kỹ thuật
            ["2.17"] = 35,   // Thẩm tra BCNC tiền khả thi
            ["2.18"] = 36,   // Thẩm tra BCNC khả thi
            ["2.19"] = 39,   // Thẩm tra thiết kế
            ["2.20"] = 41,   // Thẩm tra dự toán
            ["2.21"] = 43,   // HSMT tư vấn
            ["2.22"] = 44,   // HSMT thi công
            ["2.23"] = 45,   // HSMT mua sắm TB
            ["2.24"] = 47,   // GS thi công
            ["2.25"] = 48,   // GS lắp đặt TB
            ["2.26"] = 48,   // GS khảo sát
            ["DD1"]  = 21,   // Thiết kế phần TB (TB ≥ 50%)
            ["HTKT1"] = 31,  // Thông tin KT - Trạm BTS, CS
            ["HTKT2"] = 32,  // Thông tin KT - Truyền dẫn
        };

        /// <summary>
        /// Lấy số trang PDF từ STT bảng.
        /// VD: LaySoTrang("2.7") → 20
        /// </summary>
        public static int LaySoTrang(string sttBang)
        {
            if (string.IsNullOrWhiteSpace(sttBang)) return 1;
            return _map.TryGetValue(sttBang.Trim(), out int trang) ? trang : 1;
        }
    }
}
