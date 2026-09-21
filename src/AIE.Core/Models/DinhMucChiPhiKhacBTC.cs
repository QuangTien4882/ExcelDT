using System;
using System.Collections.Generic;
using System.Linq;

namespace AIE.Core.Models
{
    /// <summary>
    /// Các bảng định mức tỷ lệ % chi phí do Bộ Tài chính ban hành:
    /// - Phí thẩm tra, phê duyệt quyết toán
    /// - Chi phí kiểm toán độc lập
    /// - Phí thẩm định dự án đầu tư xây dựng
    /// - Phí thẩm định thiết kế xây dựng triển khai sau thiết kế cơ sở
    /// - Phí thẩm định dự toán xây dựng công trình
    /// </summary>
    public static class DinhMucChiPhiKhacBTC
    {
        // =========================================================================
        // 1. ĐỊNH MỨC PHÍ THẨM TRA, PHÊ DUYỆT QUYẾT TOÁN
        // Mốc quy mô: <=5, 10, 50, 100, 500, 1000, >=10000 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoQuyetToan = new decimal[] { 5m, 10m, 50m, 100m, 500m, 1000m, 10000m };
        public static readonly decimal[] TiLeThamTraQuyetToan = new decimal[] { 0.570m, 0.390m, 0.285m, 0.225m, 0.135m, 0.090m, 0.048m };

        // =========================================================================
        // 2. ĐỊNH MỨC CHI PHÍ KIỂM TOÁN ĐỘC LẬP
        // Mốc quy mô: <=5, 10, 50, 100, 500, 1000, >=10000 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoKiemToan = new decimal[] { 5m, 10m, 50m, 100m, 500m, 1000m, 10000m };
        public static readonly decimal[] TiLeKiemToanDocLap = new decimal[] { 0.960m, 0.645m, 0.450m, 0.345m, 0.195m, 0.129m, 0.069m };

        // =========================================================================
        // 3. ĐỊNH MỨC PHÍ THẨM ĐỊNH DỰ ÁN ĐẦU TƯ XÂY DỰNG (Theo TMĐT)
        // Mốc quy mô: <=15, 25, 50, 100, 200, 500, 1000, 2000, 5000, >=10000 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoThamDinhDuAn = new decimal[] { 15m, 25m, 50m, 100m, 200m, 500m, 1000m, 2000m, 5000m, 10000m };
        public static readonly decimal[] TiLeThamDinhDuAn = new decimal[] { 0.0190m, 0.0170m, 0.0150m, 0.0130m, 0.0100m, 0.0080m, 0.0050m, 0.0030m, 0.0020m, 0.0010m };

        // =========================================================================
        // 4 & 5. MỐC QUY MÔ CHI PHÍ XÂY DỰNG CHO THẨM ĐỊNH THIẾT KẾ VÀ DỰ TOÁN
        // Mốc quy mô: <=15, 50, 100, 200, 500, 1000, 2000, 5000, >=8000 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoThamDinhTKDT = new decimal[] { 15m, 50m, 100m, 200m, 500m, 1000m, 2000m, 5000m, 8000m };

        // 4. PHÍ THẨM ĐỊNH THIẾT KẾ XÂY DỰNG
        public static readonly Dictionary<string, decimal[]> TiLeThamDinhThietKe = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 0.165m, 0.110m, 0.085m, 0.065m, 0.050m, 0.041m, 0.029m, 0.022m, 0.019m },
            ["Công nghiệp"] = new decimal[] { 0.190m, 0.126m, 0.097m, 0.075m, 0.058m, 0.044m, 0.035m, 0.026m, 0.022m },
            ["Giao thông"] = new decimal[] { 0.109m, 0.072m, 0.055m, 0.043m, 0.033m, 0.025m, 0.021m, 0.016m, 0.014m },
            [DinhMucTT38Database.LoaiNongNghiepMoiTruong] = new decimal[] { 0.121m, 0.080m, 0.061m, 0.048m, 0.037m, 0.028m, 0.023m, 0.017m, 0.014m },
            ["Nông nghiệp & PTNT"] = new decimal[] { 0.121m, 0.080m, 0.061m, 0.048m, 0.037m, 0.028m, 0.023m, 0.017m, 0.014m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 0.126m, 0.085m, 0.065m, 0.050m, 0.039m, 0.030m, 0.026m, 0.019m, 0.017m }
        };

        // 5. PHÍ THẨM ĐỊNH DỰ TOÁN XÂY DỰNG
        public static readonly Dictionary<string, decimal[]> TiLeThamDinhDuToan = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 0.160m, 0.106m, 0.083m, 0.062m, 0.046m, 0.038m, 0.028m, 0.021m, 0.018m },
            ["Công nghiệp"] = new decimal[] { 0.185m, 0.121m, 0.094m, 0.072m, 0.055m, 0.041m, 0.033m, 0.023m, 0.020m },
            ["Giao thông"] = new decimal[] { 0.106m, 0.068m, 0.054m, 0.041m, 0.031m, 0.024m, 0.020m, 0.014m, 0.012m },
            [DinhMucTT38Database.LoaiNongNghiepMoiTruong] = new decimal[] { 0.117m, 0.076m, 0.060m, 0.046m, 0.035m, 0.026m, 0.022m, 0.016m, 0.014m },
            ["Nông nghiệp & PTNT"] = new decimal[] { 0.117m, 0.076m, 0.060m, 0.046m, 0.035m, 0.026m, 0.022m, 0.016m, 0.014m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 0.122m, 0.082m, 0.062m, 0.047m, 0.037m, 0.029m, 0.024m, 0.017m, 0.014m }
        };

        /// <summary>
        /// Hàm nội suy 2 chiều giữa các mốc quy mô theo công thức quy chuẩn:
        /// Nt = Nb - ((Nb - Na) / (Ga - Gb)) * (Gt - Gb)
        /// Trả về giá trị nội suy chính xác (KHÔNG làm tròn bên trong) — nơi cần làm tròn
        /// sẽ tự làm tròn theo ngữ cảnh (ví dụ DinhMucTT38Engine).
        /// </summary>
        public static decimal NoiSuy(decimal[] mocs, decimal[] tiLes, decimal quyMo)
        {
            if (mocs == null || tiLes == null || mocs.Length == 0 || tiLes.Length == 0)
                return 0;

            if (quyMo <= mocs[0]) return tiLes[0];
            if (quyMo >= mocs[mocs.Length - 1]) return tiLes[tiLes.Length - 1];

            for (int i = 0; i < mocs.Length - 1; i++)
            {
                decimal gb = mocs[i];
                decimal ga = mocs[i + 1];
                if (quyMo >= gb && quyMo <= ga)
                {
                    decimal nb = tiLes[i];
                    decimal na = tiLes[i + 1];
                    if (ga == gb) return nb;

                    decimal nt = nb - ((nb - na) / (ga - gb)) * (quyMo - gb);
                    return nt;
                }
            }

            return tiLes[tiLes.Length - 1];
        }
    }
}
