using System;
using System.Collections.Generic;
using System.Linq;

namespace AIE.Core.Models
{
    /// <summary>
    /// Cơ sở dữ liệu định mức chi phí Quản lý dự án và Tư vấn đầu tư xây dựng
    /// theo Thông tư số 38/2026/TT-BXD (Phụ lục VIII) - Số liệu khớp 100% văn bản PDF.
    /// </summary>
    public static class DinhMucTT38Database
    {
        /// <summary>Key chuẩn của loại công trình số 4 (nông nghiệp &amp; môi trường).</summary>
        public const string LoaiNongNghiepMoiTruong = "Nông nghiệp & Môi trường";

        // =========================================================================
        // BẢNG 1.1: ĐỊNH MỨC CHI PHÍ QUẢN LÝ DỰ ÁN (Phần trăm %)
        // Căn cứ tính: Chi phí xây dựng và thiết bị trước thuế (G_XD + G_TB)
        // Mốc quy mô: <=10, 20, 50, 100, 200, 500, 1000, 2000, 5000, 10000, 20000, >=30000 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoQLDA = new decimal[]
        {
            10m, 20m, 50m, 100m, 200m, 500m, 1000m, 2000m, 5000m, 10000m, 20000m, 30000m
        };

        public static readonly Dictionary<string, decimal[]> TiLeQLDA = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 3.446m, 2.923m, 2.610m, 2.017m, 1.886m, 1.514m, 1.239m, 0.958m, 0.711m, 0.510m, 0.381m, 0.305m },
            ["Công nghiệp"] = new decimal[] { 3.557m, 3.018m, 2.694m, 2.082m, 1.947m, 1.564m, 1.279m, 1.103m, 0.734m, 0.527m, 0.393m, 0.314m },
            ["Giao thông"] = new decimal[] { 3.024m, 2.566m, 2.292m, 1.771m, 1.655m, 1.329m, 1.088m, 0.937m, 0.624m, 0.448m, 0.335m, 0.268m },
            [LoaiNongNghiepMoiTruong] = new decimal[] { 3.263m, 2.769m, 2.473m, 1.910m, 1.786m, 1.434m, 1.174m, 1.012m, 0.674m, 0.484m, 0.361m, 0.289m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 2.901m, 2.461m, 2.198m, 1.593m, 1.560m, 1.275m, 1.071m, 0.899m, 0.599m, 0.429m, 0.321m, 0.257m }
        };

        // =========================================================================
        // BẢNG 2.1: ĐỊNH MỨC LẬP BÁO CÁO ĐỀ XUẤT CHỦ TRƯƠNG ĐẦU TƯ
        // Căn cứ tính: G_XD + G_TB
        // Mốc quy mô: <=15, 50, 100, 500, 800, 1000, 1500, 2300, 3000, 4000, >=4600 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoBCDeXuat = new decimal[]
        {
            15m, 50m, 100m, 500m, 800m, 1000m, 1500m, 2300m, 3000m, 4000m, 4600m
        };

        public static readonly Dictionary<string, decimal[]> TiLeBCDeXuat = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 0.301m, 0.169m, 0.108m, 0.045m, 0.041m, 0.039m, 0.036m, 0.032m, 0.029m, 0.024m, 0.021m },
            ["Công nghiệp"] = new decimal[] { 0.341m, 0.198m, 0.132m, 0.073m, 0.067m, 0.063m, 0.056m, 0.048m, 0.041m, 0.031m, 0.027m },
            ["Giao thông"] = new decimal[] { 0.165m, 0.100m, 0.071m, 0.028m, 0.026m, 0.025m, 0.023m, 0.021m, 0.019m, 0.017m, 0.015m },
            [LoaiNongNghiepMoiTruong] = new decimal[] { 0.226m, 0.137m, 0.086m, 0.038m, 0.035m, 0.033m, 0.030m, 0.027m, 0.024m, 0.021m, 0.018m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 0.172m, 0.105m, 0.073m, 0.030m, 0.027m, 0.026m, 0.024m, 0.021m, 0.018m, 0.015m, 0.014m }
        };

        // =========================================================================
        // BẢNG 2.2: ĐỊNH MỨC LẬP BÁO CÁO NGHIÊN CỨU TIỀN KHẢ THI
        // Căn cứ tính: G_XD + G_TB
        // Mốc quy mô: <=15, 20, 50, 100, 200, 500, 1000, 2000, 5000, 10000, 20000, >=30000 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoTienKhaThi = new decimal[]
        {
            15m, 20m, 50m, 100m, 200m, 500m, 1000m, 2000m, 5000m, 10000m, 20000m, 30000m
        };

        public static readonly Dictionary<string, decimal[]> TiLeTienKhaThi = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 0.668m, 0.503m, 0.376m, 0.240m, 0.161m, 0.100m, 0.086m, 0.073m, 0.050m, 0.040m, 0.026m, 0.022m },
            ["Công nghiệp"] = new decimal[] { 0.757m, 0.612m, 0.441m, 0.294m, 0.206m, 0.163m, 0.141m, 0.110m, 0.074m, 0.057m, 0.034m, 0.027m },
            ["Giao thông"] = new decimal[] { 0.413m, 0.345m, 0.251m, 0.177m, 0.108m, 0.071m, 0.062m, 0.053m, 0.036m, 0.029m, 0.019m, 0.016m },
            [LoaiNongNghiepMoiTruong] = new decimal[] { 0.566m, 0.472m, 0.343m, 0.216m, 0.144m, 0.096m, 0.082m, 0.070m, 0.048m, 0.039m, 0.025m, 0.021m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 0.431m, 0.360m, 0.262m, 0.183m, 0.112m, 0.074m, 0.065m, 0.055m, 0.038m, 0.030m, 0.020m, 0.017m }
        };

        // =========================================================================
        // BẢNG 2.3: ĐỊNH MỨC LẬP BÁO CÁO NGHIÊN CỨU KHẢ THI (FS)
        // Căn cứ tính: G_XD + G_TB
        // Mốc quy mô: <=15, 20, 50, 100, 200, 500, 1000, 2000, 5000, 10000, 20000, >=30000 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoFS = MocQuyMoTienKhaThi;

        public static readonly Dictionary<string, decimal[]> TiLeLapFS = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 1.114m, 0.914m, 0.751m, 0.534m, 0.402m, 0.287m, 0.246m, 0.209m, 0.167m, 0.134m, 0.102m, 0.086m },
            ["Công nghiệp"] = new decimal[] { 1.261m, 1.112m, 0.882m, 0.654m, 0.515m, 0.466m, 0.404m, 0.315m, 0.248m, 0.189m, 0.135m, 0.107m },
            ["Giao thông"] = new decimal[] { 0.689m, 0.628m, 0.501m, 0.393m, 0.271m, 0.203m, 0.177m, 0.151m, 0.120m, 0.097m, 0.075m, 0.063m },
            [LoaiNongNghiepMoiTruong] = new decimal[] { 0.943m, 0.858m, 0.685m, 0.480m, 0.361m, 0.273m, 0.234m, 0.201m, 0.161m, 0.129m, 0.100m, 0.084m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 0.719m, 0.654m, 0.524m, 0.407m, 0.280m, 0.211m, 0.185m, 0.158m, 0.127m, 0.101m, 0.078m, 0.065m }
        };

        // =========================================================================
        // BẢNG 2.4: ĐỊNH MỨC LẬP BÁO CÁO KINH TẾ - KỸ THUẬT
        // Áp dụng cho dự án có tổng mức đầu tư <= 40 tỷ (không gồm GPMB, tiền SDĐ)
        // Căn cứ tính: G_XD + G_TB
        // Mốc quy mô: <=1, 3, 7, 15, 20, >=40 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoBaoCaoKTKT = new decimal[] { 1m, 3m, 7m, 15m, 20m, 40m };

        public static readonly Dictionary<string, decimal[]> TiLeBaoCaoKTKT = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 6.5m, 4.7m, 4.2m, 3.6m, 2.9m, 2.4m },
            ["Công nghiệp"] = new decimal[] { 6.7m, 4.8m, 4.3m, 3.8m, 3.2m, 2.5m },
            ["Giao thông"] = new decimal[] { 5.4m, 3.6m, 3.0m, 2.9m, 2.7m, 2.1m },
            [LoaiNongNghiepMoiTruong] = new decimal[] { 6.2m, 4.4m, 3.9m, 3.6m, 3.1m, 2.4m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 5.8m, 4.2m, 3.4m, 3.0m, 2.8m, 2.2m }
        };

        // =========================================================================
        // BẢNG 2.5: ĐỊNH MỨC LẬP BÁO CÁO KT-KT (Dự án sửa chữa bảo trì nhóm C)
        // Căn cứ tính: G_XD
        // Mốc quy mô: <=1, 3, 5, 10, 20, 50, 100, 150, 200, >=240 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoKTKTSuaChua = new decimal[]
        {
            1m, 3m, 5m, 10m, 20m, 50m, 100m, 150m, 200m, 240m
        };

        public static readonly Dictionary<string, decimal[]> TiLeKTKTSuaChua = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 5.31m, 3.96m, 3.87m, 3.41m, 2.70m, 1.98m, 1.80m, 1.73m, 1.65m, 1.62m },
            ["Công nghiệp"] = new decimal[] { 5.49m, 4.07m, 3.96m, 3.48m, 3.00m, 2.03m, 1.85m, 1.77m, 1.70m, 1.67m },
            ["Giao thông"] = new decimal[] { 4.23m, 2.88m, 2.85m, 2.58m, 2.49m, 1.64m, 1.49m, 1.43m, 1.31m, 1.29m },
            [LoaiNongNghiepMoiTruong] = new decimal[] { 4.95m, 3.60m, 3.53m, 3.12m, 2.75m, 1.85m, 1.65m, 1.56m, 1.49m, 1.44m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 4.37m, 3.17m, 3.09m, 2.72m, 2.55m, 1.73m, 1.56m, 1.47m, 1.40m, 1.37m }
        };

        // =========================================================================
        // BẢNG 2.6: ĐỊNH MỨC LẬP BÁO CÁO KT-KT (Nạo vét duy tu luồng hàng hải, đường thủy nội địa)
        // Căn cứ tính: G_XD - Chỉ áp dụng loại Giao thông
        // Mốc quy mô: <=1, 3, 5, 10, 20, 50, 100, 150, 200, >=240 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoKTKTNaoVet = MocQuyMoKTKTSuaChua;

        public static readonly Dictionary<string, decimal[]> TiLeKTKTNaoVet = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Giao thông"] = new decimal[] { 2.82m, 1.92m, 1.90m, 1.72m, 1.66m, 1.09m, 0.99m, 0.95m, 0.87m, 0.86m }
        };

        // =========================================================================
        // BẢNG 2.7 - 2.16: ĐỊNH MỨC THIẾT KẾ XÂY DỰNG CÔNG TRÌNH
        // Căn cứ tính: Chi phí xây dựng trước thuế (G_XD)
        // Mốc quy mô: <=10, 20, 50, 100, 200, 500, 1000, 2000, 5000, 8000, >=10000 tỷ
        // Mỗi loại công trình có 02 bộ bảng:
        //  - TKKT: thiết kế kỹ thuật (dự án 3 bước - thực hiện 2 bước sau thiết kế cơ sở)
        //  - TKBVTC: thiết kế bản vẽ thi công (dự án 2 bước)
        // Bảng theo cấp: Cấp đặc biệt, Cấp I, Cấp II, Cấp III, Cấp IV.
        // Trường hợp "–" (Cấp IV quy mô lớn): dùng tỷ lệ của mốc cuối có số liệu trước đó.
        // =========================================================================
        public static readonly decimal[] MocQuyMoThietKe = new decimal[]
        {
            10m, 20m, 50m, 100m, 200m, 500m, 1000m, 2000m, 5000m, 8000m, 10000m
        };

        public const string CapDacBiet = "Cấp đặc biệt";
        public const string CapI = "Cấp I";
        public const string CapII = "Cấp II";
        public const string CapIII = "Cấp III";
        public const string CapIV = "Cấp IV";

        /// <summary>Thứ tự các cấp công trình dùng trong bảng 2.7 - 2.16.</summary>
        public static readonly string[] CapCongTrinhOrder = new[]
        {
            CapDacBiet, CapI, CapII, CapIII, CapIV
        };

        // Bảng 2.7 (TKKT) & 2.8 (TKBVTC) - Dân dụng
        public static readonly Dictionary<string, Dictionary<string, decimal[]>> TiLeThietKeKyThuat = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new(StringComparer.OrdinalIgnoreCase)
            {
                [CapDacBiet] = new decimal[] { 3.22m, 2.81m, 2.36m, 2.15m, 1.96m, 1.65m, 1.36m, 1.16m, 0.89m, 0.68m, 0.61m },
                [CapI] = new decimal[] { 2.93m, 2.55m, 2.14m, 1.94m, 1.78m, 1.50m, 1.22m, 1.05m, 0.80m, 0.61m, 0.55m },
                [CapII] = new decimal[] { 2.67m, 2.33m, 1.96m, 1.77m, 1.62m, 1.37m, 1.11m, 0.94m, 0.73m, 0.55m, 0.50m },
                [CapIII] = new decimal[] { 2.36m, 2.07m, 1.74m, 1.57m, 1.43m, 1.21m, 0.98m, 0.83m, 0.64m, 0.48m, 0.44m },
                [CapIV] = new decimal[] { 2.07m, 1.81m, 1.48m, 1.30m, 1.06m, 0.89m, 0.89m, 0.89m, 0.89m, 0.89m, 0.89m }
            },
            ["Công nghiệp"] = new(StringComparer.OrdinalIgnoreCase)
            {
                [CapDacBiet] = new decimal[] { 2.96m, 2.73m, 2.34m, 2.13m, 1.92m, 1.76m, 1.54m, 1.30m, 0.97m, 0.79m, 0.70m },
                [CapI] = new decimal[] { 2.47m, 2.27m, 1.93m, 1.77m, 1.60m, 1.46m, 1.28m, 1.09m, 0.80m, 0.65m, 0.58m },
                [CapII] = new decimal[] { 2.03m, 1.86m, 1.59m, 1.46m, 1.32m, 1.20m, 1.05m, 0.90m, 0.66m, 0.53m, 0.48m },
                [CapIII] = new decimal[] { 1.78m, 1.65m, 1.40m, 1.27m, 1.17m, 1.06m, 0.93m, 0.79m, 0.58m, 0.47m, 0.42m },
                [CapIV] = new decimal[] { 1.59m, 1.47m, 1.24m, 1.14m, 0.98m, 0.83m, 0.83m, 0.83m, 0.83m, 0.83m, 0.83m }
            },
            ["Giao thông"] = new(StringComparer.OrdinalIgnoreCase)
            {
                [CapDacBiet] = new decimal[] { 2.05m, 1.92m, 1.68m, 1.50m, 1.36m, 1.24m, 1.08m, 0.92m, 0.68m, 0.51m, 0.45m },
                [CapI] = new decimal[] { 1.44m, 1.39m, 1.13m, 1.05m, 0.95m, 0.81m, 0.68m, 0.58m, 0.44m, 0.34m, 0.28m },
                [CapII] = new decimal[] { 1.19m, 1.08m, 0.92m, 0.84m, 0.77m, 0.70m, 0.60m, 0.51m, 0.39m, 0.29m, 0.25m },
                [CapIII] = new decimal[] { 1.05m, 0.93m, 0.81m, 0.74m, 0.68m, 0.58m, 0.48m, 0.43m, 0.32m, 0.25m, 0.21m },
                [CapIV] = new decimal[] { 0.95m, 0.87m, 0.76m, 0.69m, 0.59m, 0.49m, 0.43m, 0.43m, 0.43m, 0.43m, 0.43m }
            },
            [LoaiNongNghiepMoiTruong] = new(StringComparer.OrdinalIgnoreCase)
            {
                [CapDacBiet] = new decimal[] { 2.98m, 2.60m, 2.20m, 1.98m, 1.83m, 1.54m, 1.30m, 1.13m, 0.85m, 0.66m, 0.58m },
                [CapI] = new decimal[] { 2.70m, 2.36m, 1.99m, 1.78m, 1.66m, 1.39m, 1.17m, 1.02m, 0.77m, 0.59m, 0.52m },
                [CapII] = new decimal[] { 2.48m, 2.14m, 1.80m, 1.61m, 1.51m, 1.22m, 1.05m, 0.87m, 0.67m, 0.49m, 0.42m },
                [CapIII] = new decimal[] { 2.20m, 1.90m, 1.60m, 1.43m, 1.24m, 1.06m, 0.90m, 0.77m, 0.59m, 0.43m, 0.37m },
                [CapIV] = new decimal[] { 1.74m, 1.52m, 1.27m, 1.12m, 1.01m, 0.80m, 0.64m, 0.64m, 0.64m, 0.64m, 0.64m }
            },
            ["Hạ tầng kỹ thuật"] = new(StringComparer.OrdinalIgnoreCase)
            {
                [CapDacBiet] = new decimal[] { 2.22m, 1.94m, 1.63m, 1.48m, 1.36m, 1.14m, 0.97m, 0.83m, 0.61m, 0.48m, 0.43m },
                [CapI] = new decimal[] { 2.09m, 1.83m, 1.53m, 1.38m, 1.28m, 1.04m, 0.90m, 0.75m, 0.53m, 0.39m, 0.33m },
                [CapII] = new decimal[] { 1.86m, 1.62m, 1.36m, 1.22m, 1.13m, 0.91m, 0.78m, 0.66m, 0.47m, 0.34m, 0.29m },
                [CapIII] = new decimal[] { 1.62m, 1.39m, 1.19m, 1.07m, 0.97m, 0.80m, 0.70m, 0.56m, 0.41m, 0.29m, 0.25m },
                [CapIV] = new decimal[] { 1.45m, 1.23m, 1.01m, 0.92m, 0.80m, 0.70m, 0.58m, 0.58m, 0.58m, 0.58m, 0.58m }
            }
        };

        public static readonly Dictionary<string, Dictionary<string, decimal[]>> TiLeThietKeBVTC = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new(StringComparer.OrdinalIgnoreCase)
            {
                [CapDacBiet] = new decimal[] { 4.66m, 4.05m, 3.41m, 3.10m, 2.83m, 2.39m, 1.93m, 1.65m, 1.28m, 0.99m, 0.91m },
                [CapI] = new decimal[] { 4.22m, 3.66m, 3.10m, 2.82m, 2.57m, 2.17m, 1.76m, 1.51m, 1.16m, 0.90m, 0.80m },
                [CapII] = new decimal[] { 3.85m, 3.33m, 2.80m, 2.54m, 2.34m, 1.98m, 1.61m, 1.36m, 1.06m, 0.82m, 0.72m },
                [CapIII] = new decimal[] { 3.41m, 2.95m, 2.48m, 2.25m, 2.07m, 1.75m, 1.43m, 1.20m, 0.94m, 0.72m, 0.63m },
                [CapIV] = new decimal[] { 2.92m, 2.55m, 2.12m, 1.86m, 1.51m, 1.30m, 1.30m, 1.30m, 1.30m, 1.30m, 1.30m }
            },
            ["Công nghiệp"] = new(StringComparer.OrdinalIgnoreCase)
            {
                [CapDacBiet] = new decimal[] { 4.70m, 4.27m, 3.66m, 3.32m, 3.01m, 2.75m, 2.40m, 2.03m, 1.52m, 1.21m, 1.04m },
                [CapI] = new decimal[] { 3.87m, 3.57m, 3.02m, 2.77m, 2.50m, 2.28m, 2.01m, 1.70m, 1.26m, 1.02m, 0.88m },
                [CapII] = new decimal[] { 3.13m, 2.90m, 2.43m, 2.24m, 2.03m, 1.90m, 1.66m, 1.42m, 1.04m, 0.82m, 0.72m },
                [CapIII] = new decimal[] { 2.78m, 2.57m, 2.16m, 1.99m, 1.79m, 1.68m, 1.47m, 1.25m, 0.91m, 0.72m, 0.64m },
                [CapIV] = new decimal[] { 2.46m, 2.25m, 1.89m, 1.72m, 1.47m, 1.22m, 1.22m, 1.22m, 1.22m, 1.22m, 1.22m }
            },
            ["Giao thông"] = new(StringComparer.OrdinalIgnoreCase)
            {
                [CapDacBiet] = new decimal[] { 3.01m, 2.76m, 2.36m, 2.15m, 1.95m, 1.78m, 1.52m, 1.32m, 1.02m, 0.75m, 0.66m },
                [CapI] = new decimal[] { 2.27m, 2.15m, 1.83m, 1.67m, 1.51m, 1.38m, 1.21m, 1.03m, 0.79m, 0.61m, 0.49m },
                [CapII] = new decimal[] { 1.67m, 1.55m, 1.32m, 1.20m, 1.10m, 1.01m, 0.85m, 0.72m, 0.56m, 0.42m, 0.36m },
                [CapIII] = new decimal[] { 1.48m, 1.37m, 1.17m, 1.06m, 0.97m, 0.82m, 0.70m, 0.59m, 0.45m, 0.33m, 0.29m },
                [CapIV] = new decimal[] { 1.37m, 1.26m, 1.08m, 0.98m, 0.83m, 0.71m, 0.71m, 0.71m, 0.71m, 0.71m, 0.71m }
            },
            [LoaiNongNghiepMoiTruong] = new(StringComparer.OrdinalIgnoreCase)
            {
                [CapDacBiet] = new decimal[] { 4.29m, 3.75m, 3.17m, 2.85m, 2.60m, 2.21m, 1.87m, 1.58m, 1.22m, 0.95m, 0.83m },
                [CapI] = new decimal[] { 3.89m, 3.40m, 2.87m, 2.57m, 2.36m, 2.00m, 1.69m, 1.43m, 1.10m, 0.85m, 0.74m },
                [CapII] = new decimal[] { 3.53m, 3.11m, 2.62m, 2.34m, 2.15m, 1.73m, 1.48m, 1.25m, 0.96m, 0.69m, 0.58m },
                [CapIII] = new decimal[] { 3.13m, 2.76m, 2.31m, 2.07m, 1.79m, 1.52m, 1.29m, 1.10m, 0.83m, 0.60m, 0.51m },
                [CapIV] = new decimal[] { 2.48m, 2.19m, 1.82m, 1.61m, 1.41m, 1.14m, 1.14m, 1.14m, 1.14m, 1.14m, 1.14m }
            },
            ["Hạ tầng kỹ thuật"] = new(StringComparer.OrdinalIgnoreCase)
            {
                [CapDacBiet] = new decimal[] { 3.23m, 2.79m, 2.35m, 2.13m, 1.95m, 1.64m, 1.39m, 1.19m, 0.90m, 0.70m, 0.63m },
                [CapI] = new decimal[] { 3.01m, 2.63m, 2.21m, 1.99m, 1.82m, 1.49m, 1.28m, 1.07m, 0.79m, 0.58m, 0.49m },
                [CapII] = new decimal[] { 2.68m, 2.33m, 1.97m, 1.77m, 1.58m, 1.32m, 1.14m, 0.92m, 0.70m, 0.51m, 0.43m },
                [CapIII] = new decimal[] { 2.36m, 2.01m, 1.72m, 1.55m, 1.39m, 1.16m, 1.02m, 0.81m, 0.61m, 0.44m, 0.36m },
                [CapIV] = new decimal[] { 2.07m, 1.76m, 1.49m, 1.35m, 1.15m, 0.98m, 0.98m, 0.98m, 0.98m, 0.98m, 0.98m }
            }
        };

        // =========================================================================
        // BẢNG 2.17: ĐỊNH MỨC THẨM TRA BÁO CÁO NGHIÊN CỨU TIỀN KHẢ THI
        // BẢNG 2.18: ĐỊNH MỨC THẨM TRA BÁO CÁO NGHIÊN CỨU KHẢ THI
        // Căn cứ tính: G_XD + G_TB
        // Mốc quy mô: <=15, 20, 50, 100, 200, 500, 1000, 2000, 5000, 10000, 20000, >=30000 tỷ
        // =========================================================================
        public static readonly Dictionary<string, decimal[]> TiLeLapTT_BCNCTienKhaThi = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 0.071m, 0.059m, 0.048m, 0.034m, 0.025m, 0.016m, 0.014m, 0.012m, 0.009m, 0.007m, 0.005m, 0.004m },
            ["Công nghiệp"] = new decimal[] { 0.098m, 0.083m, 0.067m, 0.049m, 0.037m, 0.028m, 0.025m, 0.020m, 0.015m, 0.010m, 0.007m, 0.005m },
            ["Giao thông"] = new decimal[] { 0.054m, 0.049m, 0.039m, 0.030m, 0.020m, 0.013m, 0.011m, 0.009m, 0.007m, 0.005m, 0.004m, 0.003m },
            [LoaiNongNghiepMoiTruong] = new decimal[] { 0.064m, 0.058m, 0.047m, 0.033m, 0.024m, 0.015m, 0.013m, 0.011m, 0.009m, 0.006m, 0.005m, 0.004m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 0.056m, 0.051m, 0.041m, 0.032m, 0.021m, 0.013m, 0.012m, 0.010m, 0.008m, 0.005m, 0.004m, 0.003m }
        };

        public static readonly Dictionary<string, decimal[]> TiLeLapTT_BCNCKhaThi = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 0.204m, 0.168m, 0.138m, 0.097m, 0.070m, 0.046m, 0.041m, 0.034m, 0.026m, 0.019m, 0.015m, 0.012m },
            ["Công nghiệp"] = new decimal[] { 0.281m, 0.238m, 0.190m, 0.141m, 0.107m, 0.080m, 0.070m, 0.056m, 0.044m, 0.029m, 0.020m, 0.015m },
            ["Giao thông"] = new decimal[] { 0.153m, 0.139m, 0.112m, 0.087m, 0.058m, 0.036m, 0.032m, 0.026m, 0.020m, 0.014m, 0.010m, 0.009m },
            [LoaiNongNghiepMoiTruong] = new decimal[] { 0.182m, 0.167m, 0.133m, 0.094m, 0.068m, 0.044m, 0.037m, 0.032m, 0.026m, 0.017m, 0.014m, 0.010m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 0.160m, 0.145m, 0.116m, 0.092m, 0.060m, 0.037m, 0.034m, 0.029m, 0.022m, 0.015m, 0.010m, 0.009m }
        };

        // =========================================================================
        // BẢNG 2.19: ĐỊNH MỨC THẨM TRA THIẾT KẾ XÂY DỰNG
        // BẢNG 2.20: ĐỊNH MỨC THẨM TRA DỰ TOÁN XÂY DỰNG
        // Căn cứ tính: G_XD
        // Mốc quy mô: <=10, 20, 50, 100, 200, 500, 1000, 2000, 5000, 8000, >=10000 tỷ
        // Chi phí tối thiểu: 2.000.000 đồng
        // =========================================================================
        public static readonly decimal[] MocQuyMoThamTraTK = MocQuyMoThietKe;

        public static readonly Dictionary<string, decimal[]> TiLeThamTraTK = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 0.258m, 0.223m, 0.172m, 0.143m, 0.108m, 0.083m, 0.068m, 0.044m, 0.033m, 0.028m, 0.026m },
            ["Công nghiệp"] = new decimal[] { 0.290m, 0.252m, 0.192m, 0.146m, 0.113m, 0.087m, 0.066m, 0.053m, 0.038m, 0.031m, 0.028m },
            ["Giao thông"] = new decimal[] { 0.170m, 0.147m, 0.113m, 0.084m, 0.073m, 0.055m, 0.042m, 0.035m, 0.024m, 0.020m, 0.017m },
            [LoaiNongNghiepMoiTruong] = new decimal[] { 0.189m, 0.163m, 0.125m, 0.093m, 0.073m, 0.056m, 0.043m, 0.035m, 0.026m, 0.022m, 0.019m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 0.197m, 0.172m, 0.133m, 0.099m, 0.076m, 0.059m, 0.046m, 0.040m, 0.029m, 0.024m, 0.021m }
        };

        public static readonly Dictionary<string, decimal[]> TiLeThamTraDuToan = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 0.250m, 0.219m, 0.166m, 0.140m, 0.105m, 0.077m, 0.064m, 0.043m, 0.032m, 0.027m, 0.025m },
            ["Công nghiệp"] = new decimal[] { 0.282m, 0.244m, 0.185m, 0.141m, 0.108m, 0.083m, 0.062m, 0.050m, 0.034m, 0.030m, 0.027m },
            ["Giao thông"] = new decimal[] { 0.166m, 0.142m, 0.106m, 0.082m, 0.069m, 0.052m, 0.041m, 0.034m, 0.021m, 0.018m, 0.016m },
            [LoaiNongNghiepMoiTruong] = new decimal[] { 0.183m, 0.158m, 0.119m, 0.092m, 0.070m, 0.053m, 0.040m, 0.034m, 0.024m, 0.021m, 0.018m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 0.191m, 0.166m, 0.128m, 0.095m, 0.072m, 0.056m, 0.044m, 0.037m, 0.026m, 0.022m, 0.020m }
        };

        // =========================================================================
        // BẢNG 2.21: HSMT, ĐÁNH GIÁ HSDT GÓI THẦU TƯ VẤN
        // Căn cứ tính: Chi phí tư vấn trong dự toán gói thầu tư vấn
        // Mốc quy mô: <=1, 3, 5, 10, 20, 50, >=100 tỷ - Một tỷ lệ chung cho mọi loại CT
        // =========================================================================
        public static readonly decimal[] MocQuyMoHSMTTuVan = new decimal[] { 1m, 3m, 5m, 10m, 20m, 50m, 100m };
        public static readonly decimal[] TiLeHSMTTuVan = new decimal[] { 0.816m, 0.583m, 0.505m, 0.389m, 0.311m, 0.176m, 0.114m };

        // =========================================================================
        // BẢNG 2.22: HSMT, ĐÁNH GIÁ HSDT THI CÔNG XÂY DỰNG
        // Căn cứ tính: G_XD trong dự toán gói thầu thi công
        // Mốc quy mô: <=10, 20, 50, 100, 200, 500, 1000, >=2000 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoHSMT = new decimal[] { 10m, 20m, 50m, 100m, 200m, 500m, 1000m, 2000m };

        public static readonly Dictionary<string, decimal[]> TiLeHSMTThiCong = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 0.432m, 0.346m, 0.195m, 0.127m, 0.078m, 0.057m, 0.040m, 0.032m },
            ["Công nghiệp"] = new decimal[] { 0.549m, 0.379m, 0.211m, 0.144m, 0.096m, 0.067m, 0.052m, 0.041m },
            ["Giao thông"] = new decimal[] { 0.346m, 0.237m, 0.151m, 0.090m, 0.057m, 0.043m, 0.029m, 0.023m },
            [LoaiNongNghiepMoiTruong] = new decimal[] { 0.361m, 0.302m, 0.166m, 0.094m, 0.066m, 0.046m, 0.031m, 0.026m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 0.388m, 0.325m, 0.172m, 0.106m, 0.069m, 0.052m, 0.038m, 0.028m }
        };

        // =========================================================================
        // BẢNG 2.23: HSMT, ĐÁNH GIÁ HSDT MUA SẮM VẬT TƯ, THIẾT BỊ
        // Căn cứ tính: Chi phí vật tư, thiết bị (chưa GTGT) trong dự toán gói thầu mua sắm
        // Mốc quy mô: <=10, 20, 50, 100, 200, 500, 1000, >=2000 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoHSMTCuaTB = MocQuyMoHSMT;

        public static readonly Dictionary<string, decimal[]> TiLeHSMTCuaTB = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 0.367m, 0.346m, 0.181m, 0.113m, 0.102m, 0.081m, 0.055m, 0.043m },
            ["Công nghiệp"] = new decimal[] { 0.549m, 0.494m, 0.280m, 0.177m, 0.152m, 0.123m, 0.084m, 0.066m },
            ["Giao thông"] = new decimal[] { 0.261m, 0.230m, 0.131m, 0.084m, 0.074m, 0.056m, 0.040m, 0.032m },
            [LoaiNongNghiepMoiTruong] = new decimal[] { 0.281m, 0.245m, 0.140m, 0.090m, 0.078m, 0.061m, 0.050m, 0.037m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 0.302m, 0.260m, 0.156m, 0.102m, 0.087m, 0.069m, 0.054m, 0.041m }
        };

        // =========================================================================
        // BẢNG 2.24: ĐỊNH MỨC GIÁM SÁT THI CÔNG XÂY DỰNG
        // Căn cứ tính: G_XD trong dự toán gói thầu thi công
        // Mốc quy mô: <=10, 20, 50, 100, 200, 500, 1000, 2000, 5000, 8000, >=10000 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoGiamSat = MocQuyMoThietKe;

        public static readonly Dictionary<string, decimal[]> TiLeGiamSatThiCong = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 3.285m, 2.853m, 2.435m, 1.845m, 1.546m, 1.188m, 0.797m, 0.694m, 0.620m, 0.530m, 0.478m },
            ["Công nghiệp"] = new decimal[] { 3.508m, 3.137m, 2.559m, 2.074m, 1.604m, 1.301m, 0.823m, 0.716m, 0.640m, 0.550m, 0.493m },
            ["Giao thông"] = new decimal[] { 3.203m, 2.700m, 2.356m, 1.714m, 1.272m, 1.003m, 0.731m, 0.636m, 0.550m, 0.480m, 0.438m },
            [LoaiNongNghiepMoiTruong] = new decimal[] { 2.598m, 2.292m, 2.075m, 1.545m, 1.189m, 0.950m, 0.631m, 0.550m, 0.490m, 0.420m, 0.378m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 2.566m, 2.256m, 1.984m, 1.461m, 1.142m, 0.912m, 0.584m, 0.509m, 0.452m, 0.390m, 0.350m }
        };

        // =========================================================================
        // BẢNG 2.25: ĐỊNH MỨC GIÁM SÁT LẮP ĐẶT THIẾT BỊ
        // Căn cứ tính: Chi phí thiết bị (chưa GTGT) trong dự toán gói thầu thiết bị
        // Mốc quy mô: <=10, 20, 50, 100, 200, 500, 1000, 2000, 5000, 8000, >=10000 tỷ
        // =========================================================================
        public static readonly Dictionary<string, decimal[]> TiLeGiamSatLapDatTB = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 0.844m, 0.715m, 0.596m, 0.394m, 0.305m, 0.261m, 0.176m, 0.153m, 0.132m, 0.112m, 0.110m },
            ["Công nghiệp"] = new decimal[] { 1.147m, 1.005m, 0.958m, 0.811m, 0.490m, 0.422m, 0.356m, 0.309m, 0.270m, 0.230m, 0.210m },
            ["Giao thông"] = new decimal[] { 0.677m, 0.580m, 0.486m, 0.320m, 0.261m, 0.217m, 0.146m, 0.127m, 0.110m, 0.092m, 0.085m },
            [LoaiNongNghiepMoiTruong] = new decimal[] { 0.718m, 0.585m, 0.520m, 0.344m, 0.276m, 0.232m, 0.159m, 0.138m, 0.120m, 0.098m, 0.091m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 0.803m, 0.690m, 0.575m, 0.383m, 0.300m, 0.261m, 0.173m, 0.150m, 0.126m, 0.105m, 0.095m }
        };

        // =========================================================================
        // BẢNG 2.26: ĐỊNH MỨC GIÁM SÁT CÔNG TÁC KHẢO SÁT XÂY DỰNG
        // Căn cứ tính: Chi phí khảo sát (chưa GTGT) trong dự toán gói thầu khảo sát
        // Mốc quy mô: <=1, 5, 10, 20, >=50 tỷ - Một tỷ lệ chung cho mọi loại CT
        // =========================================================================
        public static readonly decimal[] MocQuyMoGiamSatKhaoSat = new decimal[] { 1m, 5m, 10m, 20m, 50m };
        public static readonly decimal[] TiLeGiamSatKhaoSat = new decimal[] { 4.072m, 3.541m, 3.079m, 2.707m, 2.381m };

        // =========================================================================
        // BẢNG DD1: ĐỊNH MỨC THIẾT KẾ PHẦN THIẾT BỊ
        // Áp dụng khi chi phí thiết bị công trình >= 50% (G_XD + G_TB) trong dự toán
        // Căn cứ tính: Chi phí thiết bị (chưa GTGT)
        // Mốc quy mô: <=5, 15, 25, 50, 100, 200, 500, 1000, >=3000 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoDD1 = new decimal[] { 5m, 15m, 25m, 50m, 100m, 200m, 500m, 1000m, 3000m };
        public static readonly decimal[] TiLeDD1 = new decimal[] { 0.60m, 0.50m, 0.45m, 0.40m, 0.36m, 0.33m, 0.28m, 0.22m, 0.16m };

        /// <summary>
        /// Hệ số điều chỉnh cấp công trình dùng khi chưa tách được bảng định mức riêng theo cấp.
        /// Bảng 2.7 - 2.16 tra trực tiếp theo cấp nên không nhân hệ số này nữa.
        /// </summary>
        public static decimal GetHeSoCapCongTrinh(string capCT)
        {
            if (string.IsNullOrEmpty(capCT)) return 1.0m;
            if (capCT.Contains("đặc biệt") || capCT.Contains("ĐB") || capCT.Contains("Đặc biệt")) return 1.45m;
            if (capCT.Contains("I") && !capCT.Contains("II") && !capCT.Contains("III") && !capCT.Contains("IV")) return 1.25m;
            if (capCT.Contains("II") && !capCT.Contains("III")) return 1.12m;
            if (capCT.Contains("III")) return 1.00m; // Cấp III là chuẩn mốc
            if (capCT.Contains("IV")) return 0.85m;
            return 1.0m;
        }

        /// <summary>
        /// Chuẩn hóa chuỗi cấp công trình về một trong các key của bảng 2.7 - 2.16.
        /// Trả về "Cấp III" nếu không nhận diện được.
        /// </summary>
        public static string ChuanHoaCapCongTrinh(string capCT)
        {
            if (string.IsNullOrWhiteSpace(capCT)) return CapIII;
            string c = capCT.Trim().ToLowerInvariant();
            if (c.Contains("đặc biệt") || c.Contains("đb") || c.Contains("dac biet")) return CapDacBiet;
            if (c.Contains("iv") || c.Contains("4")) return CapIV;
            if (c.Contains("iii") || c.Contains("3")) return CapIII;
            if (c.Contains("ii") || c.Contains("2")) return CapII;
            if (c.Contains("i") || c.Contains("1")) return CapI;
            return CapIII;
        }

        /// <summary>
        /// Chuẩn hóa tên loại công trình về một trong các khóa chuẩn dùng trong các bảng định mức:
        /// "Dân dụng", "Công nghiệp", "Giao thông", "Nông nghiệp & Môi trường", "Hạ tầng kỹ thuật".
        /// Công trình thủy lợi (kênh mương, đê kè, hồ đập, trạm bơm...) thuộc loại Nông nghiệp & Môi trường.
        /// </summary>
        public static string ChuanHoaLoaiCongTrinh(string loaiCT)
        {
            if (string.IsNullOrWhiteSpace(loaiCT)) return "Dân dụng";

            string s = loaiCT.Trim();
            string l = s.ToLowerInvariant();

            if (l.Contains("thủy lợi") || l.Contains("thuy loi")
                || l.Contains("nông nghiệp") || l.Contains("nong nghiep")
                || l.Contains("nn & ptn") || l.Contains("nn&ptn")
                || l.Contains("trạm bơm")
                || l.Contains("môi trường") || l.Contains("moi truong"))
                return LoaiNongNghiepMoiTruong;
            if (l.Contains("dân dụng") || l.Contains("dan dung"))
                return "Dân dụng";
            if (l.Contains("công nghiệp") || l.Contains("cong nghiep"))
                return "Công nghiệp";
            if (l.Contains("giao thông") || l.Contains("giao thong"))
                return "Giao thông";
            if (l.Contains("hạ tầng") || l.Contains("ha tang"))
                return "Hạ tầng kỹ thuật";

            return s;
        }
    }
}