using System;
using System.Collections.Generic;
using System.Linq;
using AIE.Core.Models;

namespace AIE.Core.Services
{
    /// <summary>
    /// Engine tra cứu, nội suy tỷ lệ định mức và tính toán Tổng mức đầu tư (Bảng 1.2 TT 36)
    /// & Tổng hợp dự toán công trình (Bảng 2.1 TT 36) theo Thông tư 38/2026 và các văn bản BTC.
    /// </summary>
    public static class DinhMucTT38Engine
    {
        /// <summary>
        /// Tạo mới một BangTongHopKinhPhiModel với đầy đủ các danh mục chi phí chuẩn
        /// </summary>
        public static BangTongHopKinhPhiModel TaoBangKinhPhiMacDinh(
            string loaiCT = "Dân dụng",
            string capCT = "Cấp III",
            int soBuocTK = 2,
            decimal chiPhiXD = 10000000000m,   // 10 tỷ
            decimal chiPhiTB = 0m,
            decimal chiPhiBT = 0m,
            decimal chiPhiNhaTam = 0m)
        {
            var model = new BangTongHopKinhPhiModel
            {
                LoaiCongTrinh = loaiCT,
                CapCongTrinh = capCT,
                SoBuocThietKe = soBuocTK,
                ChiPhiXDTruocThue = chiPhiXD,
                ChiPhiNhaTamTruocThue = chiPhiNhaTam,
                ChiPhiTBTruocThue = chiPhiTB,
                ChiPhiBTTruocThue = chiPhiBT
            };

            model.Items = KhoiTaoDanhSachKhoanMucChuan();
            if (soBuocTK == 1)
            {
                var itemTK = model.Items.FirstOrDefault(x => x.MaChiPhi == "TV_TK");
                if (itemTK != null) itemTK.IsActive = false;
                var itemTDTK = model.Items.FirstOrDefault(x => x.MaChiPhi == "K_TD_TK");
                if (itemTDTK != null) itemTDTK.IsActive = false;
                var itemTDDT = model.Items.FirstOrDefault(x => x.MaChiPhi == "K_TD_DT");
                if (itemTDDT != null) itemTDDT.IsActive = false;
            }
            if (chiPhiTB > 0)
            {
                var itemGSTB = model.Items.FirstOrDefault(x => x.MaChiPhi == "TV_GS_TB");
                if (itemGSTB != null) itemGSTB.IsActive = true;
                var itemTB = model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiThietBi);
                if (itemTB != null) itemTB.IsActive = true;
            }
            CapNhatToanBoDinhMucVaTinhToan(model);
            return model;
        }

        /// <summary>
        /// Khởi tạo khung các khoản mục chi phí chuẩn theo Bảng 1.2 và Bảng 2.1
        /// </summary>
        public static List<ChiPhiKinhPhiItem> KhoiTaoDanhSachKhoanMucChuan()
        {
            var list = new List<ChiPhiKinhPhiItem>
            {
                // I. Chi phí bồi thường, hỗ trợ và tái định cư
                new ChiPhiKinhPhiItem
                {
                    STT = "I",
                    MaChiPhi = "G_BT",
                    TenChiPhi = "Chi phí bồi thường, hỗ trợ và tái định cư",
                    Nhom = NhomChiPhi.BoiThuong_TDC,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiBoiThuong,
                    KyHieu = "G_BT,TĐC",
                    ThueSuatGTGT = 0m,
                    GhiChuCachTinh = "Phương án bồi thường GPMB được duyệt",
                    IsActive = false, // Mặc định tắt nếu không có GPMB
                    IsReadOnly = true
                },

                // II. Chi phí xây dựng (Tách thành 2 mục theo quy định: Chi phí xây dựng & Chi phí nhà tạm)
                new ChiPhiKinhPhiItem
                {
                    STT = "2.1",
                    MaChiPhi = "G_XD",
                    TenChiPhi = "Chi phí xây dựng",
                    Nhom = NhomChiPhi.ChiPhiXayDung,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gxd",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng tổng hợp chi phí xây dựng (Dòng IV Bảng 3.8)",
                    IsActive = true,
                    IsReadOnly = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "2.2",
                    MaChiPhi = "G_NHA_TAM",
                    TenChiPhi = "Chi phí nhà tạm để ở và điều hành thi công",
                    Nhom = NhomChiPhi.ChiPhiXayDung,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gnt",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng tổng hợp chi phí xây dựng (Dòng VII Bảng 3.8)",
                    IsActive = true,
                    IsReadOnly = true
                },

                // III. Chi phí thiết bị
                new ChiPhiKinhPhiItem
                {
                    STT = "III",
                    MaChiPhi = "G_TB",
                    TenChiPhi = "Chi phí thiết bị",
                    Nhom = NhomChiPhi.ChiPhiThietBi,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiThietBi,
                    KyHieu = "Gtb",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Dự toán mua sắm thiết bị (Bảng 2.2)",
                    IsActive = true,
                    IsReadOnly = true
                },

                // IV. Chi phí quản lý dự án
                new ChiPhiKinhPhiItem
                {
                    STT = "IV",
                    MaChiPhi = "G_QLDA",
                    TenChiPhi = "Chi phí quản lý dự án",
                    Nhom = NhomChiPhi.QuanLyDuAn,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                    KyHieu = "Gqlda",
                    ThueSuatGTGT = 0m,
                    GhiChuCachTinh = "Bảng 1.1 Thông tư 38/2026/TT-BXD",
                    IsActive = true,
                    IsReadOnly = true
                },

                // V. Chi phí tư vấn đầu tư xây dựng
                new ChiPhiKinhPhiItem
                {
                    STT = "5.1",
                    MaChiPhi = "TV_KS",
                    TenChiPhi = "Chi phí khảo sát xây dựng",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gtv1",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Dự toán chi phí khảo sát riêng",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.2",
                    MaChiPhi = "TV_FS_KTKT",
                    TenChiPhi = "Chi phí lập Báo cáo nghiên cứu khả thi (FS) / Báo cáo KT-KT",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                    KyHieu = "Gtv2",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.2 / Bảng 2.4 Thông tư 38/2026",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.3",
                    MaChiPhi = "TV_TK",
                    TenChiPhi = "Chi phí thiết kế xây dựng công trình",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gtv3",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.7 - 2.16 Thông tư 38/2026",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.4",
                    MaChiPhi = "TV_TT_TK",
                    TenChiPhi = "Chi phí thẩm tra thiết kế xây dựng",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gtv4",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.19 Thông tư 38/2026",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.5",
                    MaChiPhi = "TV_TT_DT",
                    TenChiPhi = "Chi phí thẩm tra dự toán xây dựng",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gtv5",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.20 Thông tư 38/2026",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.6",
                    MaChiPhi = "TV_HSMT",
                    TenChiPhi = "Chi phí lập HSMT và đánh giá HSDT thi công xây dựng",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gtv6",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.21 - 2.23 Thông tư 38/2026",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.7",
                    MaChiPhi = "TV_GS_XD",
                    TenChiPhi = "Chi phí giám sát thi công xây dựng",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gtv7",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.24 Thông tư 38/2026",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.8",
                    MaChiPhi = "TV_GS_TB",
                    TenChiPhi = "Chi phí giám sát lắp đặt thiết bị",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiThietBi,
                    KyHieu = "Gtv8",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.25 Thông tư 38/2026",
                    IsActive = false // Bật khi có chi phí thiết bị
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.9",
                    MaChiPhi = "TV_BC_DX",
                    TenChiPhi = "Chi phí lập Báo cáo đề xuất chủ trương đầu tư",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                    KyHieu = "Gtv9",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.1 Thông tư 38/2026",
                    IsActive = false
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.10",
                    MaChiPhi = "TV_BC_TKT",
                    TenChiPhi = "Chi phí lập Báo cáo nghiên cứu tiền khả thi",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                    KyHieu = "Gtv10",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.2 Thông tư 38/2026",
                    IsActive = false
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.11",
                    MaChiPhi = "TV_KTKT_SUA_CHUA",
                    TenChiPhi = "Chi phí lập Báo cáo KT-KT (cải tạo, sửa chữa)",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                    KyHieu = "Gtv11",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.5 Thông tư 38/2026",
                    IsActive = false
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.12",
                    MaChiPhi = "TV_KTKT_NAO_VET",
                    TenChiPhi = "Chi phí lập Báo cáo KT-KT (nạo vét luồng)",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                    KyHieu = "Gtv12",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.6 Thông tư 38/2026",
                    IsActive = false
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.13",
                    MaChiPhi = "TV_TT_BC_TKT",
                    TenChiPhi = "Chi phí thẩm tra Báo cáo nghiên cứu tiền khả thi",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                    KyHieu = "Gtv13",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.17 Thông tư 38/2026",
                    IsActive = false
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.14",
                    MaChiPhi = "TV_TT_BC_KT",
                    TenChiPhi = "Chi phí thẩm tra Báo cáo nghiên cứu khả thi",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                    KyHieu = "Gtv14",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.18 Thông tư 38/2026",
                    IsActive = false
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.15",
                    MaChiPhi = "TV_HSMT_TV",
                    TenChiPhi = "Chi phí lập HSMT và đánh giá HSDT gói thầu tư vấn",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.TungKhoanMucRieng,
                    KyHieu = "Gtv15",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.21 Thông tư 38/2026",
                    IsActive = false
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.16",
                    MaChiPhi = "TV_HSMT_TB",
                    TenChiPhi = "Chi phí lập HSMT và đánh giá HSDT mua sắm vật tư, thiết bị",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiThietBi,
                    KyHieu = "Gtv16",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.23 Thông tư 38/2026",
                    IsActive = false
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.17",
                    MaChiPhi = "TV_GS_KS",
                    TenChiPhi = "Chi phí giám sát công tác khảo sát xây dựng",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.TungKhoanMucRieng,
                    KyHieu = "Gtv17",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.26 Thông tư 38/2026",
                    IsActive = false
                },

                // VI. Chi phí khác
                new ChiPhiKinhPhiItem
                {
                    STT = "6.1",
                    MaChiPhi = "K_TD_DA",
                    TenChiPhi = "Phí thẩm định dự án đầu tư xây dựng (hoặc Báo cáo KT-KT)",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.TongMucDauTu,
                    KyHieu = "Gk1",
                    ThueSuatGTGT = 0m, // Phí theo Luật Phí và Lệ phí
                    MinValue = 500000m,
                    MaxValue = 150000000m,
                    GhiChuCachTinh = "Thông tư BTC về phí thẩm định dự án",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "6.2",
                    MaChiPhi = "K_TD_TK",
                    TenChiPhi = "Phí thẩm định thiết kế xây dựng (sau TKCS)",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gk2",
                    ThueSuatGTGT = 0m,
                    MinValue = 500000m,
                    MaxValue = 150000000m,
                    GhiChuCachTinh = "Thông tư BTC về phí thẩm định thiết kế",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "6.3",
                    MaChiPhi = "K_TD_DT",
                    TenChiPhi = "Phí thẩm định dự toán xây dựng",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gk3",
                    ThueSuatGTGT = 0m,
                    MinValue = 500000m,
                    MaxValue = 150000000m,
                    GhiChuCachTinh = "Thông tư BTC về phí thẩm định dự toán",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "6.4",
                    MaChiPhi = "K_BH",
                    TenChiPhi = "Chi phí bảo hiểm công trình xây dựng",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    TyLePhanTram = 0.15m, // 0.15% cho dân dụng/giao thông
                    KyHieu = "Gk4",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Nghị định về bảo hiểm công trình",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "6.5",
                    MaChiPhi = "K_KT_DOCLAP",
                    TenChiPhi = "Chi phí kiểm toán độc lập",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                    KyHieu = "Gk5",
                    ThueSuatGTGT = 0.10m,
                    MinValue = 1000000m,
                    GhiChuCachTinh = "Quy định BTC về định mức chi phí kiểm toán độc lập",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "6.6",
                    MaChiPhi = "K_TT_QUYETTOAN",
                    TenChiPhi = "Chi phí thẩm tra, phê duyệt quyết toán",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                    KyHieu = "Gk6",
                    ThueSuatGTGT = 0m, // Phí ngân sách nhà nước
                    MinValue = 500000m,
                    GhiChuCachTinh = "Quy định BTC về định mức phí thẩm tra, phê duyệt quyết toán",
                    IsActive = true
                },

                // VII. Chi phí dự phòng
                new ChiPhiKinhPhiItem
                {
                    STT = "VII",
                    MaChiPhi = "G_DP",
                    TenChiPhi = "Chi phí dự phòng (Yếu tố khối lượng phát sinh)",
                    Nhom = NhomChiPhi.ChiPhiDuPhong,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.TongChiPhiTruocDuPhong,
                    TyLePhanTram = 5.0m, // 5% cho THDT (Bảng 2.1), 10% cho TMĐT (Bảng 1.2)
                    KyHieu = "Gdp",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Khoản 1 Điều 3 TT 36/2026/TT-BXD",
                    IsActive = true,
                    IsReadOnly = true
                }
            };

            return list;
        }

        /// <summary>
        /// Cập nhật lại toàn bộ tỷ lệ % định mức, hệ số điều chỉnh và tính toán lại giá trị
        /// </summary>
        public static void CapNhatToanBoDinhMucVaTinhToan(BangTongHopKinhPhiModel model)
        {
            if (model == null) return;

            string loaiCT = DinhMucTT38Database.ChuanHoaLoaiCongTrinh(model.LoaiCongTrinh);
            string capCT = model.CapCongTrinh;
            int soBuocTK = model.SoBuocThietKe;

            decimal gXD = model.ChiPhiXDTruocThue;
            decimal gTB = model.ChiPhiTBTruocThue;
            decimal gXDTB = gXD + gTB;

            // Chuyển sang đơn vị tỷ đồng để tra định mức
            decimal qmXDTy = gXD / 1_000_000_000m;
            decimal qmTBTy = gTB / 1_000_000_000m;
            decimal qmXDTBTy = gXDTB / 1_000_000_000m;
            if (qmXDTy <= 0) qmXDTy = 10m;     // Mặc định mốc tối thiểu 10 tỷ
            if (qmTBTy <= 0) qmTBTy = 10m;
            if (qmXDTBTy <= 0) qmXDTBTy = 10m;

            // Hệ số điều chỉnh chung
            decimal kQLDA = 1.0m;
            if (model.CdtTuQuanLy) kQLDA *= 0.8m;
            if (model.VungKhoKhan) kQLDA *= 1.35m;
            if (model.TuyenQuaNhieuTinh) kQLDA *= 1.1m;

            decimal kTV = 1.0m;
            if (model.VungKhoKhan) kTV *= 1.35m;

            decimal kThietKe = kTV;
            if (model.CaiTaoSuaChua) kThietKe *= 1.15m;
            if (model.ThietKeLapLai) kThietKe *= 0.36m;

            // Hệ số điều chỉnh kiểm toán & quyết toán
            decimal kKiemToan = 1.0m;
            if (model.ThietBiTren50Pct) kKiemToan *= 0.7m;

            decimal kQuyetToan = 1.0m;
            if (model.ThietBiTren50Pct) kQuyetToan *= 0.7m;
            if (model.DaKiemToanDocLap) kQuyetToan *= 0.5m;

            // Hệ số thẩm định
            decimal kThamDinh = 1.0m;
            if (model.YeuCauThueThamTra) kThamDinh *= 0.5m;

            foreach (var item in model.Items)
            {
                if (item.CachTinh != CachTinhChiPhi.TheoTyLeDinhMuc)
                    continue;

                // Phí thẩm định dự án (K_TD_DA) tính sau vì tỷ lệ tra theo Tổng mức đầu tư (mốc quy mô theo TMĐT)
                if (item.MaChiPhi == "K_TD_DA")
                    continue;

                switch (item.MaChiPhi)
                {
                    case "G_QLDA":
                        if (DinhMucTT38Database.TiLeQLDA.TryGetValue(loaiCT, out var qldaArr))
                        {
                            item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoQLDA, qldaArr, qmXDTBTy);
                            item.HeSoDieuChinh = kQLDA;
                        }
                        break;

                    case "TV_FS_KTKT":
                        if (soBuocTK == 1 || qmXDTBTy <= 40m)
                        {
                            // Báo cáo KT-KT (Bảng 2.4) - tối thiểu 5.000.000 đồng
                            if (DinhMucTT38Database.TiLeBaoCaoKTKT.TryGetValue(loaiCT, out var ktktArr))
                            {
                                item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoBaoCaoKTKT, ktktArr, qmXDTBTy);
                                item.HeSoDieuChinh = kTV;
                                item.MinValue = 5000000m;
                                item.TenChiPhi = "Chi phí lập Báo cáo kinh tế - kỹ thuật";
                            }
                        }
                        else
                        {
                            // Lập FS (Bảng 2.3)
                            if (DinhMucTT38Database.TiLeLapFS.TryGetValue(loaiCT, out var fsArr))
                            {
                                item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoFS, fsArr, qmXDTBTy);
                                item.HeSoDieuChinh = kTV;
                                item.MinValue = null;
                                item.TenChiPhi = "Chi phí lập Báo cáo nghiên cứu khả thi (FS)";
                            }
                        }
                        break;

                    case "TV_TK":
                        if (soBuocTK == 1)
                        {
                            // Đối với dự án 1 bước (Báo cáo KT-KT), Chi phí lập Báo cáo KT-KT đã bao gồm chi phí thiết kế
                            item.IsActive = false;
                        }
                        else if (soBuocTK == 2 || soBuocTK == 3)
                        {
                            item.IsActive = true;
                        }
                        if (DinhMucTT38Database.TiLeThietKeKyThuat.TryGetValue(loaiCT, out var tkLoai))
                        {
                            string capChuan = DinhMucTT38Database.ChuanHoaCapCongTrinh(capCT);
                            if (soBuocTK == 3 && tkLoai.TryGetValue(capChuan, out var tkktArr))
                            {
                                // Thiết kế 3 bước: TKKT (Bảng 2.7, 2.9, 2.11, 2.13, 2.15) x 1,55
                                decimal tlBase = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoThietKe, tkktArr, qmXDTy);
                                item.TyLePhanTram = Math.Round(tlBase * 1.55m, 4);
                            }
                            else if (DinhMucTT38Database.TiLeThietKeBVTC.TryGetValue(loaiCT, out var tkbvLoai)
                                     && tkbvLoai.TryGetValue(capChuan, out var tkbvArr))
                            {
                                // Thiết kế 1 hoặc 2 bước: TKBVTC (Bảng 2.8, 2.10, 2.12, 2.14, 2.16)
                                decimal tlBase = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoThietKe, tkbvArr, qmXDTy);
                                item.TyLePhanTram = Math.Round(tlBase, 4);
                            }
                            item.HeSoDieuChinh = kThietKe;
                        }
                        break;

                    case "TV_TT_TK":
                        if (DinhMucTT38Database.TiLeThamTraTK.TryGetValue(loaiCT, out var tttkArr))
                        {
                            item.TyLePhanTram = Math.Round(DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoThamTraTK, tttkArr, qmXDTy), 4);
                            decimal kTTTK = kTV;
                            if (soBuocTK == 1) kTTTK *= 1.2m; // Mục 4.4 TT 38/2026: Thẩm tra Báo cáo KT-KT điều chỉnh hệ số k = 1,2
                            item.HeSoDieuChinh = kTTTK;
                            item.MinValue = 2000000m;
                            item.GhiChuCachTinh = soBuocTK == 1 ? "Bảng 2.19 TT 38/2026 (k=1,2 Mục 4.4)" : "Bảng 2.19 Thông tư 38/2026";
                        }
                        break;

                    case "TV_TT_DT":
                        if (DinhMucTT38Database.TiLeThamTraDuToan.TryGetValue(loaiCT, out var ttdtArr))
                        {
                            item.TyLePhanTram = Math.Round(DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoThamTraTK, ttdtArr, qmXDTy), 4);
                            decimal kTTDT = kTV;
                            if (soBuocTK == 1) kTTDT *= 1.2m; // Mục 4.4 TT 38/2026: Thẩm tra Báo cáo KT-KT điều chỉnh hệ số k = 1,2
                            if (gXDTB > 0 && gTB / gXDTB >= 0.25m) kTTDT *= 1.2m;
                            item.HeSoDieuChinh = kTTDT;
                            item.MinValue = 2000000m;
                            item.GhiChuCachTinh = soBuocTK == 1 ? "Bảng 2.20 TT 38/2026 (k=1,2 Mục 4.4)" : "Bảng 2.20 Thông tư 38/2026";
                        }
                        break;

                    case "TV_HSMT":
                        if (DinhMucTT38Database.TiLeHSMTThiCong.TryGetValue(loaiCT, out var hsmArr))
                        {
                            item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoHSMT, hsmArr, qmXDTy);
                            item.HeSoDieuChinh = kTV;
                        }
                        break;

                    case "TV_GS_XD":
                        if (DinhMucTT38Database.TiLeGiamSatThiCong.TryGetValue(loaiCT, out var gsArr))
                        {
                            item.TyLePhanTram = Math.Round(DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoGiamSat, gsArr, qmXDTy), 4);
                            item.HeSoDieuChinh = kTV;
                            item.MinValue = 2000000m;
                        }
                        break;

                    case "TV_GS_TB":
                        if (DinhMucTT38Database.TiLeGiamSatLapDatTB.TryGetValue(loaiCT, out var gstbArr))
                        {
                            item.TyLePhanTram = Math.Round(DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoGiamSat, gstbArr, qmTBTy), 4);
                            item.HeSoDieuChinh = kTV;
                            if (gTB > 0) item.IsActive = true;
                        }
                        break;

                    case "TV_BC_DX":
                        if (DinhMucTT38Database.TiLeBCDeXuat.TryGetValue(loaiCT, out var bcDxArr))
                        {
                            item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoBCDeXuat, bcDxArr, qmXDTBTy);
                            item.HeSoDieuChinh = kTV;
                        }
                        break;

                    case "TV_BC_TKT":
                        if (DinhMucTT38Database.TiLeTienKhaThi.TryGetValue(loaiCT, out var tctArr))
                        {
                            item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoTienKhaThi, tctArr, qmXDTBTy);
                            item.HeSoDieuChinh = kTV;
                        }
                        break;

                    case "TV_KTKT_SUA_CHUA":
                        if (DinhMucTT38Database.TiLeKTKTSuaChua.TryGetValue(loaiCT, out var ktscArr))
                        {
                            item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoKTKTSuaChua, ktscArr, qmXDTBTy);
                            item.HeSoDieuChinh = kTV;
                            item.MinValue = 5000000m;
                        }
                        break;

                    case "TV_KTKT_NAO_VET":
                        if (DinhMucTT38Database.TiLeKTKTNaoVet.TryGetValue(loaiCT, out var naoVetArr))
                        {
                            item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoKTKTNaoVet, naoVetArr, qmXDTBTy);
                            item.HeSoDieuChinh = kTV;
                            item.MinValue = 5000000m;
                        }
                        break;

                    case "TV_TT_BC_TKT":
                        if (DinhMucTT38Database.TiLeLapTT_BCNCTienKhaThi.TryGetValue(loaiCT, out var ttTktArr))
                        {
                            item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoTienKhaThi, ttTktArr, qmXDTBTy);
                            item.HeSoDieuChinh = kTV;
                        }
                        break;

                    case "TV_TT_BC_KT":
                        if (DinhMucTT38Database.TiLeLapTT_BCNCKhaThi.TryGetValue(loaiCT, out var ttKtArr))
                        {
                            item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoTienKhaThi, ttKtArr, qmXDTBTy);
                            item.HeSoDieuChinh = kTV;
                        }
                        break;

                    case "TV_HSMT_TV":
                        {
                            decimal quyMoTuVanTy = LayTongTien(model, NhomChiPhi.TuVanDauTuXD) / 1_000_000_000m;
                            if (quyMoTuVanTy <= 0) quyMoTuVanTy = 1m;
                            item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoHSMTTuVan, DinhMucTT38Database.TiLeHSMTTuVan, quyMoTuVanTy);
                            item.HeSoDieuChinh = kTV;
                        }
                        break;

                    case "TV_HSMT_TB":
                        if (DinhMucTT38Database.TiLeHSMTCuaTB.TryGetValue(loaiCT, out var hsmTbArr))
                        {
                            item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoHSMTCuaTB, hsmTbArr, qmTBTy);
                            item.HeSoDieuChinh = kTV;
                        }
                        break;

                    case "TV_GS_KS":
                        {
                            var itemKS = model.Items.FirstOrDefault(x => x.MaChiPhi == "TV_KS");
                            decimal quyMoKhaoSatTy = (itemKS?.GiaTriTruocThue ?? 0m) / 1_000_000_000m;
                            if (quyMoKhaoSatTy <= 0) quyMoKhaoSatTy = 1m;
                            item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoGiamSatKhaoSat, DinhMucTT38Database.TiLeGiamSatKhaoSat, quyMoKhaoSatTy);
                            item.HeSoDieuChinh = kTV;
                        }
                        break;

                    case "K_TD_TK":
                        if (soBuocTK == 1)
                        {
                            // Đối với dự án 1 bước (Báo cáo KT-KT), Phí thẩm định Báo cáo KT-KT đã bao gồm thẩm định thiết kế
                            item.IsActive = false;
                        }
                        else if (soBuocTK == 2 || soBuocTK == 3)
                        {
                            item.IsActive = true;
                        }
                        if (DinhMucChiPhiKhacBTC.TiLeThamDinhThietKe.TryGetValue(loaiCT, out var tdtkArr))
                        {
                            item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(
                                DinhMucChiPhiKhacBTC.MocQuyMoThamDinhTKDT,
                                tdtkArr,
                                qmXDTy);
                            item.HeSoDieuChinh = kThamDinh;
                        }
                        break;

                    case "K_TD_DT":
                        if (soBuocTK == 1)
                        {
                            // Đối với dự án 1 bước (Báo cáo KT-KT), Phí thẩm định Báo cáo KT-KT đã bao gồm thẩm định dự toán
                            item.IsActive = false;
                        }
                        else if (soBuocTK == 2 || soBuocTK == 3)
                        {
                            item.IsActive = true;
                        }
                        if (DinhMucChiPhiKhacBTC.TiLeThamDinhDuToan.TryGetValue(loaiCT, out var tddtArr))
                        {
                            item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(
                                DinhMucChiPhiKhacBTC.MocQuyMoThamDinhTKDT,
                                tddtArr,
                                qmXDTy);
                            item.HeSoDieuChinh = kThamDinh;
                        }
                        break;

                    case "K_BH":
                        item.TyLePhanTram = loaiCT.Contains("Giao thông") ? 0.15m : 0.10m;
                        break;

                    case "K_KT_DOCLAP":
                        item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(
                            DinhMucChiPhiKhacBTC.MocQuyMoKiemToan,
                            DinhMucChiPhiKhacBTC.TiLeKiemToanDocLap,
                            qmXDTBTy);
                        item.HeSoDieuChinh = kKiemToan;
                        break;

                    case "K_TT_QUYETTOAN":
                        item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(
                            DinhMucChiPhiKhacBTC.MocQuyMoQuyetToan,
                            DinhMucChiPhiKhacBTC.TiLeThamTraQuyetToan,
                            qmXDTBTy);
                        item.HeSoDieuChinh = kQuyetToan;
                        break;

                    case "G_DP":
                        // Giữ nguyên tỷ lệ người dùng cài (5% cho THDT, 10% cho TMĐT)
                        break;
                }
            }

            // Tính toán trước để có Tổng mức đầu tư
            model.TinhToanLai();

            // =========================================================================
            // Phí thẩm định dự án (K_TD_DA): tỷ lệ % tra theo Tổng mức đầu tư dự án (Bảng "32. Tham dinh du an")
            // Mốc quy mô theo TMĐT (tỷ đồng), không phải G_XD + G_TB
            // =========================================================================
            CapNhatTiLeThamDinhDuAn(model, kThamDinh, model.YeuCauThueThamTra);

            model.TinhToanLai();
        }

        /// <summary>
        /// Tự động nội suy lại tỷ lệ % phí thẩm định dự án (K_TD_DA) theo Tổng mức đầu tư hiện tại.
        /// Dùng riêng để đồng bộ K_TD_DA khi người dùng sửa trực tiếp trên lưới Tab 2 (bật/tắt
        /// hạng mục, đổi tỷ lệ/hệ số...) mà không làm thay đổi toàn bộ định mức còn lại.
        /// </summary>
        public static void CapNhatTiLeThamDinhDuAn(BangTongHopKinhPhiModel model, decimal? kThamDinh = null, bool yeuCauThueThamTra = false)
        {
            if (model == null) return;
            model.TinhToanLai();

            if (kThamDinh == null)
            {
                kThamDinh = 1.0m;
                if (yeuCauThueThamTra) kThamDinh *= 0.5m;
            }

            var itemTDDA = model.Items.FirstOrDefault(x => x.MaChiPhi == "K_TD_DA");
            if (itemTDDA != null && itemTDDA.IsActive && itemTDDA.CachTinh == CachTinhChiPhi.TheoTyLeDinhMuc)
            {
                // Cơ sở tính = TMĐT trước dự phòng, trừ chính khoản phí (khớp bước 3 trong TinhToanLai)
                decimal coSoTDDA = model.Items
                    .Where(x => x.IsActive && x.Id != itemTDDA.Id && x.Nhom != NhomChiPhi.ChiPhiDuPhong)
                    .Sum(x => x.GiaTriTruocThue);
                decimal qmTMDT = coSoTDDA / 1_000_000_000m;
                if (qmTMDT <= 0) qmTMDT = 15m; // Mặc định mốc tối thiểu 15 tỷ
                itemTDDA.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(
                    DinhMucChiPhiKhacBTC.MocQuyMoThamDinhDuAn,
                    DinhMucChiPhiKhacBTC.TiLeThamDinhDuAn,
                    qmTMDT);
                itemTDDA.HeSoDieuChinh = kThamDinh.Value;
            }

            model.TinhToanLai();
        }

        // =========================================================================
        //  TRA CỨU ĐỊNH MỨC (cho popup Kiểm tra định mức - Phần D)
        // =========================================================================

        /// <summary>
        /// Trả về tổng giá trị trước thuế (đồng) của tất cả khoản mục đang kích hoạt thuộc nhóm chi phí cho trước.
        /// </summary>
        private static decimal LayTongTien(BangTongHopKinhPhiModel model, NhomChiPhi nhom)
        {
            if (model?.Items == null) return 0m;
            return model.Items
                .Where(x => x.IsActive && x.Nhom == nhom)
                .Sum(x => x.GiaTriTruocThue);
        }

        /// <summary>
        /// Tra cứu thông tin định mức TT38/2026 của một khoản mục chi phí.
        /// Trả null nếu khoản mục không có định mức áp dụng hoặc model null.
        /// Dùng cho popup "Kiểm tra định mức" (Phần D) và kiểm tra Thẩm định.
        /// </summary>
        public static ThongTinKiemTraDinhMuc? TraCuuThongTinChiPhiItem(BangTongHopKinhPhiModel model, string maChiPhi)
        {
            if (model == null || string.IsNullOrWhiteSpace(maChiPhi)) return null;

            var item = model.Items.FirstOrDefault(x => x.MaChiPhi.Equals(maChiPhi, StringComparison.OrdinalIgnoreCase));
            if (item == null) return null;

            string loaiCT = DinhMucTT38Database.ChuanHoaLoaiCongTrinh(model.LoaiCongTrinh);
            string capCT = model.CapCongTrinh;
            int soBuocTK = model.SoBuocThietKe;
            decimal gXD  = model.ChiPhiXDTruocThue;
            decimal gTB  = model.ChiPhiTBTruocThue;
            decimal gXDTB = gXD + gTB;
            decimal qmXDTy    = gXD  / 1_000_000_000m; if (qmXDTy    <= 0) qmXDTy    = 10m;
            decimal qmTBTy    = gTB  / 1_000_000_000m; if (qmTBTy    <= 0) qmTBTy    = 10m;
            decimal qmXDTBTy  = gXDTB / 1_000_000_000m; if (qmXDTBTy <= 0) qmXDTBTy  = 10m;

            decimal kTV = 1.0m;
            if (model.VungKhoKhan) kTV *= 1.35m;

            var r = new ThongTinKiemTraDinhMuc
            {
                MaChiPhi    = item.MaChiPhi,
                TenChiPhi   = item.TenChiPhi,
                TyLeHienTai = item.TyLePhanTram,
                HeSo        = item.HeSoDieuChinh,
                LoaiCongTrinh = loaiCT,
                CapCongTrinh  = DinhMucTT38Database.ChuanHoaCapCongTrinh(capCT)
            };

            switch (maChiPhi)
            {
                // ===== BẢNG 1.1: QLDA =====
                case "G_QLDA":
                    r.TenBang     = "Bảng 1.1 - Chi phí quản lý dự án";
                    r.SoTrangPdf  = BangToPdfPageMap.LaySoTrang("1.1");
                    r.MocQuyMo    = DinhMucTT38Database.MocQuyMoQLDA;
                    r.CoSoTinh    = "G_XD + G_TB (trước thuế)";
                    r.QuyMoTy     = qmXDTBTy;
                    if (DinhMucTT38Database.TiLeQLDA.TryGetValue(loaiCT, out var qlArr))
                    { r.TiLeDinhMuc = qlArr; r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, qlArr, qmXDTBTy); }
                    break;

                // ===== BẢNG 2.4 / 2.3: TV_FS_KTKT =====
                case "TV_FS_KTKT":
                    if (soBuocTK == 1 || qmXDTBTy <= 40m)
                    {
                        r.TenBang    = "Bảng 2.4 - Báo cáo kinh tế - kỹ thuật";
                        r.SoTrangPdf = BangToPdfPageMap.LaySoTrang("2.4");
                        r.MocQuyMo   = DinhMucTT38Database.MocQuyMoBaoCaoKTKT;
                        r.QuyMoTy    = qmXDTBTy;
                        r.GhiChu     = "Áp dụng khi 1 bước hoặc quy mô ≤ 40 tỷ; tối thiểu 5.000.000đ";
                        if (DinhMucTT38Database.TiLeBaoCaoKTKT.TryGetValue(loaiCT, out var ktArr))
                        { r.TiLeDinhMuc = ktArr; r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, ktArr, qmXDTBTy); }
                    }
                    else
                    {
                        r.TenBang    = "Bảng 2.3 - Báo cáo nghiên cứu khả thi (FS)";
                        r.SoTrangPdf = BangToPdfPageMap.LaySoTrang("2.3");
                        r.MocQuyMo   = DinhMucTT38Database.MocQuyMoFS;
                        r.QuyMoTy    = qmXDTBTy;
                        if (DinhMucTT38Database.TiLeLapFS.TryGetValue(loaiCT, out var fsArr))
                        { r.TiLeDinhMuc = fsArr; r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, fsArr, qmXDTBTy); }
                    }
                    r.CoSoTinh = "G_XD + G_TB (trước thuế)";
                    r.HeSo     = kTV;
                    break;

                // ===== BẢNG 2.7–2.16: TV_TK =====
                case "TV_TK":
                {
                    string capChuan = DinhMucTT38Database.ChuanHoaCapCongTrinh(capCT);
                    r.MocQuyMo   = DinhMucTT38Database.MocQuyMoThietKe;
                    r.CoSoTinh   = "G_XD (trước thuế)";
                    r.QuyMoTy    = qmXDTy;
                    decimal kTK  = kTV * (model.CaiTaoSuaChua ? 1.15m : 1.0m) * (model.ThietKeLapLai ? 0.36m : 1.0m);
                    r.HeSo       = kTK;
                    if (soBuocTK == 3 && DinhMucTT38Database.TiLeThietKeKyThuat.TryGetValue(loaiCT, out var tkLoai)
                        && tkLoai.TryGetValue(capChuan, out var tkktArr))
                    {
                        r.TenBang      = "Bảng 2.7/2.9/2.11/2.13/2.15 - TKKT × 1,55";
                        r.SoTrangPdf   = BangToPdfPageMap.LaySoTrang("2.7");
                        r.TiLeDinhMuc  = tkktArr;
                        r.TyLeNoiSuy   = Math.Round(DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, tkktArr, qmXDTy) * 1.55m, 4);
                        r.GhiChu       = $"Cấp {capChuan}, 3 bước → nội suy × 1,55";
                    }
                    else if (DinhMucTT38Database.TiLeThietKeBVTC.TryGetValue(loaiCT, out var bvLoai)
                             && bvLoai.TryGetValue(capChuan, out var bvArr))
                    {
                        r.TenBang      = "Bảng 2.8/2.10/2.12/2.14/2.16 - TKBVTC";
                        r.SoTrangPdf   = BangToPdfPageMap.LaySoTrang("2.8");
                        r.TiLeDinhMuc  = bvArr;
                        r.TyLeNoiSuy   = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, bvArr, qmXDTy);
                        r.GhiChu       = $"Cấp {capChuan}, {(soBuocTK == 3 ? "3 bước (BVTC)" : soBuocTK + " bước")}";
                    }
                    break;
                }

                // ===== BẢNG 2.19: TV_TT_TK =====
                case "TV_TT_TK":
                    r.TenBang     = "Bảng 2.19 - Thẩm tra thiết kế";
                    r.SoTrangPdf  = BangToPdfPageMap.LaySoTrang("2.19");
                    r.MocQuyMo    = DinhMucTT38Database.MocQuyMoThamTraTK;
                    r.CoSoTinh    = "G_XD (trước thuế)";
                    decimal kTK_TT = kTV;
                    if (soBuocTK == 1) kTK_TT *= 1.2m; // Mục 4.4 TT 38/2026: BCKT-KT k=1,2
                    r.HeSo        = kTK_TT;
                    r.GhiChu      = soBuocTK == 1 ? "Mục 4.4 TT 38/2026: Báo cáo KT-KT điều chỉnh k=1,2; tối thiểu 2.000.000đ" : "Tối thiểu 2.000.000đ";
                    r.QuyMoTy     = qmXDTy;
                    if (DinhMucTT38Database.TiLeThamTraTK.TryGetValue(loaiCT, out var tttkArr))
                    { r.TiLeDinhMuc = tttkArr; r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, tttkArr, qmXDTy); }
                    break;

                // ===== BẢNG 2.20: TV_TT_DT =====
                case "TV_TT_DT":
                {
                    r.TenBang     = "Bảng 2.20 - Thẩm tra dự toán";
                    r.SoTrangPdf  = BangToPdfPageMap.LaySoTrang("2.20");
                    r.MocQuyMo    = DinhMucTT38Database.MocQuyMoThamTraTK;
                    r.CoSoTinh    = "G_XD (trước thuế)";
                    decimal kD = kTV;
                    if (soBuocTK == 1) kD *= 1.2m; // Mục 4.4 TT 38/2026: BCKT-KT k=1,2
                    if (gXDTB > 0 && gTB / gXDTB >= 0.25m) kD *= 1.2m;
                    r.HeSo        = kD;
                    r.GhiChu      = soBuocTK == 1
                        ? "Mục 4.4 TT 38/2026: Báo cáo KT-KT điều chỉnh k=1,2; tối thiểu 2.000.000đ"
                        : ((gXDTB > 0 && gTB / gXDTB >= 0.25m)
                            ? "Tối thiểu 2.000.000đ; G_TB ≥ 25% tổng → k=1,2"
                            : "Tối thiểu 2.000.000đ");
                    r.QuyMoTy     = qmXDTy;
                    if (DinhMucTT38Database.TiLeThamTraDuToan.TryGetValue(loaiCT, out var tdArr))
                    { r.TiLeDinhMuc = tdArr; r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, tdArr, qmXDTy); }
                    break;
                }

                // ===== BẢNG 2.22: TV_HSMT =====
                case "TV_HSMT":
                    r.TenBang     = "Bảng 2.22 - HSMT & ĐG HSDT thi công";
                    r.SoTrangPdf  = BangToPdfPageMap.LaySoTrang("2.22");
                    r.MocQuyMo    = DinhMucTT38Database.MocQuyMoHSMT;
                    r.CoSoTinh    = "G_XD (trước thuế)";
                    r.HeSo        = kTV;
                    r.QuyMoTy     = qmXDTy;
                    if (DinhMucTT38Database.TiLeHSMTThiCong.TryGetValue(loaiCT, out var hsArr))
                    { r.TiLeDinhMuc = hsArr; r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, hsArr, qmXDTy); }
                    break;

                // ===== BẢNG 2.24: TV_GS_XD =====
                case "TV_GS_XD":
                    r.TenBang     = "Bảng 2.24 - Giám sát thi công";
                    r.SoTrangPdf  = BangToPdfPageMap.LaySoTrang("2.24");
                    r.MocQuyMo    = DinhMucTT38Database.MocQuyMoGiamSat;
                    r.CoSoTinh    = "G_XD (trước thuế)";
                    r.HeSo        = kTV;
                    r.GhiChu      = "Tối thiểu 2.000.000đ";
                    r.QuyMoTy     = qmXDTy;
                    if (DinhMucTT38Database.TiLeGiamSatThiCong.TryGetValue(loaiCT, out var gsArr))
                    { r.TiLeDinhMuc = gsArr; r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, gsArr, qmXDTy); }
                    break;

                // ===== BẢNG 2.25: TV_GS_TB =====
                case "TV_GS_TB":
                    r.TenBang     = "Bảng 2.25 - Giám sát lắp đặt thiết bị";
                    r.SoTrangPdf  = BangToPdfPageMap.LaySoTrang("2.25");
                    r.MocQuyMo    = DinhMucTT38Database.MocQuyMoGiamSat;
                    r.CoSoTinh    = "G_TB (trước thuế)";
                    r.HeSo        = kTV;
                    r.QuyMoTy     = qmTBTy;
                    if (DinhMucTT38Database.TiLeGiamSatLapDatTB.TryGetValue(loaiCT, out var gtArr))
                    { r.TiLeDinhMuc = gtArr; r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, gtArr, qmTBTy); }
                    break;

                // ===== BẢNG 2.1: TV_BC_DX =====
                case "TV_BC_DX":
                    r.TenBang     = "Bảng 2.1 - Đề xuất chủ trương";
                    r.SoTrangPdf  = BangToPdfPageMap.LaySoTrang("2.1");
                    r.MocQuyMo    = DinhMucTT38Database.MocQuyMoBCDeXuat;
                    r.CoSoTinh    = "G_XD + G_TB (trước thuế)";
                    r.HeSo        = kTV;
                    r.QuyMoTy     = qmXDTBTy;
                    if (DinhMucTT38Database.TiLeBCDeXuat.TryGetValue(loaiCT, out var bcArr))
                    { r.TiLeDinhMuc = bcArr; r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, bcArr, qmXDTBTy); }
                    break;

                // ===== BẢNG 2.2: TV_BC_TKT =====
                case "TV_BC_TKT":
                    r.TenBang     = "Bảng 2.2 - NC tiền khả thi";
                    r.SoTrangPdf  = BangToPdfPageMap.LaySoTrang("2.2");
                    r.MocQuyMo    = DinhMucTT38Database.MocQuyMoTienKhaThi;
                    r.CoSoTinh    = "G_XD + G_TB (trước thuế)";
                    r.HeSo        = kTV;
                    r.QuyMoTy     = qmXDTBTy;
                    if (DinhMucTT38Database.TiLeTienKhaThi.TryGetValue(loaiCT, out var tkttArr))
                    { r.TiLeDinhMuc = tkttArr; r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, tkttArr, qmXDTBTy); }
                    break;

                // ===== BẢNG 2.5: TV_KTKT_SUA_CHUA =====
                case "TV_KTKT_SUA_CHUA":
                    r.TenBang     = "Bảng 2.5 - KT-KT sửa chữa";
                    r.SoTrangPdf  = BangToPdfPageMap.LaySoTrang("2.5");
                    r.MocQuyMo    = DinhMucTT38Database.MocQuyMoKTKTSuaChua;
                    r.CoSoTinh    = "G_XD + G_TB (trước thuế)";
                    r.HeSo        = kTV;
                    r.GhiChu      = "Tối thiểu 5.000.000đ";
                    r.QuyMoTy     = qmXDTBTy;
                    if (DinhMucTT38Database.TiLeKTKTSuaChua.TryGetValue(loaiCT, out var ktscArr))
                    { r.TiLeDinhMuc = ktscArr; r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, ktscArr, qmXDTBTy); }
                    break;

                // ===== BẢNG 2.6: TV_KTKT_NAO_VET =====
                case "TV_KTKT_NAO_VET":
                    r.TenBang     = "Bảng 2.6 - KT-KT nạo vét";
                    r.SoTrangPdf  = BangToPdfPageMap.LaySoTrang("2.6");
                    r.MocQuyMo    = DinhMucTT38Database.MocQuyMoKTKTNaoVet;
                    r.CoSoTinh    = "G_XD + G_TB (trước thuế)";
                    r.HeSo        = kTV;
                    r.GhiChu      = "Chỉ áp dụng cho Giao thông; tối thiểu 5.000.000đ";
                    r.QuyMoTy     = qmXDTBTy;
                    if (DinhMucTT38Database.TiLeKTKTNaoVet.TryGetValue(loaiCT, out var nvArr))
                    { r.TiLeDinhMuc = nvArr; r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, nvArr, qmXDTBTy); }
                    break;

                // ===== BẢNG 2.17: TV_TT_BC_TKT =====
                case "TV_TT_BC_TKT":
                    r.TenBang     = "Bảng 2.17 - Thẩm tra BCNC tiền khả thi";
                    r.SoTrangPdf  = BangToPdfPageMap.LaySoTrang("2.17");
                    r.MocQuyMo    = DinhMucTT38Database.MocQuyMoTienKhaThi;
                    r.CoSoTinh    = "G_XD + G_TB (trước thuế)";
                    r.HeSo        = kTV;
                    r.QuyMoTy     = qmXDTBTy;
                    if (DinhMucTT38Database.TiLeLapTT_BCNCTienKhaThi.TryGetValue(loaiCT, out var ttTktArr))
                    { r.TiLeDinhMuc = ttTktArr; r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, ttTktArr, qmXDTBTy); }
                    break;

                // ===== BẢNG 2.18: TV_TT_BC_KT =====
                case "TV_TT_BC_KT":
                    r.TenBang     = "Bảng 2.18 - Thẩm tra BCNC khả thi";
                    r.SoTrangPdf  = BangToPdfPageMap.LaySoTrang("2.18");
                    r.MocQuyMo    = DinhMucTT38Database.MocQuyMoTienKhaThi;
                    r.CoSoTinh    = "G_XD + G_TB (trước thuế)";
                    r.HeSo        = kTV;
                    r.QuyMoTy     = qmXDTBTy;
                    if (DinhMucTT38Database.TiLeLapTT_BCNCKhaThi.TryGetValue(loaiCT, out var ttKtArr))
                    { r.TiLeDinhMuc = ttKtArr; r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, ttKtArr, qmXDTBTy); }
                    break;

                // ===== BẢNG 2.21: TV_HSMT_TV =====
                case "TV_HSMT_TV":
                    r.TenBang     = "Bảng 2.21 - HSMT tư vấn";
                    r.SoTrangPdf  = BangToPdfPageMap.LaySoTrang("2.21");
                    r.MocQuyMo    = DinhMucTT38Database.MocQuyMoHSMTTuVan;
                    r.TiLeDinhMuc = DinhMucTT38Database.TiLeHSMTTuVan;
                    r.CoSoTinh    = "Chi phí tư vấn gói thầu";
                    r.HeSo        = kTV;
                    { decimal qTV = LayTongTien(model, NhomChiPhi.TuVanDauTuXD) / 1_000_000_000m;
                      if (qTV <= 0) qTV = 1m; r.QuyMoTy = qTV;
                      r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, r.TiLeDinhMuc, qTV); }
                    break;

                // ===== BẢNG 2.23: TV_HSMT_TB =====
                case "TV_HSMT_TB":
                    r.TenBang     = "Bảng 2.23 - HSMT mua sắm TB";
                    r.SoTrangPdf  = BangToPdfPageMap.LaySoTrang("2.23");
                    r.MocQuyMo    = DinhMucTT38Database.MocQuyMoHSMTCuaTB;
                    r.CoSoTinh    = "G_TB (trước thuế)";
                    r.HeSo        = kTV;
                    r.QuyMoTy     = qmTBTy;
                    if (DinhMucTT38Database.TiLeHSMTCuaTB.TryGetValue(loaiCT, out var htArr))
                    { r.TiLeDinhMuc = htArr; r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, htArr, qmTBTy); }
                    break;

                // ===== BẢNG 2.26: TV_GS_KS =====
                case "TV_GS_KS":
                    r.TenBang     = "Bảng 2.26 - Giám sát khảo sát";
                    r.SoTrangPdf  = BangToPdfPageMap.LaySoTrang("2.26");
                    r.MocQuyMo    = DinhMucTT38Database.MocQuyMoGiamSatKhaoSat;
                    r.TiLeDinhMuc = DinhMucTT38Database.TiLeGiamSatKhaoSat;
                    r.CoSoTinh    = "Chi phí khảo sát";
                    r.HeSo        = kTV;
                    { var iKS = model.Items.FirstOrDefault(x => x.MaChiPhi == "TV_KS");
                      decimal qKS = (iKS?.GiaTriTruocThue ?? 0m) / 1_000_000_000m;
                      if (qKS <= 0) qKS = 1m; r.QuyMoTy = qKS;
                      r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, r.TiLeDinhMuc, qKS); }
                    break;

                // ===== Bảng 32 BTC: K_TD_DA =====
                case "K_TD_DA":
                    r.TenBang     = "Bảng 32 - Phí thẩm định dự án (BTC)";
                    r.SoTrangPdf  = 0;
                    r.MocQuyMo    = DinhMucChiPhiKhacBTC.MocQuyMoThamDinhDuAn;
                    r.TiLeDinhMuc = DinhMucChiPhiKhacBTC.TiLeThamDinhDuAn;
                    r.CoSoTinh    = "Tổng mức đầu tư (TMĐT)";
                    r.HeSo        = model.YeuCauThueThamTra ? 0.5m : 1.0m;
                    { decimal coSo = model.Items
                        .Where(x => x.IsActive && x.Id != item.Id && x.Nhom != NhomChiPhi.ChiPhiDuPhong)
                        .Sum(x => x.GiaTriTruocThue);
                      decimal qm = coSo / 1_000_000_000m;
                      if (qm <= 0) qm = 15m; r.QuyMoTy = qm;
                      r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, r.TiLeDinhMuc, qm); }
                    break;

                // ===== BTC: K_TD_TK, K_TD_DT, K_TT_QUYETTOAN, K_KT_DOCLAP =====
                case "K_TD_TK":
                    r.TenBang     = "Phí thẩm định thiết kế (BTC)";
                    r.SoTrangPdf  = 0;
                    r.MocQuyMo    = DinhMucChiPhiKhacBTC.MocQuyMoThamDinhTKDT;
                    r.TiLeDinhMuc = DinhMucChiPhiKhacBTC.TiLeThamDinhThietKe.TryGetValue(loaiCT, out var tdtkArr) ? tdtkArr : Array.Empty<decimal>();
                    r.CoSoTinh    = "G_XD (trước thuế)";
                    r.QuyMoTy     = qmXDTy;
                    if (r.TiLeDinhMuc.Length > 0) r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, r.TiLeDinhMuc, qmXDTy);
                    break;
                case "K_TD_DT":
                    r.TenBang     = "Phí thẩm định dự toán (BTC)";
                    r.SoTrangPdf  = 0;
                    r.MocQuyMo    = DinhMucChiPhiKhacBTC.MocQuyMoThamDinhTKDT;
                    r.TiLeDinhMuc = DinhMucChiPhiKhacBTC.TiLeThamDinhDuToan.TryGetValue(loaiCT, out var tddtArr2) ? tddtArr2 : Array.Empty<decimal>();
                    r.CoSoTinh    = "G_XD (trước thuế)";
                    r.QuyMoTy     = qmXDTy;
                    if (r.TiLeDinhMuc.Length > 0) r.TyLeNoiSuy = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, r.TiLeDinhMuc, qmXDTy);
                    break;
                case "K_TT_QUYETTOAN":
                    r.TenBang     = "Phí thẩm tra quyết toán (BTC)";
                    r.SoTrangPdf  = 0;
                    r.MocQuyMo    = DinhMucChiPhiKhacBTC.MocQuyMoQuyetToan;
                    r.TiLeDinhMuc = DinhMucChiPhiKhacBTC.TiLeThamTraQuyetToan;
                    r.CoSoTinh    = "G_XD + G_TB (trước thuế)";
                    r.QuyMoTy     = qmXDTBTy;
                    r.TyLeNoiSuy  = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, r.TiLeDinhMuc, qmXDTBTy);
                    break;
                case "K_KT_DOCLAP":
                    r.TenBang     = "Phí kiểm toán độc lập (BTC)";
                    r.SoTrangPdf  = 0;
                    r.MocQuyMo    = DinhMucChiPhiKhacBTC.MocQuyMoKiemToan;
                    r.TiLeDinhMuc = DinhMucChiPhiKhacBTC.TiLeKiemToanDocLap;
                    r.CoSoTinh    = "G_XD + G_TB (trước thuế)";
                    r.QuyMoTy     = qmXDTBTy;
                    r.TyLeNoiSuy  = DinhMucChiPhiKhacBTC.NoiSuy(r.MocQuyMo, r.TiLeDinhMuc, qmXDTBTy);
                    break;

                default:
                    return null;
            }
            return r;
        }

        /// <summary>
        /// Thư viện các khoản mục chi phí chuẩn để người dùng lựa chọn bổ sung vào dự toán
        /// </summary>
        public static List<ChiPhiKinhPhiItem> LayThuVienChiPhiChuan()
        {
            return new List<ChiPhiKinhPhiItem>
            {
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "TV_KHAO_SAT_DH",
                    TenChiPhi = "Chi phí khảo sát địa hình",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Dự toán khảo sát địa hình riêng"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "TV_KHAO_SAT_DC",
                    TenChiPhi = "Chi phí khảo sát địa chất công trình",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Dự toán khảo sát địa chất riêng"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "TV_DTM",
                    TenChiPhi = "Chi phí đánh giá tác động môi trường (ĐTM)",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Lập báo cáo ĐTM theo quy định Luật Môi trường"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "TV_THAM_DINH_GIA_TB",
                    TenChiPhi = "Chi phí thẩm định giá thiết bị",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiThietBi,
                    TyLePhanTram = 0.25m,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Chứng thư thẩm định giá thiết bị"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "TV_QUAN_TRAC",
                    TenChiPhi = "Chi phí quan trắc biến dạng công trình (lún, nghiêng)",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Phương án quan trắc biến dạng công trình"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "TV_THI_NGHIEM_CHUYEN_NGANH",
                    TenChiPhi = "Chi phí thí nghiệm chuyên ngành xây dựng (nén cọc, siêu âm bê tông)",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Đề cương thí nghiệm kiểm định chất lượng"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "K_RA_PHA_BOM_MIN",
                    TenChiPhi = "Chi phí rà phá bom mìn, vật nổ",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    ThueSuatGTGT = 0m,
                    GhiChuCachTinh = "Dự toán rà phá bom mìn do Quân khu/Bộ CHQS phê duyệt"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "K_KHOI_CONG_KHANH_THANH",
                    TenChiPhi = "Chi phí tổ chức khởi công, khánh thành công trình",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Theo dự toán được cấp có thẩm quyền phê duyệt"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "K_DI_DOI_HA_TANG",
                    TenChiPhi = "Chi phí di dời các công trình hạ tầng kỹ thuật (điện, nước, viễn thông)",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Dự toán di dời hoàn trả hạ tầng"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "K_DANG_KIEM_AN_TOAN",
                    TenChiPhi = "Chi phí đăng kiểm an toàn thiết bị, PCCC",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiThietBi,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Kiểm định an toàn và nghiệm thu PCCC"
                }
            };
        }
    }
}
