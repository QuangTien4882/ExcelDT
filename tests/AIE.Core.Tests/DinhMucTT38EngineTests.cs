using System;
using System.Linq;
using Xunit;
using AIE.Core.Models;
using AIE.Core.Services;

namespace AIE.Core.Tests
{
    public class DinhMucTT38EngineTests
    {
        [Fact]
        public void Test_NoiSuy_QLDA_DanDung()
        {
            // Mốc quy mô QLDA (TT 38/2026 Bảng 1.1): <=10 tỷ (3.446%), <=20 tỷ (2.923%)
            // Tại 15 tỷ: Nt = 3.446 - ((3.446 - 2.923) / (20 - 10)) * (15 - 10) = 3.446 - (0.523 / 10) * 5 = 3.446 - 0.2615 = 3.1845%
            decimal quyMo15Ty = 15m;
            decimal tyLe = DinhMucChiPhiKhacBTC.NoiSuy(
                DinhMucTT38Database.MocQuyMoQLDA,
                DinhMucTT38Database.TiLeQLDA["Dân dụng"],
                quyMo15Ty);

            Assert.Equal(3.1845m, tyLe);
        }

        [Fact]
        public void Test_ThamTraQuyetToan_Clamp_And_Discounts()
        {
            // Mốc <= 5 tỷ tỷ lệ là 0.570%
            // Nếu quy mô 50 triệu (0.05 tỷ): 50.000.000 * 0.57% = 285.000 đồng (< 500.000 đồng)
            // Phải clamp lên tối thiểu 500.000 đồng
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Dân dụng",
                capCT: "Cấp III",
                soBuocTK: 2,
                chiPhiXD: 50_000_000m,
                chiPhiTB: 0m,
                chiPhiBT: 0m);

            var itemQuyetToan = model.Items.FirstOrDefault(x => x.MaChiPhi == "K_TT_QUYETTOAN");
            Assert.NotNull(itemQuyetToan);
            Assert.Equal(500_000m, itemQuyetToan.GiaTriTruocThue);

            // Test hệ số khi dự án đã kiểm toán độc lập (k = 0.5)
            model.DaKiemToanDocLap = true;
            model.ChiPhiXDTruocThue = 10_000_000_000m; // 10 tỷ
            DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(model);

            itemQuyetToan = model.Items.FirstOrDefault(x => x.MaChiPhi == "K_TT_QUYETTOAN");
            Assert.NotNull(itemQuyetToan);
            Assert.Equal(0.5m, itemQuyetToan.HeSoDieuChinh);
            // Tỷ lệ tại 10 tỷ: 0.390%. Tiền = 10 tỷ * 0.39% * 0.5 = 19.500.000 đồng
            Assert.Equal(19_500_000m, itemQuyetToan.GiaTriTruocThue);
        }

        [Fact]
        public void Test_KiemToanDocLap_Min1Trieu_And_ThietBiDiscount()
        {
            // Dự án có thiết bị >= 50%
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Công nghiệp",
                capCT: "Cấp II",
                soBuocTK: 2,
                chiPhiXD: 10_000_000_000m,
                chiPhiTB: 15_000_000_000m, // Thiết bị chiếm > 50%
                chiPhiBT: 0m);

            model.ThietBiTren50Pct = true;
            DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(model);

            var itemKT = model.Items.FirstOrDefault(x => x.MaChiPhi == "K_KT_DOCLAP");
            Assert.NotNull(itemKT);
            Assert.Equal(0.7m, itemKT.HeSoDieuChinh);
        }

        [Fact]
        public void Test_ThamDinhDuAn_Max150Trieu_Clamp()
        {
            // Dự án cực lớn: 50.000 tỷ
            // 50.000 tỷ * 0.001% = 500.000.000 đồng > Max 150.000.000 đồng -> Phải clamp về 150 triệu
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Giao thông",
                capCT: "Cấp I",
                soBuocTK: 3,
                chiPhiXD: 30_000_000_000_000m,
                chiPhiTB: 20_000_000_000_000m,
                chiPhiBT: 0m);

            var itemTD = model.Items.FirstOrDefault(x => x.MaChiPhi == "K_TD_DA");
            Assert.NotNull(itemTD);
            Assert.Equal(150_000_000m, itemTD.GiaTriTruocThue);
        }

        [Fact]
        public void Test_Dynamic_Add_Remove_Items()
        {
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh();
            int countBefore = model.Items.Count;

            // Thêm chi phí rà phá bom mìn
            var itemMoi = new ChiPhiKinhPhiItem
            {
                STT = "6.7",
                MaChiPhi = "K_BOM_MIN",
                TenChiPhi = "Chi phí rà phá bom mìn vật nổ",
                Nhom = NhomChiPhi.ChiPhiKhac,
                CachTinh = CachTinhChiPhi.NhapTruocThue,
                GiaTriTruocThue = 120_000_000m,
                ThueSuatGTGT = 0m,
                IsActive = true,
                IsUserAdded = true
            };
            model.Items.Add(itemMoi);
            model.TinhToanLai();

            Assert.Equal(countBefore + 1, model.Items.Count);
            Assert.Contains(model.Items, x => x.MaChiPhi == "K_BOM_MIN");

            // Xóa dòng vừa thêm
            model.Items.Remove(itemMoi);
            model.TinhToanLai();
            Assert.Equal(countBefore, model.Items.Count);
        }

        [Fact]
        public void Test_SwitchMode_THDT_vs_TMDT()
        {
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Dân dụng",
                capCT: "Cấp III",
                soBuocTK: 2,
                chiPhiXD: 10_000_000_000m,
                chiPhiTB: 2_000_000_000m,
                chiPhiBT: 1_000_000_000m
            );

            // Chế độ THDT (Bảng 2.1): G_BT ẩn/không dùng, Dự phòng 5%
            var itemBT = model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.BoiThuong_TDC);
            var itemDP = model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiDuPhong);
            Assert.NotNull(itemBT);
            Assert.NotNull(itemDP);

            itemBT.IsActive = false;
            itemDP.TyLePhanTram = 5.0m;
            model.TinhToanLai();

            decimal tongSauThueTHDT = model.TongSauThue;
            Assert.True(tongSauThueTHDT > 0);

            // Chuyển sang TMĐT (Bảng 1.2): G_BT bật, Dự phòng 10%
            itemBT.IsActive = true;
            itemDP.TyLePhanTram = 10.0m;
            model.TinhToanLai();

            decimal tongSauThueTMDT = model.TongSauThue;
            // TMĐT phải lớn hơn THDT do có thêm chi phí BT 1 tỷ và dự phòng tăng từ 5% lên 10%
            Assert.True(tongSauThueTMDT > tongSauThueTHDT);
        }

        [Fact]
        public void Test_Master_VAT_And_Row_Level_VAT()
        {
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Dân dụng",
                capCT: "Cấp III",
                soBuocTK: 2,
                chiPhiXD: 10_000_000_000m,
                chiPhiTB: 0m,
                chiPhiBT: 0m
            );

            // Đổi toàn bộ VAT sang 8% (trừ các khoản phí 0%)
            foreach (var item in model.Items)
            {
                if (item.MaChiPhi == "K_TD_DA" || item.MaChiPhi == "K_TT_QUYETTOAN") continue;
                item.ThueSuatGTGT = 0.08m;
            }
            model.TinhToanLai();

            var itemXD = model.Items.First(x => x.Nhom == NhomChiPhi.ChiPhiXayDung);
            Assert.Equal(0.08m, itemXD.ThueSuatGTGT);
            Assert.Equal(800_000_000m, itemXD.TienThueGTGT);
            Assert.Equal(10_800_000_000m, itemXD.GiaTriSauThue);

            // Chỉnh sửa từng dòng riêng lẻ (chuyển thiết kế sang 10%)
            var itemTK = model.Items.FirstOrDefault(x => x.MaChiPhi == "TV_TK");
            if (itemTK != null)
            {
                itemTK.ThueSuatGTGT = 0.10m;
                model.TinhToanLai();
                Assert.Equal(0.10m, itemTK.ThueSuatGTGT);
                Assert.Equal(Math.Round(itemTK.GiaTriTruocThue * 0.10m, 0), itemTK.TienThueGTGT);
            }
        }

        [Fact]
        public void Test_CdtTuQuanLy_Coefficient()
        {
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Dân dụng",
                capCT: "Cấp III",
                soBuocTK: 2,
                chiPhiXD: 10_000_000_000m,
                chiPhiTB: 0m,
                chiPhiBT: 0m
            );

            // Mặc định CdtTuQuanLy = false -> kQLDA = 1.0
            Assert.False(model.CdtTuQuanLy);
            DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(model);
            var itemQLDA = model.Items.First(x => x.Nhom == NhomChiPhi.QuanLyDuAn);
            Assert.Equal(1.0m, itemQLDA.HeSoDieuChinh);

            // Khi chọn CĐT tự QLDA -> kQLDA = 0.8
            model.CdtTuQuanLy = true;
            DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(model);
            Assert.Equal(0.8m, itemQLDA.HeSoDieuChinh);
        }

        [Fact]
        public void Test_EquipmentDependent_Costs()
        {
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Dân dụng",
                capCT: "Cấp III",
                soBuocTK: 2,
                chiPhiXD: 10_000_000_000m,
                chiPhiTB: 0m,
                chiPhiBT: 0m
            );

            // Không có thiết bị (chiPhiTB = 0): TV_GS_TB không active
            DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(model);
            var itemGSTB = model.Items.FirstOrDefault(x => x.MaChiPhi == "TV_GS_TB");
            Assert.NotNull(itemGSTB);
            Assert.False(itemGSTB.IsActive);

            // Có thiết bị (chiPhiTB > 0): TV_GS_TB tự động active
            model.ChiPhiTBTruocThue = 500_000_000m;
            DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(model);
            Assert.True(itemGSTB.IsActive);
        }

        [Fact]
        public void Test_TraCuu_QLDA_DoiChieu_DinhMuc()
        {
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Dân dụng",
                capCT: "Cấp III",
                soBuocTK: 2,
                chiPhiXD: 10_000_000_000m,
                chiPhiTB: 5_000_000_000m,
                chiPhiBT: 0m
            );
            DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(model);

            var tt = DinhMucTT38Engine.TraCuuThongTinChiPhiItem(model, "G_QLDA");
            Assert.NotNull(tt);
            Assert.Contains("1.1", tt.TenBang);
            Assert.Equal(7, tt.SoTrangPdf);
            Assert.Equal("G_XD + G_TB (trước thuế)", tt.CoSoTinh);
            // quy mô = (10 + 5) tỷ
            Assert.Equal(15m, tt.QuyMoTy);
            Assert.True(tt.TyLeNoiSuy > 0);

            // Tỷ lệ nội suy phải khớp đúng bảng định mức
            var tyLeKyVong = DinhMucChiPhiKhacBTC.NoiSuy(
                DinhMucTT38Database.MocQuyMoQLDA,
                DinhMucTT38Database.TiLeQLDA["Dân dụng"],
                15m);
            Assert.Equal(tyLeKyVong, tt.TyLeNoiSuy);

            // Tỷ lệ hiện tại trên dòng phải khớp tỷ lệ nội suy sau khi engine tính
            var itemQLDA = model.Items.First(x => x.MaChiPhi == "G_QLDA");
            Assert.Equal(itemQLDA.TyLePhanTram, tt.TyLeHienTai);
        }

        [Fact]
        public void Test_TraCuu_TVTK_3Buoc_Nhan155()
        {
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Giao thông",
                capCT: "Cấp II",
                soBuocTK: 3,
                chiPhiXD: 20_000_000_000m,
                chiPhiTB: 0m,
                chiPhiBT: 0m
            );
            DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(model);

            var tt = DinhMucTT38Engine.TraCuuThongTinChiPhiItem(model, "TV_TK");
            Assert.NotNull(tt);
            Assert.True(tt.SoTrangPdf > 0);
            Assert.True(tt.TiLeDinhMuc.Length > 0);
            // 3 bước => hệ số k phải ≥ 1 (kTV = 1) và tỷ lệ nội suy = định mức BVTC × 1.55
            var tyLeCoSo = DinhMucChiPhiKhacBTC.NoiSuy(tt.MocQuyMo, tt.TiLeDinhMuc, tt.QuyMoTy);
            Assert.Equal(Math.Round(tyLeCoSo * 1.55m, 4), tt.TyLeNoiSuy);
        }

        [Fact]
        public void Test_TraCuu_KhongCoDinhMuc_TraVeNull()
        {
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh();
            // G_XD nhập trực tiếp, không có định mức tỷ lệ TT38
            Assert.Null(DinhMucTT38Engine.TraCuuThongTinChiPhiItem(model, "G_XD"));
            Assert.Null(DinhMucTT38Engine.TraCuuThongTinChiPhiItem(model, "KHONG_TON_TAI"));
        }

        [Fact]
        public void Test_TraCuu_KTDDA_HeSoThueThamTra()
        {
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Dân dụng",
                capCT: "Cấp III",
                soBuocTK: 2,
                chiPhiXD: 10_000_000_000m,
                chiPhiTB: 0m,
                chiPhiBT: 0m
            );
            model.YeuCauThueThamTra = true;
            DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(model);

            var tt = DinhMucTT38Engine.TraCuuThongTinChiPhiItem(model, "K_TD_DA");
            Assert.NotNull(tt);
            Assert.Contains("Tổng mức đầu tư", tt.CoSoTinh);
            Assert.Equal(0.5m, tt.HeSo);
            Assert.True(tt.QuyMoTy > 0);
            Assert.True(tt.TyLeNoiSuy > 0);
        }

        [Fact]
        public void Test_BaoCaoKTKT_1Buoc_KhongTich_TVTK_KTDTK_KTDDT()
        {
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Dân dụng",
                capCT: "Cấp III",
                soBuocTK: 1,
                chiPhiXD: 5_000_000_000m,
                chiPhiTB: 0m,
                chiPhiBT: 0m
            );

            var itemTK = model.Items.FirstOrDefault(x => x.MaChiPhi == "TV_TK");
            var itemTDTK = model.Items.FirstOrDefault(x => x.MaChiPhi == "K_TD_TK");
            var itemTDDT = model.Items.FirstOrDefault(x => x.MaChiPhi == "K_TD_DT");

            Assert.NotNull(itemTK);
            Assert.False(itemTK.IsActive); // Chi phí thiết kế đã nằm trong Chi phí lập BCKT-KT

            Assert.NotNull(itemTDTK);
            Assert.False(itemTDTK.IsActive); // Phí thẩm định thiết kế đã nằm trong Phí thẩm định BCKT-KT

            Assert.NotNull(itemTDDT);
            Assert.False(itemTDDT.IsActive); // Phí thẩm định dự toán đã nằm trong Phí thẩm định BCKT-KT
        }

        [Fact]
        public void Test_BaoCaoKTKT_1Buoc_ThamTra_HeSo1_2()
        {
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Dân dụng",
                capCT: "Cấp III",
                soBuocTK: 1,
                chiPhiXD: 5_000_000_000m,
                chiPhiTB: 0m,
                chiPhiBT: 0m
            );

            var itemTTTK = model.Items.FirstOrDefault(x => x.MaChiPhi == "TV_TT_TK");
            var itemTTDT = model.Items.FirstOrDefault(x => x.MaChiPhi == "TV_TT_DT");

            Assert.NotNull(itemTTTK);
            Assert.Equal(1.2m, itemTTTK.HeSoDieuChinh); // Mục 4.4 TT 38/2026: k = 1,2

            Assert.NotNull(itemTTDT);
            Assert.Equal(1.2m, itemTTDT.HeSoDieuChinh); // Mục 4.4 TT 38/2026: k = 1,2

            // Kiểm tra qua TraCuuThongTinChiPhiItem
            var ttTK = DinhMucTT38Engine.TraCuuThongTinChiPhiItem(model, "TV_TT_TK");
            Assert.NotNull(ttTK);
            Assert.Equal(1.2m, ttTK.HeSo);
            Assert.Contains("Mục 4.4 TT 38/2026", ttTK.GhiChu);

            var ttDT = DinhMucTT38Engine.TraCuuThongTinChiPhiItem(model, "TV_TT_DT");
            Assert.NotNull(ttDT);
            Assert.Equal(1.2m, ttDT.HeSo);
            Assert.Contains("Mục 4.4 TT 38/2026", ttDT.GhiChu);
        }
    }
}
