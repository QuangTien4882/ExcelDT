using System;
using System.Linq;
using Xunit;
using AIE.Core.Models;
using AIE.Core.Services;
using AIE.Core.Services.Shared;

namespace AIE.Core.Tests
{
    public class ThuyLoiMappingTests
    {
        [Fact]
        public void Test_ThuyLoi_DinhMucBang_QLDA_Thuoc_PTNT()
        {
            // Công trình thủy lợi thuộc NN & PTNT: tỷ lệ QLDA phải lấy theo cột "Nông nghiệp & PTNT"
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Thủy lợi",
                capCT: "Cấp III",
                soBuocTK: 2,
                chiPhiXD: 10_000_000_000m,   // 10 tỷ
                chiPhiTB: 0m,
                chiPhiBT: 0m);

            var itemQLDA = model.Items.FirstOrDefault(x => x.MaChiPhi == "G_QLDA");
            Assert.NotNull(itemQLDA);
            // Bảng 1.1 TT 38/2026: Nông nghiệp & Môi trường tại 10 tỷ = 3.263%
            Assert.Equal(3.263m, itemQLDA.TyLePhanTram);
            Assert.True(itemQLDA.GiaTriTruocThue > 0);
        }

        [Fact]
        public void Test_ThuyLoi_DinhMucBang_ThietKe_Thuoc_PTNT()
        {
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Thủy lợi",
                capCT: "Cấp III",
                soBuocTK: 2,
                chiPhiXD: 10_000_000_000m);

            var itemTK = model.Items.FirstOrDefault(x => x.MaChiPhi == "TV_TK");
            Assert.NotNull(itemTK);
            // Bảng 2.7-2.16 TT 38/2026: Nông nghiệp & Môi trường cấp III tại 10 tỷ = 3.13%
            Assert.Equal(3.13m, itemTK.TyLePhanTram);
            Assert.True(itemTK.GiaTriTruocThue > 0);
        }

        [Fact]
        public void Test_ThuyLoi_SoKhop_NN_PTNT()
        {
            // Kết quả (QLDA, TK, thẩm tra, giám sát, thẩm định TK/DT) phải khớp nhau
            var modelThuyLoi = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Thủy lợi",
                capCT: "Cấp III",
                soBuocTK: 2,
                chiPhiXD: 10_000_000_000m,
                chiPhiTB: 1_000_000_000m,
                chiPhiBT: 0m);

            var modelPTNT = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Nông nghiệp & PTNT",
                capCT: "Cấp III",
                soBuocTK: 2,
                chiPhiXD: 10_000_000_000m,
                chiPhiTB: 1_000_000_000m,
                chiPhiBT: 0m);

            var maGoc = new[] { "G_QLDA", "TV_TK", "TV_TT_TK", "TV_TT_DT", "TV_GS_XD", "TV_GS_TB", "K_TD_TK", "K_TD_DT" };
            foreach (var ma in maGoc)
            {
                var iTL = modelThuyLoi.Items.First(x => x.MaChiPhi == ma);
                var iPT = modelPTNT.Items.First(x => x.MaChiPhi == ma);
                Assert.Equal(iPT.TyLePhanTram, iTL.TyLePhanTram);
                Assert.Equal(iPT.GiaTriTruocThue, iTL.GiaTriTruocThue);
            }
        }

        [Fact]
        public void Test_ThuyLoi_LayTiLeCPCTMDT_6_1()
        {
            // Bảng 3.2: NN&PTNT không phải hầm = 6.1%
            Assert.Equal(6.1m, InterpolationHelper.LayTiLeCPCTMDT("Thủy lợi"));
            Assert.Equal(6.1m, InterpolationHelper.LayTiLeCPCTMDT("Công trình Nông nghiệp và PTNT"));
        }

        [Fact]
        public void Test_ThuyLoi_LayTiLeTNCTTT_5_5()
        {
            // Bảng 3.6: Thủy lợi (NN&PTNT) = 5.5%
            Assert.Equal(5.5m, InterpolationHelper.LayTiLeTNCTTT("Thủy lợi"));
        }

        [Fact]
        public void Test_ChuanHoaLoaiCongTrinh_CacBienThe()
        {
            Assert.Equal(DinhMucTT38Database.LoaiNongNghiepMoiTruong, DinhMucTT38Database.ChuanHoaLoaiCongTrinh("Thủy lợi"));
            Assert.Equal(DinhMucTT38Database.LoaiNongNghiepMoiTruong, DinhMucTT38Database.ChuanHoaLoaiCongTrinh("Công ty Thủy lợi Bắc Kạn"));
            Assert.Equal(DinhMucTT38Database.LoaiNongNghiepMoiTruong, DinhMucTT38Database.ChuanHoaLoaiCongTrinh("Công trình Nông nghiệp và PTNT"));
            Assert.Equal(DinhMucTT38Database.LoaiNongNghiepMoiTruong, DinhMucTT38Database.ChuanHoaLoaiCongTrinh("Nông nghiệp và môi trường"));
            Assert.Equal("Dân dụng", DinhMucTT38Database.ChuanHoaLoaiCongTrinh("Dân dụng"));
            Assert.Equal("Giao thông", DinhMucTT38Database.ChuanHoaLoaiCongTrinh("Công trình Giao thông"));
            Assert.Equal("Công nghiệp", DinhMucTT38Database.ChuanHoaLoaiCongTrinh("Công nghiệp"));
            Assert.Equal("Hạ tầng kỹ thuật", DinhMucTT38Database.ChuanHoaLoaiCongTrinh("Hạ tầng kỹ thuật"));
        }

        [Fact]
        public void Test_ThamDinhDuAn_PhanTram_Theo_TMDT_KhongPhai_XDTB()
        {
            // Dự án có G_XD + G_TB = 10 tỷ nhưng G_BT lớn khiến TMĐT rơi vào mốc >15 tỷ
            // Tỷ lệ phí thẩm định dự án phải tra theo TMĐT (bảng "32. Tham dinh du an") chứ không phải G_XD+G_TB
            var modelKhongBT = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Dân dụng",
                capCT: "Cấp III",
                soBuocTK: 2,
                chiPhiXD: 8_000_000_000m,
                chiPhiTB: 2_000_000_000m,
                chiPhiBT: 0m);

            var modelCoBT = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Dân dụng",
                capCT: "Cấp III",
                soBuocTK: 2,
                chiPhiXD: 8_000_000_000m,
                chiPhiTB: 2_000_000_000m,
                chiPhiBT: 20_000_000_000m);

            // Bật chi phí bồi thường (mặc định tắt)
            var itemBT = modelCoBT.Items.First(x => x.MaChiPhi == "G_BT");
            itemBT.IsActive = true;
            DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(modelCoBT);

            var itemTDKhongBT = modelKhongBT.Items.First(x => x.MaChiPhi == "K_TD_DA");
            var itemTDCoBT = modelCoBT.Items.First(x => x.MaChiPhi == "K_TD_DA");

            // Không có BT: TMĐT gần 10 tỷ -> mốc 15 tỷ -> 0.019%
            Assert.Equal(0.019m, itemTDKhongBT.TyLePhanTram);

            // Có BT 20 tỷ: TMĐT lớn hơn hẳn -> tỷ lệ phải giảm xuống theo mốc lớn hơn
            Assert.True(itemTDCoBT.TyLePhanTram < itemTDKhongBT.TyLePhanTram);
            Assert.True(itemTDCoBT.TyLePhanTram < 0.019m);

            // Cơ sở tính = TMĐT (tổng các khoản trước dự phòng, trừ chính phí)
            decimal coSo = modelCoBT.Items
                .Where(x => x.IsActive && x.Id != itemTDCoBT.Id && x.Nhom != NhomChiPhi.ChiPhiDuPhong)
                .Sum(x => x.GiaTriTruocThue);
            decimal kq = Math.Round(coSo * (itemTDCoBT.TyLePhanTram / 100m) * 1.0m, 0);
            Assert.Equal(kq, itemTDCoBT.GiaTriTruocThue);
            Assert.True(coSo > 10_000_000_000m); // TMĐT phải > G_XD+G_TB
        }

        [Fact]
        public void Test_ThamDinhDuAn_ChiPhiTheoTMDT_Clamp_TienToi()
        {
            // Ở mốc nhỏ, cơ sở tính phải tính trên TMĐT (gồm cả QLDA/TV/K) chứ không phải riêng G_XD+G_TB
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Dân dụng",
                capCT: "Cấp III",
                soBuocTK: 2,
                chiPhiXD: 10_000_000_000m,
                chiPhiTB: 0m,
                chiPhiBT: 0m);

            var itemTD = model.Items.First(x => x.MaChiPhi == "K_TD_DA");
            // 10 tỷ XD -> TMĐT (XD + QLDA + TV + ...) > 10 tỷ -> mốc 15 tỷ -> 0.019%
            Assert.Equal(0.019m, itemTD.TyLePhanTram);

            var coSo = model.Items
                .Where(x => x.IsActive && x.Id != itemTD.Id && x.Nhom != NhomChiPhi.ChiPhiDuPhong)
                .Sum(x => x.GiaTriTruocThue);
            Assert.True(coSo > 10_000_000_000m); // Gồm QLDA + tư vấn
        }
    }
}