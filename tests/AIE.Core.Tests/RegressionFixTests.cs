using System.Linq;
using Xunit;
using AIE.Core.Models;
using AIE.Core.Services;
using AIE.Core.Services.Shared;

namespace AIE.Core.Tests
{
    /// <summary>
    /// Các bài kiểm tra hồi quy cho đợt sửa chữa 2026-09:
    /// - A1: Mặc định Báo cáo Thẩm định (QLDA/CPC/TNCTTT) lấy từ CSDL định mức TT38/36 (không hardcode).
    /// - A2: Khi sửa lưới Tab 2, chỉ K_TD_DA được nội suy lại, các mục khác giữ nguyên tỷ lệ/hệ số người dùng sửa.
    /// </summary>
    public class RegressionFixTests
    {
        // ================= A1: Mặc định Báo cáo Thẩm định khớp CSDL =================

        [Theory]
        [InlineData("Dân dụng", 3.446)]
        [InlineData("Công nghiệp", 3.557)]
        [InlineData("Giao thông", 3.024)]
        [InlineData("Hạ tầng kỹ thuật", 2.901)]
        public void BaoCao_QLDA_MacDinh_Khop_CSDL(string loai, double expected)
        {
            var norm = DinhMucTT38Database.ChuanHoaLoaiCongTrinh(loai);
            Assert.Equal((decimal)expected, DinhMucTT38Database.TiLeQLDA[norm][0]);
        }

        [Fact]
        public void BaoCao_QLDA_ThuyLoi_Khop_CSDL_NongNghiepPTNT()
        {
            var norm = DinhMucTT38Database.ChuanHoaLoaiCongTrinh("Thủy lợi");
            Assert.Equal(DinhMucTT38Database.LoaiNongNghiepMoiTruong, norm);
            Assert.Equal(3.263m, DinhMucTT38Database.TiLeQLDA[norm][0]);
        }

        [Theory]
        [InlineData("Dân dụng", 7.3)]
        [InlineData("Công nghiệp", 6.2)]
        [InlineData("Giao thông", 6.2)]
        [InlineData("Nông nghiệp & PTNT", 6.1)]
        [InlineData("Hạ tầng kỹ thuật", 5.5)]
        public void BaoCao_CPC_MacDinh_Khop_CSDL(string loai, double expected)
        {
            Assert.Equal((decimal)expected, InterpolationHelper.LayTiLeCPCTMDT(loai));
        }

        [Theory]
        [InlineData("Dân dụng", 5.5)]
        [InlineData("Công nghiệp", 6.0)]
        [InlineData("Giao thông", 6.0)]
        [InlineData("Nông nghiệp & PTNT", 5.5)]
        [InlineData("Hạ tầng kỹ thuật", 5.5)]
        public void BaoCao_TNCTTT_MacDinh_Khop_CSDL(string loai, double expected)
        {
            Assert.Equal((decimal)expected, InterpolationHelper.LayTiLeTNCTTT(loai));
        }

        [Fact]
        public void BaoCao_TNCTTT_ThuyLoi_Thuoc_NongNghiepPTNT()
        {
            Assert.Equal(5.5m, InterpolationHelper.LayTiLeTNCTTT("Thủy lợi"));
        }

        // ================= A2: Chỉ K_TD_DA được nội suy lại khi sửa lưới =================

        [Fact]
        public void SuaLuoi_ChiNoiSuyLaiKTieuDuAn_KhongDoiCacMucKhac()
        {
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Dân dụng",
                capCT: "Cấp III",
                soBuocTK: 2,
                chiPhiXD: 1_000_000_000_000m,   // 1.000 tỷ: TMĐT lớn -> mốc >=1000 tỷ
                chiPhiTB: 0m,
                chiPhiBT: 0m);

            // Mô phỏng người dùng sửa thủ công trên lưới: đổi tỷ lệ + hệ số của mục khác (VD: TV_TT_TK)
            var tttk = model.Items.First(x => x.MaChiPhi == "TV_TT_TK");
            tttk.TyLePhanTram = 9.99m;
            tttk.HeSoDieuChinh = 0.55m;

            // Sửa lưới gọi CapNhatTiLeThamDinhDuAn (không gọi lại toàn bộ CapNhatToanBoDinhMuc)
            DinhMucTT38Engine.CapNhatTiLeThamDinhDuAn(model);

            // Chỉ K_TD_DA bị nội suy lại theo TMĐT hiện tại
            var tddaSau = model.Items.First(x => x.MaChiPhi == "K_TD_DA").TyLePhanTram;
            Assert.True(tddaSau > 0m, "K_TD_DA phải được nội suy lại");

            // Không được đổi trạng thái đã sửa của các mục khác
            var tttkSau = model.Items.First(x => x.MaChiPhi == "TV_TT_TK");
            Assert.Equal(9.99m, tttkSau.TyLePhanTram);
            Assert.Equal(0.55m, tttkSau.HeSoDieuChinh);
        }

        [Fact]
        public void SuaLuoi_ThayDoiTMĐT_KDau_QuaMoc_DuocNoiSuyLai()
        {
            // TMĐT chỉ ~10 tỷ -> mốc <=15 tỷ -> 0.019%
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Dân dụng",
                capCT: "Cấp III",
                soBuocTK: 2,
                chiPhiXD: 10_000_000_000m,
                chiPhiTB: 0m,
                chiPhiBT: 0m);

            // Kích hoạt thêm chi phí bồi thường 20 tỷ -> TMĐT nhảy mốc
            var itemBT = model.Items.First(x => x.MaChiPhi == "G_BT");
            itemBT.IsActive = true;
            model.ChiPhiBTTruocThue = 20_000_000_000m;
            model.YeuCauThueThamTra = true;

            DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(model);

            var tdda = model.Items.First(x => x.MaChiPhi == "K_TD_DA");
            // TMĐT > 15 tỷ nên tỷ lệ nhỏ hơn mốc 0.019%
            Assert.True(tdda.TyLePhanTram < 0.019m);

            // Có yêu cầu thuê thẩm tra: hệ số điều chỉnh = 0.5
            Assert.Equal(0.5m, tdda.HeSoDieuChinh);
            Assert.True(tdda.GiaTriTruocThue > 0);
        }
    }
}