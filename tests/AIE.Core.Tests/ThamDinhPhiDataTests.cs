using System.Linq;
using Xunit;
using AIE.Core.Models;
using AIE.Core.Services;

namespace AIE.Core.Tests
{
    /// <summary>
    /// Kiểm tra dữ liệu định mức phí thẩm định khớp 3 file Excel BTC:
    /// "32. Tham dinh du an.xlsx", "33. Tham dinh thiet ke.xlsx", "34. Tham dinh du toan.xlsx"
    /// </summary>
    public class ThamDinhPhiDataTests
    {
        [Fact]
        public void Test_PhiThamDinhDuAn_Khop_File32()
        {
            // File 32: mốc quy mô theo TMĐT (tỷ đồng): <=15, 25, 50, 100, 200, 500, 1000, 2000, 5000, >=10000
            var mocs = DinhMucChiPhiKhacBTC.MocQuyMoThamDinhDuAn;
            var tyle = DinhMucChiPhiKhacBTC.TiLeThamDinhDuAn;

            Assert.Equal(new decimal[] { 15m, 25m, 50m, 100m, 200m, 500m, 1000m, 2000m, 5000m, 10000m }, mocs);
            Assert.Equal(
                new decimal[] { 0.0190m, 0.0170m, 0.0150m, 0.0130m, 0.0100m, 0.0080m, 0.0050m, 0.0030m, 0.0020m, 0.0010m },
                tyle);

            Assert.Equal(10, mocs.Length);
            Assert.Equal(10, tyle.Length);
        }

        [Fact]
        public void Test_PhiThamDinhThietKe_Khop_File33()
        {
            // File 33: mốc quy mô theo G_XD (chưa VAT, tỷ đồng): <=15, 50, 100, 200, 500, 1000, 2000, 5000, >=8000
            Assert.Equal(
                new decimal[] { 15m, 50m, 100m, 200m, 500m, 1000m, 2000m, 5000m, 8000m },
                DinhMucChiPhiKhacBTC.MocQuyMoThamDinhTKDT);

            var danDung = DinhMucChiPhiKhacBTC.TiLeThamDinhThietKe["Dân dụng"];
            var NN_PTNT = DinhMucChiPhiKhacBTC.TiLeThamDinhThietKe["Nông nghiệp & PTNT"];

            Assert.Equal(new decimal[] { 0.165m, 0.110m, 0.085m, 0.065m, 0.050m, 0.041m, 0.029m, 0.022m, 0.019m }, danDung);
            Assert.Equal(new decimal[] { 0.121m, 0.080m, 0.061m, 0.048m, 0.037m, 0.028m, 0.023m, 0.017m, 0.014m }, NN_PTNT);
            Assert.Equal(9, danDung.Length);
            Assert.Equal(9, NN_PTNT.Length);
        }

        [Fact]
        public void Test_PhiThamDinhDuToan_Khop_File34()
        {
            var danDung = DinhMucChiPhiKhacBTC.TiLeThamDinhDuToan["Dân dụng"];
            var NN_PTNT = DinhMucChiPhiKhacBTC.TiLeThamDinhDuToan["Nông nghiệp & PTNT"];

            Assert.Equal(new decimal[] { 0.160m, 0.106m, 0.083m, 0.062m, 0.046m, 0.038m, 0.028m, 0.021m, 0.018m }, danDung);
            Assert.Equal(new decimal[] { 0.117m, 0.076m, 0.060m, 0.046m, 0.035m, 0.026m, 0.022m, 0.016m, 0.014m }, NN_PTNT);
            Assert.Equal(9, danDung.Length);
            Assert.Equal(9, NN_PTNT.Length);
        }

        [Fact]
        public void Test_ThamDinhThietKe_ThuyLoi_CoTyLe_Khop_Files()
        {
            // Thủy lợi (thuộc NN & PTNT) phải tra đúng cột "Nông nghiệp & PTNT" trong file 33/34
            var model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: "Thủy lợi",
                capCT: "Cấp III",
                soBuocTK: 2,
                chiPhiXD: 10_000_000_000m,
                chiPhiTB: 0m,
                chiPhiBT: 0m);

            var kTD_TK = model.Items.First(x => x.MaChiPhi == "K_TD_TK");
            // File 33, NN&PTNT tại mốc <=15 tỷ (G_XD=10 tỷ): 0.121%
            Assert.Equal(0.121m, kTD_TK.TyLePhanTram);
            Assert.True(kTD_TK.GiaTriTruocThue > 0);

            var kTD_DT = model.Items.First(x => x.MaChiPhi == "K_TD_DT");
            // File 34, NN&PTNT tại mốc <=15 tỷ: 0.117%
            Assert.Equal(0.117m, kTD_DT.TyLePhanTram);
            Assert.True(kTD_DT.GiaTriTruocThue > 0);
        }
    }
}