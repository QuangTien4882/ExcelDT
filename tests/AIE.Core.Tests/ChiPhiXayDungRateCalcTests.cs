using System.Linq;
using Xunit;
using AIE.Core.Models;
using AIE.Core.Services;

namespace AIE.Core.Tests
{
    /// <summary>
    /// Kiểm tra ChiPhiXayDungRateCalc — tính tỷ lệ % chi phí xây dựng thuần túy
    /// (CPC, TT, TNCTTT, GTGT, Nhà tạm) + cảnh báo khi thiếu dữ liệu.
    /// </summary>
    public class ChiPhiXayDungRateCalcTests
    {
        [Fact]
        public void TinhTyLe_KhongCoCSDL_DungBangChuanVaBaoCanhBao()
        {
            var r = ChiPhiXayDungRateCalc.TinhTyLe(
                loaiCongTrinh: "Dân dụng",
                phanLoai: null,
                gxdtt: 10_000_000_000m,
                tongT: 0m,
                soBuocThietKe: 2,
                laVungSauXa: false,
                loaiCongTrinhNhaTam: "Công trình xây dựng còn lại",
                tiLeCPC: 0m, tiLeTT: 0m, tiLeTNCTTT: 0m, tiLeGTGT: 0m, tiLeNhaTam: 0m,
                cpcTheoQuyMo: null, ttTheoLoai: null);

            Assert.Equal(7.3m, r.TiLeCPC);        // Bảng 3.2 Dân dụng
            Assert.Equal(2.5m, r.TiLeTT);         // mặc định Bảng 3.5
            Assert.Equal(5.5m, r.TiLeTNCTTT);     // Bảng 3.6 Dân dụng
            Assert.Equal(10.0m, r.TiLeGTGT);      // thuế GTGT mặc định
            Assert.Equal(1.1m, r.TiLeNhaTam);     // Bảng 3.7 quy mô ≤ 15 tỷ, CT còn lại
            Assert.True(r.CanhBao.Count >= 2, "Phải cảnh báo khi thiếu Bảng 3.3 và Bảng 3.5.");
        }

        [Fact]
        public void TinhTyLe_GiuNguyenTyLeDaNhapKhongCanhBao()
        {
            var r = ChiPhiXayDungRateCalc.TinhTyLe(
                loaiCongTrinh: "Giao thông",
                phanLoai: null,
                gxdtt: 15_000_000_000m,
                tongT: 0m,
                soBuocThietKe: 2,
                laVungSauXa: false,
                loaiCongTrinhNhaTam: "Công trình xây dựng còn lại",
                tiLeCPC: 6.2m, tiLeTT: 3.0m, tiLeTNCTTT: 6.0m, tiLeGTGT: 10.0m, tiLeNhaTam: 1.1m);

            Assert.Equal(6.2m, r.TiLeCPC);
            Assert.Equal(3.0m, r.TiLeTT);
            Assert.Equal(6.0m, r.TiLeTNCTTT);
            Assert.Equal(10.0m, r.TiLeGTGT);
            Assert.Equal(1.1m, r.TiLeNhaTam);
            Assert.True(r.CanhBao.Count == 0);
        }

        [Fact]
        public void TinhTyLe_NoiSuyCPC_TheoQuyMoKhiCoCSDL()
        {
            // Bảng 3.3: mốc ≤40 tỷ → 7.0%; ≤100 tỷ → 6.0%; quy mô 60 tỷ: 7.0 - (1.0/60)x20 = 6.667
            var dsCPC = new[]
            {
                new DinhMucCPC { QuyMoMin = 0m, QuyMoMax = 40m, TiLe = 7.0m },
                new DinhMucCPC { QuyMoMin = 40m, QuyMoMax = 100m, TiLe = 6.0m },
                new DinhMucCPC { QuyMoMin = 100m, QuyMoMax = 1000m, TiLe = 5.0m }
            };

            var r = ChiPhiXayDungRateCalc.TinhTyLe(
                loaiCongTrinh: "Dân dụng",
                phanLoai: null,
                gxdtt: 60_000_000_000m,
                tongT: 0m,
                soBuocThietKe: 2,
                laVungSauXa: false,
                loaiCongTrinhNhaTam: "Công trình xây dựng còn lại",
                tiLeCPC: 0m, tiLeTT: 0m, tiLeTNCTTT: 0m, tiLeGTGT: 0m, tiLeNhaTam: 0m,
                cpcTheoQuyMo: dsCPC, ttTheoLoai: null);

            Assert.Equal(6.667m, r.TiLeCPC);
            Assert.True(r.CanhBao.All(w => !w.Contains("Chi phí chung")), "Có CSDL CPC thì không được cảnh báo CPC.");
        }

        [Fact]
        public void TinhTyLe_VungSauXa_NhanHeSo11()
        {
            var r = ChiPhiXayDungRateCalc.TinhTyLe(
                loaiCongTrinh: "Dân dụng",
                phanLoai: null,
                gxdtt: 10_000_000_000m,
                tongT: 0m,
                soBuocThietKe: 2,
                laVungSauXa: true,
                loaiCongTrinhNhaTam: "Công trình xây dựng còn lại",
                tiLeCPC: 7.3m, tiLeTT: 2.5m, tiLeTNCTTT: 5.5m, tiLeGTGT: 10.0m, tiLeNhaTam: 1.1m);

            Assert.Equal(8.03m, r.TiLeCPC); // 7.3 × 1.1 = 8.03
        }

        [Fact]
        public void TinhQuyMo_UocLuongTuTongChiPhiTrucTiep()
        {
            // T = 10 tỷ, với CPC 7.3% + TT 2.5% + TNCTTT 5.5% ⇒ GXDTT ≈ 10 × 1.098 × 1.055 ≈ 11.5839
            decimal quyMo = ChiPhiXayDungRateCalc.TinhQuyMoTuChiPhiXD(
                gxdtt: 0m, tongT: 10_000_000_000m, tiLeCPC: 7.3m, tiLeTT: 2.5m, tiLeTNCTTT: 5.5m);
            Assert.Equal(11.5839m, quyMo);

            // Quy mô 20 tỷ trước thuế truyền thẳng vào
            Assert.Equal(20m, ChiPhiXayDungRateCalc.TinhQuyMoTuChiPhiXD(
                gxdtt: 20_000_000_000m, tongT: 0m, tiLeCPC: 0m, tiLeTT: 0m, tiLeTNCTTT: 0m));
        }

        [Fact]
        public void TinhTyLe_QuyMoKhongXacDinh_BaoCanhBao()
        {
            var r = ChiPhiXayDungRateCalc.TinhTyLe(
                loaiCongTrinh: "Dân dụng",
                phanLoai: null,
                gxdtt: 0m, tongT: 0m,
                soBuocThietKe: 2,
                laVungSauXa: false,
                loaiCongTrinhNhaTam: "Công trình xây dựng còn lại",
                tiLeCPC: 0m, tiLeTT: 0m, tiLeTNCTTT: 0m, tiLeGTGT: 0m, tiLeNhaTam: 0m);

            Assert.True(r.CanhBao.Any(w => w.Contains("quy mô")), "Phải cảnh báo khi không xác định được quy mô.");
        }
    }
}