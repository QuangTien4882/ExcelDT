using System;
using Xunit;
using AIE.Core.Models;

namespace AIE.Core.Tests
{
    /// <summary>
    /// Kiểm tra hàm nội suy định mức tỷ lệ % theo quy mô (DinhMucChiPhiKhacBTC.NoiSuy):
    /// - Nội suy tuyến tính giữa các mốc
    /// - Giá trị clamp ở ranh giới dưới/trên
    /// - Không làm tròn bên trong (trả giá trị chính xác)
    /// </summary>
    public class DinhMucNoiSuyTests
    {
        // Bảng 1.1 QLDA — Dân dụng (khớp 100% văn bản TT 38/2026)
        private static readonly decimal[] Mocs = { 10m, 20m, 50m, 100m };
        private static readonly decimal[] TiLes = { 3.446m, 2.923m, 2.610m, 2.017m };

        [Fact]
        public void NoiSuy_TrongKhoang_NoiSuyTuyenTinhChinhXac()
        {
            // 15 tỷ: 3.446 - (0.523/10)*5 = 3.1845 (chính xác, không làm tròn)
            Assert.Equal(3.1845m, DinhMucChiPhiKhacBTC.NoiSuy(Mocs, TiLes, 15m));
        }

        [Fact]
        public void NoiSuy_TaiMoc_ChinhXacBangGiaTriMoc()
        {
            Assert.Equal(3.446m, DinhMucChiPhiKhacBTC.NoiSuy(Mocs, TiLes, 10m));
            Assert.Equal(2.923m, DinhMucChiPhiKhacBTC.NoiSuy(Mocs, TiLes, 20m));
        }

        [Fact]
        public void NoiSuy_DuoiMocNhoNhat_ClampLayMocDau()
        {
            Assert.Equal(3.446m, DinhMucChiPhiKhacBTC.NoiSuy(Mocs, TiLes, 1m));
            Assert.Equal(3.446m, DinhMucChiPhiKhacBTC.NoiSuy(Mocs, TiLes, 0m));
        }

        [Fact]
        public void NoiSuy_TrenMocLonNhat_ClampLayMocCuoi()
        {
            Assert.Equal(2.017m, DinhMucChiPhiKhacBTC.NoiSuy(Mocs, TiLes, 500m));
            Assert.Equal(2.017m, DinhMucChiPhiKhacBTC.NoiSuy(Mocs, TiLes, 100m));
        }

        [Fact]
        public void NoiSuy_ChietKhauDonGian_DungCongThucQuyChuan()
        {
            // Nt = Nb - ((Nb - Na)/(Ga - Gb)) * (Gt - Gb)  tại Gt = 30 (giữa 20 và 50)
            // = 2.923 - ((2.923 - 2.610)/(50-20)) * (30-20) = 2.923 - (0.313/30)*10 = 2.923 - 0.1043333...
            decimal expected = 2.923m - (2.923m - 2.610m) / (50m - 20m) * (30m - 20m);
            Assert.Equal(expected, DinhMucChiPhiKhacBTC.NoiSuy(Mocs, TiLes, 30m));
        }
    }
}