using System.Collections.Generic;
using Xunit;
using AIE.Core.Models;
using AIE.Core.Services.LapDuToan;

namespace AIE.Core.Tests
{
    public class DuToanSyncHelperTests
    {
        [Fact]
        public void Merge_KhiHangMucBiXoa_GhepChinhXacTheoTenHangMucKhongBiLechIndex()
        {
            // Dự toán cũ có 2 hạng mục
            var duToanCu = new DuToan
            {
                TenCongTrinh = "Công trình Test",
                DanhSachHangMuc = new List<HangMuc>
                {
                    new HangMuc
                    {
                        TenHangMuc = "Hạng mục 1: Phần Móng",
                        MaHangMuc = "HM01",
                        HeSoVL = 1.1m,
                        HeSoNC = 1.2m,
                        HeSoM = 1.3m,
                        DanhSachCongTac = new List<DongDuToan>
                        {
                            new DongDuToan { MaHieu = "AF.11111", TenCongTac = "Bê tông móng", KhoiLuong = 10 }
                        }
                    },
                    new HangMuc
                    {
                        TenHangMuc = "Hạng mục 2: Phần Thân",
                        MaHangMuc = "HM02",
                        HeSoVL = 2.1m,
                        HeSoNC = 2.2m,
                        HeSoM = 2.3m,
                        DanhSachCongTac = new List<DongDuToan>
                        {
                            new DongDuToan { MaHieu = "AF.22222", TenCongTac = "Bê tông cột", KhoiLuong = 20 }
                        }
                    }
                }
            };

            // Người dùng trên Excel xóa Hạng mục 1, chỉ còn lại Hạng mục 2
            var latestBoq = new DuToan
            {
                TenCongTrinh = "Công trình Test",
                DanhSachHangMuc = new List<HangMuc>
                {
                    new HangMuc
                    {
                        TenHangMuc = "Hạng mục 2: Phần Thân",
                        MaHangMuc = "HM02",
                        DanhSachCongTac = new List<DongDuToan>
                        {
                            new DongDuToan { MaHieu = "AF.22222", TenCongTac = "Bê tông cột", KhoiLuong = 25 }
                        }
                    }
                }
            };

            // Thực hiện Merge
            var merged = DuToanSyncHelper.Merge(latestBoq, duToanCu);

            // Kiểm tra:
            // 1. Chỉ còn 1 hạng mục
            Assert.Single(merged.DanhSachHangMuc);
            var hmMerged = merged.DanhSachHangMuc[0];
            Assert.Equal("Hạng mục 2: Phần Thân", hmMerged.TenHangMuc);

            // 2. Hệ số phải là của Hạng mục 2 (2.1, 2.2, 2.3), KHÔNG được bị gán nhầm thành của Hạng mục 1 (1.1, 1.2, 1.3)
            Assert.Equal(2.1m, hmMerged.HeSoVL);
            Assert.Equal(2.2m, hmMerged.HeSoNC);
            Assert.Equal(2.3m, hmMerged.HeSoM);

            // 3. Khối lượng được cập nhật mới (25)
            Assert.Equal(25m, hmMerged.DanhSachCongTac[0].KhoiLuong);
        }
    }
}
