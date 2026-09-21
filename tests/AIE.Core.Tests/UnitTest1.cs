using Xunit;
using AIE.Core.Models;

namespace AIE.Core.Tests;

public class DuToanInitializationTests
{
    [Fact]
    public void DuToan_KhoiTao_DiaDiemKhongBiHardcodeDaNang()
    {
        var dt = new DuToan();
        Assert.True(string.IsNullOrEmpty(dt.DiaDiem), "DiaDiem mặc định phải là rỗng, không được hardcode Đà Nẵng.");
    }

    [Fact]
    public void HangMuc_KhoiTao_HeSoDieuChinhMacDinhLa1()
    {
        var hm = new HangMuc();
        Assert.Equal(1.0m, hm.HeSoVL);
        Assert.Equal(1.0m, hm.HeSoNC);
        Assert.Equal(1.0m, hm.HeSoM);
    }

    [Fact]
    public void ChiPhiKinhPhiItem_ApDungMinValue_KhiGiaTriNhoHonToiThieu()
    {
        var model = new BangTongHopKinhPhiModel();
        var item = new ChiPhiKinhPhiItem
        {
            Id = "TV_TT_TK",
            TenChiPhi = "Chi phí thẩm tra thiết kế xây dựng",
            Nhom = NhomChiPhi.TuVanDauTuXD,
            CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
            CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
            TyLePhanTram = 0.197m,
            HeSoDieuChinh = 1.0m,
            MinValue = 2000000m,
            ThueSuatGTGT = 0.10m,
            IsActive = true
        };
        model.Items.Add(item);
        model.ChiPhiXDTruocThue = 744174705m;
        model.ChiPhiNhaTamTruocThue = 8185922m;
        model.ChiPhiTBTruocThue = 0m;
        model.TinhToanLai();

        Assert.Equal(2000000m, item.GiaTriTruocThue);
        Assert.Equal(2200000m, item.GiaTriSauThue);
    }
}
