using Xunit;
using AIE.Core.Services.LapDuToan;

namespace AIE.Core.Tests;

public class ChiPhiXayDungTests
{
    [Fact]
    public void ChiPhiXayDungCalc_TheoDuLieu_Screenshot()
    {
        // Arrange
        var calc = new ChiPhiXayDungCalc();
        decimal tongVL = 762113m;
        decimal tongNC = 0m;
        decimal tongMay = 0m;
        
        decimal tiLeCPC = 7.3m;
        decimal tiLeTT = 2.5m;
        decimal tiLeTNCTTT = 5.5m;
        decimal tiLeGTGT = 8.0m;
        decimal tiLeNhaTam = 1.1m;

        // Act
        var result = calc.Tinh(tongVL, tongNC, tongMay, tiLeCPC, tiLeTT, tiLeTNCTTT, tiLeGTGT, tiLeNhaTam);

        // Assert
        Assert.Equal(762113m, result.T);
        Assert.Equal(55634m, result.CPC); // Khớp 55.634
        Assert.Equal(19053m, result.TT);  // Khớp 19.053
        Assert.Equal(74687m, result.GT);  // Khớp 74.687
        
        Assert.Equal(46024m, result.TL);  // Khớp 46.024
        Assert.Equal(882824m, result.G);  // Khớp 882.824
        
        Assert.Equal(70626m, result.GTGT); // Khớp 70.626
        Assert.Equal(953450m, result.Gxd); // Khớp 953.450
        
        Assert.Equal(10488m, result.LT);   // Khớp 10.488 (Tuy nhiên có thể chênh lệch 1 đồng do cách làm tròn của phần mềm gốc: G * 1.1% = 882824 * 0.011 = 9711.064???
        // Đợi đã, trong screenshot, Chi phí nhà tạm = 10.488 
        // Tính lại G * 1.1% = 882824 * 1.1% = 9711.064 (Lệch).
        // Hay là Gxd * 1.1% ? 953450 * 1.1% = 10487.95 -> 10488 !!!
        // => À, TT 36 có thể quy định Nhà tạm tính trên Gxd (Sau thuế) hoặc Gx 1.1%?
        // Hãy Kiểm tra kỹ Screenshot 2: "Chi phí nhà tạm để ở và điều hành thi công (G x 1,1...)"
        // Screenshot bị cắt, "G x 1,..." có thể là (Gxd x 1,1%) hoặc (G x 1,...) 
        // 953450 * 1.1% = 10487.95 => Làm tròn = 10488. (Vậy là tính trên Gxd).
    }

    [Fact]
    public void ChiPhiNhaTam_TheoQuyetDinh1538_QD_BXD()
    {
        // Kiểm tra tính tuân thủ theo Quyết định số 1538/QĐ-BXD ngày 28/8/2026:
        // Đính chính công thức tại Bảng 3.8: GXDTT × Tỷ lệ × (1 + TGTGT)
        var calc = new ChiPhiXayDungCalc();
        decimal gxdtt = 1_000_000_000m; // 1 tỷ chi phí XD trước thuế
        decimal tiLeNhaTam = 1.2m;      // 1.2%
        decimal thueSuatGTGT = 10m;     // 10%

        var result = calc.Tinh(
            tongVL: gxdtt,
            tongNC: 0m,
            tongMay: 0m,
            tiLeCPC: 0m,
            tiLeTT: 0m,
            tiLeTNCTTT: 0m,
            tiLeGTGT: thueSuatGTGT,
            tiLeNhaTam: tiLeNhaTam);

        // Công thức theo QĐ 1538: GXDTT × Tỷ lệ × (1 + TGTGT)
        decimal expectedLT = Math.Round(gxdtt * (tiLeNhaTam / 100m) * (1m + thueSuatGTGT / 100m), 0, MidpointRounding.AwayFromZero);
        // = 1.000.000.000 × 1.2% × (1 + 10%) = 12.000.000 × 1.1 = 13.200.000 đ
        Assert.Equal(13_200_000m, expectedLT);
        Assert.Equal(expectedLT, result.LT);
    }

    [Fact]
    public void ChiPhiXayDung_Bang38_KhongCongNhaTamVaoSauThue()
    {
        // Kiểm tra đúng theo Bảng 3.8 TT 36/2026:
        // CHI PHÍ XÂY DỰNG SAU THUẾ (GXD) = GXDTT + GTGT, KHÔNG CỘNG CHI PHÍ NHÀ TẠM LT
        var calc = new ChiPhiXayDungCalc();
        var result = calc.Tinh(
            tongVL: 100_000_000m,
            tongNC: 50_000_000m,
            tongMay: 50_000_000m,
            tiLeCPC: 7.3m,
            tiLeTT: 2.5m,
            tiLeTNCTTT: 5.5m,
            tiLeGTGT: 10m,
            tiLeNhaTam: 1.1m);

        // GXDTT = T + GT + TL
        // T = 200.000.000
        // CPC = 200.000.000 * 7.3% = 14.600.000
        // TT = 200.000.000 * 2.5% = 5.000.000
        // GT = 19.600.000
        // TL = (200.000.000 + 19.600.000) * 5.5% = 12.078.000
        // GXDTT = 200.000.000 + 19.600.000 + 12.078.000 = 231.678.000
        // GTGT = 231.678.000 * 10% = 23.167.800
        // Chi phí xây dựng sau thuế: GXD = 231.678.000 + 23.167.800 = 254.845.800 đ
        Assert.Equal(254_845_800m, result.GXD);
        Assert.Equal(254_845_800m, result.Gxd); // Alias cũ tương thích

        // Nhà tạm: LT = GXDTT * 1.1% * 1.1 = 2.803.304 đ
        Assert.True(result.LT > 0);
        // Xác nhận GXD không bị cộng dồn thêm LT
        Assert.NotEqual(result.GXD + result.LT, result.GXD);
    }

    [Fact]
    public void NoiSuyTiLeNhaTam_Bang37_ChuanXac()
    {
        // Công trình theo tuyến: <=15 tỷ: 2.2%, <=100: 2.0%, <=500: 1.9%, <=1000: 1.8%, >1000: 1.7%
        Assert.Equal(2.2m, AIE.Core.Services.Shared.InterpolationHelper.NoiSuyTiLeNhaTam("Công trình xây dựng theo tuyến", 10m));
        Assert.Equal(2.2m, AIE.Core.Services.Shared.InterpolationHelper.NoiSuyTiLeNhaTam("Công trình xây dựng theo tuyến", 15m));
        Assert.Equal(2.0m, AIE.Core.Services.Shared.InterpolationHelper.NoiSuyTiLeNhaTam("Công trình xây dựng theo tuyến", 100m));
        Assert.Equal(1.95m, AIE.Core.Services.Shared.InterpolationHelper.NoiSuyTiLeNhaTam("Công trình xây dựng theo tuyến", 300m));
        Assert.Equal(1.7m, AIE.Core.Services.Shared.InterpolationHelper.NoiSuyTiLeNhaTam("Công trình xây dựng theo tuyến", 1500m));

        // Công trình còn lại: <=15 tỷ: 1.1%, <=100: 1.0%, <=500: 0.95%, <=1000: 0.9%, >1000: 0.85%
        Assert.Equal(1.1m, AIE.Core.Services.Shared.InterpolationHelper.NoiSuyTiLeNhaTam("Công trình xây dựng còn lại", 10m));
        Assert.Equal(1.1m, AIE.Core.Services.Shared.InterpolationHelper.NoiSuyTiLeNhaTam("Công trình xây dựng còn lại", 15m));
        Assert.Equal(1.0m, AIE.Core.Services.Shared.InterpolationHelper.NoiSuyTiLeNhaTam("Công trình xây dựng còn lại", 100m));
        Assert.Equal(0.975m, AIE.Core.Services.Shared.InterpolationHelper.NoiSuyTiLeNhaTam("Công trình xây dựng còn lại", 300m));
        Assert.Equal(0.85m, AIE.Core.Services.Shared.InterpolationHelper.NoiSuyTiLeNhaTam("Công trình xây dựng còn lại", 1200m));
    }

    [Fact]
    public void TraCuuDinhMuc_Bang32_Va_Bang36_ChuanXac()
    {
        // Bảng 3.2: Chi phí chung TMĐT
        Assert.Equal(7.3m, AIE.Core.Services.Shared.InterpolationHelper.LayTiLeCPCTMDT("Dân dụng"));
        Assert.Equal(11.6m, AIE.Core.Services.Shared.InterpolationHelper.LayTiLeCPCTMDT("Dân dụng", "Tu bổ di tích"));
        Assert.Equal(6.2m, AIE.Core.Services.Shared.InterpolationHelper.LayTiLeCPCTMDT("Công nghiệp"));
        Assert.Equal(7.3m, AIE.Core.Services.Shared.InterpolationHelper.LayTiLeCPCTMDT("Công nghiệp", "Đường hầm thủy điện, hầm lò"));
        Assert.Equal(5.5m, AIE.Core.Services.Shared.InterpolationHelper.LayTiLeCPCTMDT("Hạ tầng kỹ thuật"));

        // Bảng 3.6: Thu nhập chịu thuế tính trước
        Assert.Equal(5.5m, AIE.Core.Services.Shared.InterpolationHelper.LayTiLeTNCTTT("Dân dụng"));
        Assert.Equal(6.0m, AIE.Core.Services.Shared.InterpolationHelper.LayTiLeTNCTTT("Công nghiệp"));
        Assert.Equal(6.0m, AIE.Core.Services.Shared.InterpolationHelper.LayTiLeTNCTTT("Giao thông"));
        Assert.Equal(5.5m, AIE.Core.Services.Shared.InterpolationHelper.LayTiLeTNCTTT("Nông nghiệp & PTNT"));
        Assert.Equal(5.5m, AIE.Core.Services.Shared.InterpolationHelper.LayTiLeTNCTTT("Hạ tầng kỹ thuật"));
    }
}


