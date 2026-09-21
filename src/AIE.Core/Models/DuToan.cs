namespace AIE.Core.Models;

/// <summary>
/// Dự toán xây dựng công trình — tổng hợp tất cả khoản mục chi phí.
/// G_XDCT = GXD + GTB + GQLDA + GTV + GK + GDP
/// </summary>
public class DuToan
{
    /// <summary>Tên dự án</summary>
    public string TenDuAn { get; set; } = string.Empty;

    /// <summary>Tên công trình</summary>
    public string TenCongTrinh { get; set; } = string.Empty;

    /// <summary>Loại công trình (Dân dụng, Giao thông,...)</summary>
    public string LoaiCongTrinh { get; set; } = string.Empty;

    /// <summary>Cấp công trình (nếu có)</summary>
    public string? CapCongTrinh { get; set; }

    /// <summary>Số bước thiết kế (0 = Chưa chọn, 1 bước, 2 bước, 3 bước)</summary>
    public int SoBuocThietKe { get; set; } = 0;

    /// <summary>Chủ đầu tư</summary>
    public string? ChuDauTu { get; set; }

    /// <summary>Địa điểm</summary>
    public string DiaDiem { get; set; } = string.Empty;

    /// <summary>Vùng áp dụng (để tính lương, máy)</summary>
    public AIE.Core.Enums.Vung VungApDung { get; set; } = AIE.Core.Enums.Vung.VungII;

    /// <summary>ID của Bộ Đơn Giá được sử dụng cho dự toán này (để lấy giá VL, NC, Máy)</summary>
    public int? BoDonGiaId { get; set; }

    /// <summary>Vùng áp dụng cho dự toán (2, 3, 4, hoặc 5 - Cù Lao Chàm)</summary>
    public AIE.Core.Enums.Vung Vung { get; set; } = AIE.Core.Enums.Vung.VungII;

    /// <summary>Danh sách hạng mục</summary>
    public List<HangMuc> DanhSachHangMuc { get; set; } = [];

    /// <summary>Bảng tổng hợp vật tư (Vật liệu, Nhân công, Máy thi công) phân tích từ danh sách hạng mục</summary>
    public BangTongHopVatTu BangTongHop { get; set; } = new();

    /// <summary>Chi phí xây dựng (GXD)</summary>
    public ChiPhiXayDung? ChiPhiXD { get; set; }

    /// <summary>Chi phí thiết bị (GTB)</summary>
    public decimal ChiPhiThietBi { get; set; }

    /// <summary>Chi phí quản lý dự án (GQLDA)</summary>
    public decimal ChiPhiQLDA { get; set; }
    public decimal TiLeQLDA { get; set; }

    /// <summary>Chi phí tư vấn (GTV) — tổng các loại tư vấn</summary>
    public decimal ChiPhiTuVan { get; set; }
    public List<ChiPhiTuVanChiTiet> ChiTietTuVan { get; set; } = [];

    /// <summary>Chi phí khác (GK)</summary>
    public decimal ChiPhiKhac { get; set; }

    /// <summary>Chi phí dự phòng (GDP)</summary>
    public decimal ChiPhiDuPhong { get; set; }
    public decimal DuPhongKhoiLuong { get; set; }
    public decimal DuPhongTruotGia { get; set; }

    /// <summary>Bảng tổng hợp kinh phí & Tổng mức đầu tư chi tiết (TT 36/2026 & TT 38/2026)</summary>
    public BangTongHopKinhPhiModel? BangKinhPhi { get; set; }

    /// <summary>G_XDCT = GXD + GTB + GQLDA + GTV + GK + GDP</summary>
    public decimal TongDuToan =>
        (ChiPhiXD?.GXD ?? 0) + ChiPhiThietBi + ChiPhiQLDA +
        ChiPhiTuVan + ChiPhiKhac + ChiPhiDuPhong;
}

/// <summary>
/// Hạng mục trong dự toán (Level 1 — Đơn vị xác định Loại công trình và định mức TT 36).
/// </summary>
public class HangMuc
{
    public int STT { get; set; }
    public int RowIndex { get; set; }
    public string MaHangMuc { get; set; } = string.Empty;
    public string TenHangMuc { get; set; } = string.Empty;
    public string LoaiCongTrinh { get; set; } = string.Empty;
    public string PhanLoaiPhu { get; set; } = string.Empty;
    public string? CapCongTrinh { get; set; }

    /// <summary>Danh sách các hạng mục con (Level 2) nếu có</summary>
    public List<HangMucCon> DanhSachHangMucCon { get; set; } = [];

    /// <summary>Danh sách toàn bộ công tác thuộc hạng mục này</summary>
    public List<DongDuToan> DanhSachCongTac { get; set; } = [];

    /// <summary>Chi phí xây dựng tính riêng theo Bảng 3.8 TT 36 cho Hạng mục này</summary>
    public ChiPhiXayDung? ChiPhiXD { get; set; }

    /// <summary>Hệ số điều chỉnh riêng của hạng mục theo Vật liệu (nhập tay trên sheet HeSo_DieuChinh)</summary>
    public decimal HeSoVL { get; set; } = 1.0m;

    /// <summary>Hệ số điều chỉnh riêng của hạng mục theo Nhân công</summary>
    public decimal HeSoNC { get; set; } = 1.0m;

    /// <summary>Hệ số điều chỉnh riêng của hạng mục theo Máy thi công</summary>
    public decimal HeSoM { get; set; } = 1.0m;

    /// <summary>Tổng thành tiền Vật liệu của hạng mục</summary>
    public decimal TongVL => DanhSachCongTac.Sum(x => x.ThanhTienVL);

    /// <summary>Tổng thành tiền Nhân công của hạng mục</summary>
    public decimal TongNC => DanhSachCongTac.Sum(x => x.ThanhTienNC);

    /// <summary>Tổng thành tiền Máy thi công của hạng mục</summary>
    public decimal TongMay => DanhSachCongTac.Sum(x => x.ThanhTienMay);

    /// <summary>Tổng chi phí trực tiếp của hạng mục T = VL + NC + M</summary>
    public decimal TongChiPhi => TongVL + TongNC + TongMay;
}

/// <summary>
/// 1 dòng dự toán = 1 công tác XD với Khối lượng cụ thể.
/// </summary>
public class DongDuToan
{
    public int STT { get; set; }

    /// <summary>Tên hạng mục con (nếu thuộc hạng mục con)</summary>
    public string? TenHangMucCon { get; set; }

    /// <summary>Mã hiệu định mức</summary>
    public string MaHieu { get; set; } = string.Empty;

    /// <summary>Tên công tác</summary>
    public string TenCongTac { get; set; } = string.Empty;

    /// <summary>Đơn vị</summary>
    public string DonVi { get; set; } = string.Empty;
    
    /// <summary>Danh sách hao phí (vật liệu, nhân công, máy)</summary>
    public List<HaoPhi> DanhSachHaoPhi { get; set; } = [];

    /// <summary>Khối lượng</summary>
    public decimal KhoiLuong { get; set; }

    /// <summary>Đơn giá vật liệu</summary>
    public decimal DonGiaVL { get; set; }

    /// <summary>Đơn giá nhân công</summary>
    public decimal DonGiaNC { get; set; }

    /// <summary>Đơn giá máy</summary>
    public decimal DonGiaMay { get; set; }

    /// <summary>Đơn giá tổng hợp = VL + NC + M</summary>
    public decimal DonGiaTongHop => DonGiaVL + DonGiaNC + DonGiaMay;

    /// <summary>Thành tiền VL = KL × Đơn giá VL (làm tròn 0 chữ số thập phân theo RoundingRules)</summary>
    public decimal ThanhTienVL => Math.Round(KhoiLuong * DonGiaVL, 0);

    /// <summary>Thành tiền NC = KL × Đơn giá NC (làm tròn 0 chữ số thập phân theo RoundingRules)</summary>
    public decimal ThanhTienNC => Math.Round(KhoiLuong * DonGiaNC, 0);

    /// <summary>Thành tiền Máy = KL × Đơn giá Máy (làm tròn 0 chữ số thập phân theo RoundingRules)</summary>
    public decimal ThanhTienMay => Math.Round(KhoiLuong * DonGiaMay, 0);

    /// <summary>Thành tiền = Thành tiền VL + NC + Máy</summary>
    public decimal ThanhTien => ThanhTienVL + ThanhTienNC + ThanhTienMay;
}

/// <summary>
/// Chi tiết 1 loại tư vấn trong dự toán.
/// </summary>
public class ChiPhiTuVanChiTiet
{
    public string LoaiTuVan { get; set; } = string.Empty;
    public decimal TiLe { get; set; }
    public decimal GiaTri { get; set; }
}
