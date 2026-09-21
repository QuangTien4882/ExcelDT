using AIE.Core.Enums;
using System.Collections.Generic;

namespace AIE.Core.Models;

/// <summary>
/// Cấu hình mapping cột cho sheet Đơn Giá Chi Tiết
/// </summary>
public class ThamDinhConfig
{
    public string SheetName { get; set; } = string.Empty;
    public string ColMaHieu { get; set; } = "B";
    public string ColTenCT { get; set; } = "E";
    public string ColDonVi { get; set; } = "F";
    public string ColDinhMuc { get; set; } = "G";
    public string ColDonGia { get; set; } = "H";
    public string ColThanhTien { get; set; } = "I";
    public int DongBatDau { get; set; } = 6;
    public Vung Vung { get; set; } = Vung.VungII;
}

/// <summary>
/// 1 công tác đọc được từ file dự toán cần thẩm định
/// </summary>
public class CongTacThamDinh
{
    public int SoDongExcel { get; set; }
    public string MaHieu { get; set; } = string.Empty;
    public string TenCongTac { get; set; } = string.Empty;
    public string DonVi { get; set; } = string.Empty;
    public decimal KhoiLuong { get; set; }
    public decimal DonGia { get; set; }
    public decimal ThanhTien { get; set; }
    public List<HaoPhiThamDinh> DanhSachHaoPhi { get; set; } = new();
}

/// <summary>
/// 1 hao phí đọc được từ file dự toán
/// </summary>
public class HaoPhiThamDinh
{
    public int SoDongExcel { get; set; }
    public LoaiHaoPhi Loai { get; set; }
    public string MaHieuHP { get; set; } = string.Empty;
    public string TenHaoPhi { get; set; } = string.Empty;
    public string DonVi { get; set; } = string.Empty;
    public decimal DinhMuc { get; set; }
    public decimal DonGia { get; set; }
    public decimal ThanhTien { get; set; } // Nếu cần
}

/// <summary>
/// Kết quả thẩm định 1 công tác
/// </summary>
public class KetQuaCongTacThamDinh
{
    public CongTacThamDinh DuToan { get; set; } = new();
    public CongTacXayDung? DinhMucChuan { get; set; }
    public List<SaiLechDinhMuc> DanhSachSaiLech { get; set; } = new();
    public bool DaDat => DanhSachSaiLech.Count == 0 && DinhMucChuan != null;
}

/// <summary>
/// Chi tiết 1 sai lệch (trên 1 dòng Excel)
/// </summary>
public class SaiLechDinhMuc
{
    public string LoaiLoi { get; set; } = string.Empty;
    public string MoTa { get; set; } = string.Empty;
    public LoaiHaoPhi? LoaiHP { get; set; }
    public int? SoDongExcel { get; set; }
    
    // Dữ liệu dùng để xuất cột
    public HaoPhi? HaoPhiChuan { get; set; }
    public decimal? DonGiaChuan { get; set; } // Lấy từ Repo
    public string GhiChuDonGia { get; set; } = string.Empty;
    
    // Thuộc tính tiện ích
    public decimal ChenhLechDinhMuc => (HaoPhiChuan != null && HaoPhiDuToan != null) ? HaoPhiDuToan.DinhMuc - HaoPhiChuan.DinhMuc : 0;
    public decimal ChenhLechDonGia => (DonGiaChuan.HasValue && HaoPhiDuToan != null) ? HaoPhiDuToan.DonGia - DonGiaChuan.Value : 0;
    
    // Tính toán Thành tiền thẩm định
    public decimal ThanhTienChuan => (HaoPhiChuan != null && DonGiaChuan.HasValue) ? HaoPhiChuan.DinhMuc * DonGiaChuan.Value : 0;
    public decimal ThanhTienDuToan => (HaoPhiDuToan != null) ? (HaoPhiDuToan.ThanhTien > 0 ? HaoPhiDuToan.ThanhTien : HaoPhiDuToan.DinhMuc * HaoPhiDuToan.DonGia) : 0;
    public decimal ChenhLechThanhTien => ThanhTienDuToan - ThanhTienChuan;
    
    public HaoPhiThamDinh? HaoPhiDuToan { get; set; }
}

/// <summary>
/// Tổng hợp vật tư từ dự toán để so sánh đơn giá
/// </summary>
public class VatTuGiaModel
{
    public string MaHieu { get; set; } = string.Empty;
    public string TenVatTu { get; set; } = string.Empty;
    public string DonVi { get; set; } = string.Empty;
    public LoaiHaoPhi LoaiHP { get; set; }
    public decimal KhoiLuong { get; set; }
    public decimal GiaDuToan { get; set; }
    public decimal? GiaChuan { get; set; }
    
    public decimal ChenhLechGia => (GiaChuan.HasValue) ? GiaDuToan - GiaChuan.Value : 0;
}

/// <summary>
/// Thông tin hành chính & pháp lý để xuất Báo cáo Thẩm định theo TT 36/2026 và NĐ 10/2021
/// </summary>
public class ThongTinBaoCaoThamDinh
{
    public string TenDuAn { get; set; } = string.Empty;
    public string TenHangMuc { get; set; } = string.Empty;
    public string DiaDiem { get; set; } = string.Empty;
    public string ChuDauTu { get; set; } = string.Empty;
    public string DonViTuVan { get; set; } = string.Empty;
    public string DonViThamDinh { get; set; } = string.Empty;
    public string SoVanBan { get; set; } = string.Empty;
    public System.DateTime NgayLap { get; set; } = System.DateTime.Today;
    public string LoaiCongTrinh { get; set; } = "Công trình Dân dụng";
    public string CapCongTrinh { get; set; } = "Cấp III";
    public string BuocThietKe { get; set; } = "Thiết kế Bản vẽ thi công";

    // Tỷ lệ % chi phí gián tiếp theo TT 36/2026/TT-BXD
    public decimal TiLeCPC { get; set; } = 7.3m;
    public decimal TiLeNhaTam { get; set; } = 1.2m;
    public decimal TiLeKhongXacDinh { get; set; } = 2.0m;
    public decimal TiLeTNCTTT { get; set; } = 5.5m;
    public decimal TiLeVAT { get; set; } = 10.0m;

    // Tỷ lệ % theo Thông tư 38/2026 & Thông tư 36/2026 cho Tổng hợp dự toán
    public decimal TiLeQLDA { get; set; } = 3.2m;
    public decimal TiLeTuVan { get; set; } = 6.5m;
    public decimal TiLeChiPhiKhac { get; set; } = 2.5m;
    public decimal TiLeDuPhong { get; set; } = 5.0m;

    // Tùy chọn sheet xuất
    public bool XuatTongHop { get; set; } = true;
    public bool XuatChiPhiXD { get; set; } = true;
    public bool XuatDoiChieuChiTiet { get; set; } = true;
    public bool XuatGiaVatTu { get; set; } = true;
}

/// <summary>
/// Đơn giá sau thẩm định cho 1 đơn vị công tác (chốt sau Giai đoạn 1)
/// </summary>
public class DonGiaThamDinhItem
{
    public string MaHieu { get; set; } = string.Empty;
    public string TenCongTac { get; set; } = string.Empty;
    public string DonVi { get; set; } = string.Empty;

    // Đơn giá Dự toán gốc (cho 1 đơn vị sản phẩm)
    public decimal DonGiaVL_DT { get; set; }
    public decimal DonGiaNC_DT { get; set; }
    public decimal DonGiaMay_DT { get; set; }
    public decimal DonGiaTong_DT => DonGiaVL_DT + DonGiaNC_DT + DonGiaMay_DT;

    // Đơn giá Thẩm định chuẩn (cho 1 đơn vị sản phẩm)
    public decimal DonGiaVL_TD { get; set; }
    public decimal DonGiaNC_TD { get; set; }
    public decimal DonGiaMay_TD { get; set; }
    public decimal DonGiaTong_TD => DonGiaVL_TD + DonGiaNC_TD + DonGiaMay_TD;

    public string LuaChon { get; set; } = "Theo Thẩm định";
    public bool CoSaiKhac { get; set; }
}

