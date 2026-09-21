#nullable enable
using AIE.Core.Models;

namespace AIE.Core.Services.LapDuToan;

/// <summary>
/// Đồng bộ dữ liệu giữa bản quét từ sheet Excel (BOQ) và bản ghi nhớ trong bộ nhớ (CurrentDuToan).
///
/// Nguyên tắc nguồn dữ liệu duy nhất:
/// - Khối lượng/đơn giá/tiên lượng (BOQ)  →  lấy từ sheet Excel (người dùng sửa trực tiếp).
/// - Dữ liệu kinh phí & thông tin cấu hình (BangKinhPhi, ChiPhiXD, ChiPhiThietBi, QLDA,
///   Tư vấn, Chi phí khác, Dự phòng, SoBuocThietKe, CapCongTrinh, ChuDauTu, ...)  →  lấy từ bộ nhớ,
///   vì các sheet Excel hiện KHÔNG lưu trữ/đọc lại các chỉ tiêu này.
/// </summary>
public static class DuToanSyncHelper
{
    /// <summary>
    /// Gộp bản quét từ sheet với bản ghi nhớ, trả về đối tượng DuToan mới chứa đầy đủ dữ liệu
    /// (BOQ từ sheet + toàn bộ dữ liệu kinh phí/cấu hình từ bộ nhớ).
    /// </summary>
    public static DuToan Merge(DuToan scanned, DuToan memory)
    {
        if (scanned == null) throw new System.ArgumentNullException(nameof(scanned));
        if (memory == null) return scanned;

        var result = new DuToan();

        // --- Từ sheet (BOQ) ---
        result.TenCongTrinh = scanned.TenCongTrinh;
        result.DiaDiem = scanned.DiaDiem;
        result.VungApDung = scanned.VungApDung;
        result.DanhSachHangMuc = scanned.DanhSachHangMuc;

        // --- Từ bộ nhớ: dữ liệu kinh phí & cấu hình không đọc được từ sheet ---
        result.BoDonGiaId = memory.BoDonGiaId;
        result.ChiPhiXD = memory.ChiPhiXD;
        result.BangTongHop = memory.BangTongHop;
        result.LoaiCongTrinh = memory.LoaiCongTrinh;
        result.TenDuAn = !string.IsNullOrWhiteSpace(memory.TenDuAn) ? memory.TenDuAn : scanned.TenDuAn;
        result.CapCongTrinh = memory.CapCongTrinh;
        result.ChuDauTu = memory.ChuDauTu;
        result.SoBuocThietKe = memory.SoBuocThietKe;
        result.Vung = memory.Vung;

        result.ChiPhiThietBi = memory.ChiPhiThietBi;
        result.ChiPhiQLDA = memory.ChiPhiQLDA;
        result.TiLeQLDA = memory.TiLeQLDA;
        result.ChiPhiTuVan = memory.ChiPhiTuVan;
        result.ChiTietTuVan = memory.ChiTietTuVan;
        result.ChiPhiKhac = memory.ChiPhiKhac;
        result.ChiPhiDuPhong = memory.ChiPhiDuPhong;
        result.DuPhongKhoiLuong = memory.DuPhongKhoiLuong;
        result.DuPhongTruotGia = memory.DuPhongTruotGia;
        result.BangKinhPhi = memory.BangKinhPhi;

        // --- Sao chép hao phí & dữ liệu kinh phí từng hạng mục ---
        GanHaoPhiVaChiPhiHangMuc(result.DanhSachHangMuc, memory.DanhSachHangMuc);

        return result;
    }

    /// <summary>
    /// Sao chép hao phí và dữ liệu kinh phí từng hạng mục (ChiPhiXD, PhanLoaiPhu, CapCongTrinh)
    /// từ bộ nhớ sang danh sách hạng mục vừa quét từ sheet, dùng bản ghi nhớ làm nguồn ưu tiên.
    /// Khớp hạng mục theo RowIndex, dự phòng theo STT; khớp công tác theo MaHieu, dự phòng theo vị trí.
    /// </summary>
    public static void GanHaoPhiVaChiPhiHangMuc(List<HangMuc> scannedHms, List<HangMuc> memoryHms)
    {
        if (scannedHms == null || memoryHms == null || memoryHms.Count == 0) return;

        var memByRow = new System.Collections.Generic.Dictionary<int, HangMuc>();
        foreach (var memHm in memoryHms)
            if (memHm.RowIndex > 0 && !memByRow.ContainsKey(memHm.RowIndex))
                memByRow[memHm.RowIndex] = memHm;

        for (int i = 0; i < scannedHms.Count; i++)
        {
            var sc = scannedHms[i];
            if (sc == null) continue;

            HangMuc? memHm = null;
            // 1. Ưu tiên khớp theo Tên hạng mục (chính xác nhất khi hạng mục bị thêm/xóa/đổi chỗ)
            if (!string.IsNullOrWhiteSpace(sc.TenHangMuc))
            {
                memHm = memoryHms.Find(m => string.Equals(m.TenHangMuc?.Trim(), sc.TenHangMuc.Trim(), System.StringComparison.OrdinalIgnoreCase));
            }
            // 2. Khớp theo Mã hạng mục nếu có
            if (memHm == null && !string.IsNullOrWhiteSpace(sc.MaHangMuc))
            {
                memHm = memoryHms.Find(m => string.Equals(m.MaHangMuc?.Trim(), sc.MaHangMuc.Trim(), System.StringComparison.OrdinalIgnoreCase));
            }
            // 3. Khớp theo RowIndex (nếu hàng dòng khớp)
            if (memHm == null && sc.RowIndex > 0 && memByRow.TryGetValue(sc.RowIndex, out var byRow))
            {
                memHm = byRow;
            }
            // 4. Dự phòng khi số lượng hạng mục bằng nhau
            if (memHm == null && scannedHms.Count == memoryHms.Count && i < memoryHms.Count)
            {
                memHm = memoryHms[i];
            }

            if (memHm == null) continue;

            GanHaoPhiChoTungCongTac(sc.DanhSachCongTac, memHm.DanhSachCongTac);

            // Dữ liệu kinh phí riêng từng hạng mục (sheet không đọc các chỉ tiêu này)
            if (sc.ChiPhiXD == null) sc.ChiPhiXD = memHm.ChiPhiXD;
            if (string.IsNullOrEmpty(sc.PhanLoaiPhu)) sc.PhanLoaiPhu = memHm.PhanLoaiPhu;
            if (string.IsNullOrEmpty(sc.CapCongTrinh)) sc.CapCongTrinh = memHm.CapCongTrinh;
            if (string.IsNullOrEmpty(sc.LoaiCongTrinh)) sc.LoaiCongTrinh = memHm.LoaiCongTrinh;
            if (sc.HeSoVL <= 0 || sc.HeSoVL == 1.0m) sc.HeSoVL = memHm.HeSoVL;
            if (sc.HeSoNC <= 0 || sc.HeSoNC == 1.0m) sc.HeSoNC = memHm.HeSoNC;
            if (sc.HeSoM <= 0 || sc.HeSoM == 1.0m) sc.HeSoM = memHm.HeSoM;
        }
    }

    private static void GanHaoPhiChoTungCongTac(
        System.Collections.Generic.List<DongDuToan> scannedCts,
        System.Collections.Generic.List<DongDuToan> memoryCts)
    {
        if (scannedCts == null || memoryCts == null || memoryCts.Count == 0) return;

        var memByMaHieu = new System.Collections.Generic.Dictionary<string, DongDuToan>();
        foreach (var ct in memoryCts)
        {
            if (ct == null || string.IsNullOrEmpty(ct.MaHieu)) continue;
            if (!memByMaHieu.ContainsKey(ct.MaHieu)) memByMaHieu[ct.MaHieu] = ct;
        }

        for (int i = 0; i < scannedCts.Count; i++)
        {
            var sc = scannedCts[i];
            if (sc == null) continue;

            DongDuToan? memCt = null;
            if (!string.IsNullOrEmpty(sc.MaHieu)) memByMaHieu.TryGetValue(sc.MaHieu, out memCt);
            if (memCt == null && i < memoryCts.Count) memCt = memoryCts[i];
            if (memCt != null && memCt.DanhSachHaoPhi != null)
                sc.DanhSachHaoPhi = memCt.DanhSachHaoPhi;
        }
    }

    /// <summary>
    /// Ghi toàn bộ dữ liệu của <paramref name="source"/> vào chính instance <paramref name="target"/>
    /// (GIỮ NGUYÊN tham chiếu object) để các form đang mở vẫn tham chiếu đúng đối tượng.
    /// </summary>
    public static void OverwriteInto(DuToan target, DuToan source)
    {
        if (target == null || source == null) return;

        target.TenDuAn = source.TenDuAn;
        target.TenCongTrinh = source.TenCongTrinh;
        target.LoaiCongTrinh = source.LoaiCongTrinh;
        target.CapCongTrinh = source.CapCongTrinh;
        target.SoBuocThietKe = source.SoBuocThietKe;
        target.ChuDauTu = source.ChuDauTu;
        target.DiaDiem = source.DiaDiem;
        target.VungApDung = source.VungApDung;
        target.BoDonGiaId = source.BoDonGiaId;
        target.Vung = source.Vung;

        target.DanhSachHangMuc = source.DanhSachHangMuc;
        target.BangTongHop = source.BangTongHop;
        target.ChiPhiXD = source.ChiPhiXD;
        target.ChiPhiThietBi = source.ChiPhiThietBi;
        target.ChiPhiQLDA = source.ChiPhiQLDA;
        target.TiLeQLDA = source.TiLeQLDA;
        target.ChiPhiTuVan = source.ChiPhiTuVan;
        target.ChiTietTuVan = source.ChiTietTuVan;
        target.ChiPhiKhac = source.ChiPhiKhac;
        target.ChiPhiDuPhong = source.ChiPhiDuPhong;
        target.DuPhongKhoiLuong = source.DuPhongKhoiLuong;
        target.DuPhongTruotGia = source.DuPhongTruotGia;
        target.BangKinhPhi = source.BangKinhPhi;
    }
}