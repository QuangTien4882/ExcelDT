using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using AIE.Core.Enums;
using AIE.Core.Models;
using AIE.ExcelAddIn.Helpers;
using AIE.ExcelAddIn.Services;
using ExcelDna.Integration;
using ExcelWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
using ExcelHAlign = Microsoft.Office.Interop.Excel.XlHAlign;
using ExcelVAlign = Microsoft.Office.Interop.Excel.XlVAlign;
using ExcelLineStyle = Microsoft.Office.Interop.Excel.XlLineStyle;
using ExcelBorderWeight = Microsoft.Office.Interop.Excel.XlBorderWeight;
using ExcelPageOrientation = Microsoft.Office.Interop.Excel.XlPageOrientation;

namespace AIE.ExcelAddIn.Forms;

/// <summary>
/// Nhóm đại diện cho một Công tác có sai khác
/// </summary>
public class CongTacSaiKhacGroup
{
    public int STT { get; set; }
    public string TenHangMuc { get; set; } = string.Empty;
    public int DongExcel { get; set; }
    public string MaCongTac { get; set; } = string.Empty;
    public string TenCongTac { get; set; } = string.Empty;
    public string DonVi { get; set; } = string.Empty;
    public decimal KhoiLuong { get; set; } // Khối lượng toàn bộ công tác (có thể chỉnh sửa trực tiếp)

    public KetQuaCongTacThamDinh KetQuaCongTac { get; set; } = null!;
    public List<HaoPhiSaiKhacItem> DanhSachHaoPhi { get; set; } = new();

    public decimal DonGiaDT { get; set; }
    public decimal DonGiaTD { get; set; }

    public bool IsTamTinh => KetQuaCongTac?.DinhMucChuan == null;

    // Thành tiền toàn bộ công tác theo Dự toán: Luôn bằng KL x Đơn giá dự toán
    public decimal ThanhTienDT => Math.Round(KhoiLuong * DonGiaDT, 0);

    // Chênh lệch do thẩm định các hao phí được chọn áp dụng
    public decimal ChenhLech => DanhSachHaoPhi.Count > 0
        ? DanhSachHaoPhi.Sum(h => h.LuaChon == "Theo Dự toán" ? 0m : h.ChenhLech)
        : (LuaChon == "Theo Dự toán" ? 0m : Math.Round(KhoiLuong * (DonGiaTD - DonGiaDT), 0));

    // Thành tiền toàn bộ công tác theo Thẩm định = Thành tiền Dự toán + Chênh lệch thẩm định
    public decimal ThanhTienTD => ThanhTienDT + ChenhLech;

    // Đơn giá Thẩm định hiển thị cho 1 đơn vị công tác
    public decimal DonGiaTDHienThi => KhoiLuong > 0 ? Math.Round(ThanhTienTD / KhoiLuong, 0) : DonGiaTD;

    // Lựa chọn cấp công tác: "Theo Thẩm định", "Theo Dự toán", hoặc "Tùy biến"
    public string LuaChon { get; set; } = "Theo Thẩm định";

    // Trạng thái mở rộng xem chi tiết hao phí (mặc định thu gọn)
    public bool IsExpanded { get; set; } = false;
}

/// <summary>
/// Đại diện cho một hao phí cụ thể có sai khác thuộc một Công tác
/// </summary>
public class HaoPhiSaiKhacItem
{
    public CongTacSaiKhacGroup ParentGroup { get; set; } = null!;
    public SaiLechDinhMuc? SaiLech { get; set; }

    public string MaHaoPhi { get; set; } = string.Empty;
    public string TenHaoPhi { get; set; } = string.Empty;
    public LoaiHaoPhi? LoaiHP { get; set; }
    public string DonVi { get; set; } = string.Empty;
    public string LoaiSaiKhac { get; set; } = string.Empty;

    // Định mức 1 đơn vị công tác
    public decimal DinhMucDT { get; set; }
    public decimal DinhMucTD { get; set; }

    // Khối lượng hao phí toàn bộ = KL_CT * DinhMuc
    public decimal KhoiLuongHaoPhiDT => ParentGroup.KhoiLuong * DinhMucDT;
    public decimal KhoiLuongHaoPhiTD => ParentGroup.KhoiLuong * DinhMucTD;

    // Đơn giá của 1 đơn vị vật tư (ví dụ: 1 kg xi măng, 1 m3 đá)
    public decimal DonGiaDT { get; set; }
    public decimal DonGiaTD { get; set; }

    // Đơn giá tính cho 1 đơn vị công tác theo định mức (ĐM * ĐG)
    public decimal DonGiaTheoDinhMucDT => Math.Round(DinhMucDT * DonGiaDT, 0);
    public decimal DonGiaTheoDinhMucTD => Math.Round(DinhMucTD * DonGiaTD, 0);

    // Thành tiền toàn bộ = KL_CT * DinhMuc * DonGia
    public decimal ThanhTienDT => Math.Round(ParentGroup.KhoiLuong * DinhMucDT * DonGiaDT, 0);
    public decimal ThanhTienTD => Math.Round(ParentGroup.KhoiLuong * DinhMucTD * DonGiaTD, 0);
    public decimal ChenhLech => ThanhTienTD - ThanhTienDT;

    public string LuaChon { get; set; } = "Theo Thẩm định";
}

/// <summary>
/// Form Xác nhận sai khác định mức, đơn giá theo nhóm mã hiệu công tác
/// </summary>
public class XacNhanSaiKhacForm : Form
{
    private readonly List<KetQuaCongTacThamDinh> _results;
    private readonly ThamDinhConfig _config;
    private readonly int? _boDonGiaId;
    private readonly string _tenCongTrinh;

    private List<CongTacSaiKhacGroup> _allGroups = new();
    private List<CongTacSaiKhacGroup> _displayGroups = new();
    private bool _isUpdatingGrid = false;
    private string _selectedSheetName = string.Empty;

    // Controls
    private Label lblTieuDe;
    private Label lblThongTin;
    private TextBox txtSearch;
    private ComboBox cboFilterLoaiLoi;
    private ComboBox cboFilterLoaiHP;
    private ComboBox cboFilterLuaChon;
    private ComboBox cboFilterTyLeChenhLech;
    private ComboBox cboFilterGiaTriChenhLech;
    private Button btnPresetAll;
    private Button btnPresetBigDiff;
    private Button btnPresetCut;
    private Button btnPresetIncrease;
    private Button btnPresetSaiDinhMuc;
    private Button btnPresetSaiDonGia;
    private Label lblFilterStatus;
    private Button btnMoRongTatCa;
    private Button btnThuGonTatCa;
    private Button btnApDungTatCaTD;
    private Button btnApDungTatCaDT;
    private Button btnChiDinhKL;
    private Button btnXoaCongTac;

    private DataGridView dgvSaiKhac;

    // Summary controls
    private Label lblCountInfo;
    private Label lblTongVL;
    private Label lblTongNC;
    private Label lblTongMay;
    private Label lblTongT;
    private Label lblChenhLechSoVoiDT;

    private Button btnTiepTuc;
    private Button btnDong;
    private Button btnXuatDoiChieu;

    /// <summary>
    /// Bắt buộc Windows tạo cửa sổ cấp ứng dụng độc lập trên Taskbar và danh sách Alt+Tab
    /// </summary>
    protected override CreateParams CreateParams
    {
        get
        {
            CreateParams cp = base.CreateParams;
            cp.ExStyle |= 0x00040000; // WS_EX_APPWINDOW
            cp.Parent = IntPtr.Zero;
            return cp;
        }
    }

    public XacNhanSaiKhacForm(
        List<KetQuaCongTacThamDinh> results,
        ThamDinhConfig config,
        int? boDonGiaId,
        string tenCongTrinh = "")
    {
        _results = results ?? new List<KetQuaCongTacThamDinh>();
        _config = config ?? new ThamDinhConfig();
        _boDonGiaId = boDonGiaId;
        _tenCongTrinh = !string.IsNullOrEmpty(tenCongTrinh) ? tenCongTrinh : "Công trình xây dựng";

        KhoiTaoDuLieuNhom();
        InitializeComponent();
        NapDuLieuLenGrid();
        CapNhatTongHopChiPhiTrucTiep();

        FormStateHelper.Attach(this);
    }

    private void KhoiTaoDuLieuNhom()
    {
        _allGroups = new List<CongTacSaiKhacGroup>();
        int stt = 1;

        foreach (var kq in _results)
        {
            decimal kl = kq.DuToan.KhoiLuong > 0 ? kq.DuToan.KhoiLuong : 1.0m;

            var group = new CongTacSaiKhacGroup
            {
                STT = stt++,
                MaCongTac = kq.DuToan.MaHieu,
                TenCongTac = kq.DuToan.TenCongTac,
                DonVi = kq.DuToan.DonVi,
                KhoiLuong = kl,
                DonGiaDT = kq.DuToan.DonGia,
                DonGiaTD = kq.DuToan.DonGia,
                KetQuaCongTac = kq,
                LuaChon = "Theo Thẩm định"
            };

            // 1. Trường hợp công tác tạm tính (không có hao phí chi tiết trong TT38)
            if (kq.DinhMucChuan == null)
            {
                decimal ttDT = kq.DuToan.ThanhTien > 0 ? kq.DuToan.ThanhTien : Math.Round(kl * kq.DuToan.DonGia, 0);
                group.DonGiaDT = kq.DuToan.DonGia > 0 ? kq.DuToan.DonGia : (kl > 0 ? Math.Round(ttDT / kl, 0) : ttDT);
                group.DonGiaTD = group.DonGiaDT;

                _allGroups.Add(group);
                continue;
            }

            // 2. Trường hợp công tác có hao phí chi tiết: thu thập các hao phí có sai khác
            foreach (var s in kq.DanhSachSaiLech)
            {
                bool hasSaiKhac = !string.IsNullOrEmpty(s.LoaiLoi)
                    || Math.Abs(s.ChenhLechDinhMuc) > 0.0001m
                    || (s.DonGiaChuan.HasValue && s.HaoPhiDuToan != null && Math.Abs(s.HaoPhiDuToan.DonGia - s.DonGiaChuan.Value) > 0.01m)
                    || Math.Abs(s.ChenhLechThanhTien) > 0.01m;

                if (!hasSaiKhac) continue;

                LoaiHaoPhi? loai = s.LoaiHP ?? s.HaoPhiChuan?.LoaiHaoPhi ?? s.HaoPhiDuToan?.Loai;
                string maHp = s.HaoPhiChuan?.MaHieuHP ?? s.HaoPhiDuToan?.MaHieuHP ?? "";
                string tenHp = s.HaoPhiChuan?.TenHaoPhi ?? s.HaoPhiDuToan?.TenHaoPhi ?? "";
                string donVi = s.HaoPhiChuan?.DonVi ?? s.HaoPhiDuToan?.DonVi ?? "";

                decimal dmDT = s.HaoPhiDuToan?.DinhMuc ?? 0m;
                decimal dgDT = s.HaoPhiDuToan?.DonGia ?? 0m;

                decimal dmTD = s.HaoPhiChuan?.DinhMuc ?? 0m;
                decimal dgTD = s.DonGiaChuan ?? dgDT;

                string moTaLoi = "";
                if (s.LoaiLoi == "Hao phí thừa") moTaLoi = "Hao phí thừa (Dự toán có, TT38 không)";
                else if (s.LoaiLoi == "Thiếu hao phí") moTaLoi = "Thiếu hao phí (TT38 có, Dự toán thiếu)";
                else if (Math.Abs(s.ChenhLechDinhMuc) > 0.0001m && s.DonGiaChuan.HasValue && Math.Abs(dgDT - dgTD) > 0.01m) moTaLoi = "Sai cả ĐM & Đơn giá";
                else if (Math.Abs(s.ChenhLechDinhMuc) > 0.0001m) moTaLoi = "Sai định mức hao phí";
                else if (s.DonGiaChuan.HasValue && Math.Abs(dgDT - dgTD) > 0.01m) moTaLoi = "Sai đơn giá";
                else if (!string.IsNullOrEmpty(s.LoaiLoi)) moTaLoi = s.LoaiLoi;
                else moTaLoi = "Chênh lệch đơn giá";

                var child = new HaoPhiSaiKhacItem
                {
                    ParentGroup = group,
                    SaiLech = s,
                    MaHaoPhi = maHp,
                    TenHaoPhi = tenHp,
                    LoaiHP = loai,
                    DonVi = donVi,
                    LoaiSaiKhac = moTaLoi,
                    DinhMucDT = dmDT,
                    DinhMucTD = dmTD,
                    DonGiaDT = dgDT,
                    DonGiaTD = dgTD,
                    LuaChon = "Theo Thẩm định"
                };

                group.DanhSachHaoPhi.Add(child);
            }

            // Nếu công tác này có sai khác hoặc là công tác tạm tính, đưa vào danh sách hiển thị
            if (group.DanhSachHaoPhi.Count > 0 || group.IsTamTinh)
            {
                _allGroups.Add(group);
            }
        }
    }

    /// <summary>
    /// Tái tạo danh sách các công tác độc lập theo từng Hạng mục từ kết quả chọn cột khối lượng Excel.
    /// Gán định mức và đơn giá chuẩn cho từng dòng mã hiệu thuộc từng Hạng mục (tuyệt đối không cộng dồn khối lượng)
    /// để bố cục đối chiếu đồng bộ 1:1 với dự toán đã lập.
    /// </summary>
    public void CapNhatDanhSachCongTacTheoTungHangMuc(List<KhoiLuongMappingItem> selectedItems, string sheetName)
    {
        if (selectedItems == null || selectedItems.Count == 0) return;

        _selectedSheetName = sheetName;
        _allGroups.Clear();

        // Tạo map tra cứu theo mã đã chuẩn hóa
        var mapResults = new Dictionary<string, KetQuaCongTacThamDinh>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in _results)
        {
            string norm = NormalizeMa(r.DuToan.MaHieu);
            if (!string.IsNullOrEmpty(norm) && !mapResults.ContainsKey(norm))
            {
                mapResults[norm] = r;
            }
        }

        int stt = 1;
        foreach (var item in selectedItems)
        {
            var kq = item.MatchedResult;
            if (kq == null)
            {
                string normMa = NormalizeMa(item.MaCongTac);
                if (!string.IsNullOrEmpty(normMa))
                {
                    mapResults.TryGetValue(normMa, out kq);
                }
            }

            decimal kl = item.KhoiLuongQuet > 0 ? item.KhoiLuongQuet : 1.0m;

            var group = new CongTacSaiKhacGroup
            {
                STT = stt++,
                TenHangMuc = !string.IsNullOrEmpty(item.TenHangMuc) ? item.TenHangMuc : "Hạng mục chung",
                DongExcel = item.DongExcel,
                MaCongTac = item.MaCongTac,
                TenCongTac = !string.IsNullOrEmpty(item.TenCongTac) && !item.TenCongTac.StartsWith("(") ? item.TenCongTac : item.TenExcel,
                DonVi = !string.IsNullOrEmpty(item.DonVi) ? item.DonVi : (kq?.DuToan.DonVi ?? ""),
                KhoiLuong = kl,
                DonGiaDT = kq?.DuToan.DonGia ?? item.DonGiaDT,
                DonGiaTD = kq?.DuToan.DonGia ?? item.DonGiaDT,
                KetQuaCongTac = kq ?? new KetQuaCongTacThamDinh
                {
                    DuToan = new CongTacThamDinh
                    {
                        MaHieu = item.MaCongTac,
                        TenCongTac = item.TenExcel,
                        DonVi = item.DonVi,
                        KhoiLuong = kl,
                        DonGia = item.DonGiaDT
                    }
                },
                LuaChon = "Theo Thẩm định"
            };

            // 1. Tạm tính
            if (kq == null || kq.DinhMucChuan == null)
            {
                group.DonGiaDT = kq?.DuToan.DonGia > 0 ? kq.DuToan.DonGia : item.DonGiaDT;
                group.DonGiaTD = group.DonGiaDT;
                _allGroups.Add(group);
                continue;
            }

            // 2. Có hao phí chi tiết: clone các hao phí có sai khác
            foreach (var s in kq.DanhSachSaiLech)
            {
                bool hasSaiKhac = !string.IsNullOrEmpty(s.LoaiLoi)
                    || Math.Abs(s.ChenhLechDinhMuc) > 0.0001m
                    || (s.DonGiaChuan.HasValue && s.HaoPhiDuToan != null && Math.Abs(s.HaoPhiDuToan.DonGia - s.DonGiaChuan.Value) > 0.01m)
                    || Math.Abs(s.ChenhLechThanhTien) > 0.01m;

                if (!hasSaiKhac) continue;

                LoaiHaoPhi? loai = s.LoaiHP ?? s.HaoPhiChuan?.LoaiHaoPhi ?? s.HaoPhiDuToan?.Loai;
                string maHp = s.HaoPhiChuan?.MaHieuHP ?? s.HaoPhiDuToan?.MaHieuHP ?? "";
                string tenHp = s.HaoPhiChuan?.TenHaoPhi ?? s.HaoPhiDuToan?.TenHaoPhi ?? "";
                string donVi = s.HaoPhiChuan?.DonVi ?? s.HaoPhiDuToan?.DonVi ?? "";

                decimal dmDT = s.HaoPhiDuToan?.DinhMuc ?? 0m;
                decimal dgDT = s.HaoPhiDuToan?.DonGia ?? 0m;
                decimal dmTD = s.HaoPhiChuan?.DinhMuc ?? 0m;
                decimal dgTD = s.DonGiaChuan ?? dgDT;

                string moTaLoi = "";
                if (s.LoaiLoi == "Hao phí thừa") moTaLoi = "Hao phí thừa (Dự toán có, TT38 không)";
                else if (s.LoaiLoi == "Thiếu hao phí") moTaLoi = "Thiếu hao phí (TT38 có, Dự toán thiếu)";
                else if (Math.Abs(s.ChenhLechDinhMuc) > 0.0001m && s.DonGiaChuan.HasValue && Math.Abs(dgDT - dgTD) > 0.01m) moTaLoi = "Sai cả ĐM & Đơn giá";
                else if (Math.Abs(s.ChenhLechDinhMuc) > 0.0001m) moTaLoi = "Sai định mức hao phí";
                else if (s.DonGiaChuan.HasValue && Math.Abs(dgDT - dgTD) > 0.01m) moTaLoi = "Sai đơn giá";
                else if (!string.IsNullOrEmpty(s.LoaiLoi)) moTaLoi = s.LoaiLoi;
                else moTaLoi = "Chênh lệch đơn giá";

                var child = new HaoPhiSaiKhacItem
                {
                    ParentGroup = group,
                    SaiLech = s,
                    MaHaoPhi = maHp,
                    TenHaoPhi = tenHp,
                    LoaiHP = loai,
                    DonVi = donVi,
                    LoaiSaiKhac = moTaLoi,
                    DinhMucDT = dmDT,
                    DinhMucTD = dmTD,
                    DonGiaDT = dgDT,
                    DonGiaTD = dgTD,
                    LuaChon = "Theo Thẩm định"
                };

                group.DanhSachHaoPhi.Add(child);
            }

            _allGroups.Add(group);
        }

        // Cập nhật nhãn thông tin header
        int tongSoLoi = _allGroups.Sum(g => g.DanhSachHaoPhi.Count > 0 ? g.DanhSachHaoPhi.Count : 1);
        int soHM = _allGroups.Where(x => !string.IsNullOrEmpty(x.TenHangMuc)).Select(x => x.TenHangMuc).Distinct().Count();
        lblThongTin.Text = $"Dự án: {_tenCongTrinh}   |   {soHM} Hạng mục   |   Tổng số: {_allGroups.Count} công tác độc lập   |   Tổng mục sai khác: {tongSoLoi}";
    }

    private static string NormalizeMa(string ma)
    {
        if (string.IsNullOrEmpty(ma)) return string.Empty;
        return ma.Replace(".", "").Replace(" ", "").Replace("-", "").ToUpper().Trim();
    }

    private void InitializeComponent()
    {
        this.Text = "Xác Nhận Sai Khác Định Mức, Đơn Giá & Tính Chi Phí Trực Tiếp Thẩm Định";
        this.Size = new Size(1360, 850);
        this.MinimumSize = new Size(1100, 650);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.WindowState = FormWindowState.Maximized; // Mở toàn màn hình để bao quát hết dữ liệu
        this.ShowInTaskbar = true;
        this.MinimizeBox = true;
        this.MaximizeBox = true;
        this.TopMost = false;
        this.Font = UIHelper.GetFont(10f);
        this.BackColor = Color.FromArgb(246, 248, 252);

        // =========================================================================
        // 1. TOP HEADER PANEL: TIÊU ĐỀ & THÔNG TIN DỰ ÁN (CHIỀU CAO RỘNG RÃI, KHÔNG CHE KHUẤT)
        // =========================================================================
        var pnlHeader = new Panel
        {
            Dock = DockStyle.Top,
            Height = 90,
            BackColor = Color.FromArgb(0, 51, 102),
            Padding = new Padding(20, 10, 20, 10)
        };

        lblTieuDe = new Label
        {
            Text = "📋 BƯỚC 1: XÁC NHẬN SAI KHÁC ĐỊNH MỨC & ĐƠN GIÁ (THEO TỪNG MÃ HIỆU)",
            Font = UIHelper.GetFont(12.5f, FontStyle.Bold),
            ForeColor = Color.White,
            AutoSize = true,
            Location = new Point(20, 12)
        };

        int tongSoLoi = _allGroups.Sum(g => g.DanhSachHaoPhi.Count > 0 ? g.DanhSachHaoPhi.Count : 1);
        lblThongTin = new Label
        {
            Text = $"Dự án: {_tenCongTrinh}   |   Tổng số công tác: {_results.Count}   |   Số công tác có sai khác: {_allGroups.Count}   |   Tổng mục sai khác: {tongSoLoi}",
            Font = UIHelper.GetFont(9.5f),
            ForeColor = Color.FromArgb(210, 230, 255),
            AutoSize = true,
            Location = new Point(20, 48)
        };

        pnlHeader.Controls.Add(lblTieuDe);
        pnlHeader.Controls.Add(lblThongTin);

        // =========================================================================
        // 2. TOOLBAR PANEL: TÌM KIẾM, BỘ LỌC, CHỈ ĐỊNH CỘT KL VÀ XÓA CÔNG TÁC THỪA
        // =========================================================================
        // =========================================================================
        // 2. TOOLBAR PANEL: SMART AUDIT FILTER (TIÊU ĐIỂM NÓNG, PRESETS, DROPDOWNS)
        // =========================================================================
        var pnlFilterArea = new Panel
        {
            Dock = DockStyle.Top,
            Height = 96,
            BackColor = Color.White
        };

        // Hàng 1: Tiêu điểm nhanh (Smart Presets) & Trạng thái thống kê bộ lọc
        var pnlPresets = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 44,
            BackColor = Color.FromArgb(246, 249, 253),
            WrapContents = false,
            AutoScroll = true,
            Padding = new Padding(10, 5, 10, 3)
        };

        var lblPresetTitle = new Label
        {
            Text = "⚡ Tiêu điểm:",
            AutoSize = true,
            Margin = new Padding(2, 8, 4, 0),
            Font = UIHelper.GetFont(9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 102, 204)
        };
        pnlPresets.Controls.Add(lblPresetTitle);

        btnPresetAll = new Button
        {
            Text = "🔍 Tất cả (Đặt lại)",
            AutoSize = true,
            Height = 30,
            Padding = new Padding(8, 0, 8, 0),
            Margin = new Padding(0, 2, 6, 0),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(50, 50, 50),
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnPresetAll.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
        btnPresetAll.Click += (s, e) =>
        {
            txtSearch.Text = "";
            if (cboFilterLoaiLoi != null) cboFilterLoaiLoi.SelectedIndex = 0;
            if (cboFilterLoaiHP != null) cboFilterLoaiHP.SelectedIndex = 0;
            if (cboFilterLuaChon != null) cboFilterLuaChon.SelectedIndex = 0;
            if (cboFilterTyLeChenhLech != null) cboFilterTyLeChenhLech.SelectedIndex = 0;
            if (cboFilterGiaTriChenhLech != null) cboFilterGiaTriChenhLech.SelectedIndex = 0;
            NapDuLieuLenGrid();
        };
        pnlPresets.Controls.Add(btnPresetAll);

        btnPresetBigDiff = new Button
        {
            Text = "🔥 Chênh lệch lớn (>5% hoặc >5tr)",
            AutoSize = true,
            Height = 30,
            Padding = new Padding(8, 0, 8, 0),
            Margin = new Padding(0, 2, 6, 0),
            BackColor = Color.FromArgb(255, 245, 235),
            ForeColor = Color.FromArgb(204, 85, 0),
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnPresetBigDiff.FlatAppearance.BorderColor = Color.FromArgb(240, 160, 80);
        btnPresetBigDiff.Click += (s, e) =>
        {
            if (cboFilterTyLeChenhLech != null) cboFilterTyLeChenhLech.SelectedItem = "Chênh lệch > 5%";
            if (cboFilterGiaTriChenhLech != null) cboFilterGiaTriChenhLech.SelectedIndex = 0;
            NapDuLieuLenGrid();
        };
        pnlPresets.Controls.Add(btnPresetBigDiff);

        btnPresetCut = new Button
        {
            Text = "🔻 Cắt giảm (>1tr đ)",
            AutoSize = true,
            Height = 30,
            Padding = new Padding(8, 0, 8, 0),
            Margin = new Padding(0, 2, 6, 0),
            BackColor = Color.FromArgb(255, 240, 240),
            ForeColor = Color.FromArgb(190, 30, 30),
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnPresetCut.FlatAppearance.BorderColor = Color.FromArgb(240, 120, 120);
        btnPresetCut.Click += (s, e) =>
        {
            if (cboFilterGiaTriChenhLech != null) cboFilterGiaTriChenhLech.SelectedItem = "🔻 Chỉ mục GIẢM (TĐ < ĐT)";
            if (cboFilterTyLeChenhLech != null) cboFilterTyLeChenhLech.SelectedIndex = 0;
            NapDuLieuLenGrid();
        };
        pnlPresets.Controls.Add(btnPresetCut);

        btnPresetIncrease = new Button
        {
            Text = "🔺 Tăng chi phí (>1tr đ)",
            AutoSize = true,
            Height = 30,
            Padding = new Padding(8, 0, 8, 0),
            Margin = new Padding(0, 2, 6, 0),
            BackColor = Color.FromArgb(235, 248, 240),
            ForeColor = Color.FromArgb(0, 140, 60),
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnPresetIncrease.FlatAppearance.BorderColor = Color.FromArgb(100, 200, 140);
        btnPresetIncrease.Click += (s, e) =>
        {
            if (cboFilterGiaTriChenhLech != null) cboFilterGiaTriChenhLech.SelectedItem = "🔺 Chỉ mục TĂNG (TĐ > ĐT)";
            if (cboFilterTyLeChenhLech != null) cboFilterTyLeChenhLech.SelectedIndex = 0;
            NapDuLieuLenGrid();
        };
        pnlPresets.Controls.Add(btnPresetIncrease);

        btnPresetSaiDinhMuc = new Button
        {
            Text = "📋 Sai định mức TT38",
            AutoSize = true,
            Height = 30,
            Padding = new Padding(8, 0, 8, 0),
            Margin = new Padding(0, 2, 6, 0),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(0, 102, 204),
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnPresetSaiDinhMuc.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 215);
        btnPresetSaiDinhMuc.Click += (s, e) =>
        {
            if (cboFilterLoaiLoi != null) cboFilterLoaiLoi.SelectedItem = "Sai định mức";
            NapDuLieuLenGrid();
        };
        pnlPresets.Controls.Add(btnPresetSaiDinhMuc);

        btnPresetSaiDonGia = new Button
        {
            Text = "🏷 Sai đơn giá VT",
            AutoSize = true,
            Height = 30,
            Padding = new Padding(8, 0, 8, 0),
            Margin = new Padding(0, 2, 12, 0),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(120, 40, 160),
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnPresetSaiDonGia.FlatAppearance.BorderColor = Color.FromArgb(160, 100, 200);
        btnPresetSaiDonGia.Click += (s, e) =>
        {
            if (cboFilterLoaiLoi != null) cboFilterLoaiLoi.SelectedItem = "Sai đơn giá";
            NapDuLieuLenGrid();
        };
        pnlPresets.Controls.Add(btnPresetSaiDonGia);

        lblFilterStatus = new Label
        {
            Text = "📊 Đang hiển thị...",
            AutoSize = true,
            Margin = new Padding(6, 8, 6, 0),
            Font = UIHelper.GetFont(9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 102, 204)
        };
        pnlPresets.Controls.Add(lblFilterStatus);

        // Hàng 2: Bộ lọc chi tiết, tìm kiếm & Các nút thao tác
        var pnlFilter = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            WrapContents = false,
            AutoScroll = true,
            Padding = new Padding(10, 6, 10, 6)
        };

        var lblSearch = new Label
        {
            Text = "🔍 Tìm kiếm:",
            AutoSize = true,
            Margin = new Padding(3, 7, 3, 0),
            Font = UIHelper.GetFont(9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(50, 50, 50)
        };

        txtSearch = new TextBox
        {
            Width = 125,
            Margin = new Padding(0, 4, 10, 0),
            Font = UIHelper.GetFont(9.5f)
        };
        txtSearch.TextChanged += (s, e) => NapDuLieuLenGrid();

        var lblTyLe = new Label
        {
            Text = "Tỷ lệ CL:",
            AutoSize = true,
            Margin = new Padding(3, 7, 3, 0),
            Font = UIHelper.GetFont(9.5f),
            ForeColor = Color.FromArgb(50, 50, 50)
        };

        cboFilterTyLeChenhLech = new ComboBox
        {
            Width = 125,
            Margin = new Padding(0, 4, 10, 0),
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = UIHelper.GetFont(9.5f)
        };
        cboFilterTyLeChenhLech.Items.AddRange(new object[] { "Tất cả tỷ lệ", "Chênh lệch > 5%", "Chênh lệch > 10%", "Chênh lệch > 20%", "Chênh lệch > 50%", "Khớp (0%)" });
        cboFilterTyLeChenhLech.SelectedIndex = 0;
        cboFilterTyLeChenhLech.SelectedIndexChanged += (s, e) => NapDuLieuLenGrid();

        var lblGiaTriTien = new Label
        {
            Text = "Mức tiền:",
            AutoSize = true,
            Margin = new Padding(3, 7, 3, 0),
            Font = UIHelper.GetFont(9.5f),
            ForeColor = Color.FromArgb(50, 50, 50)
        };

        cboFilterGiaTriChenhLech = new ComboBox
        {
            Width = 135,
            Margin = new Padding(0, 4, 10, 0),
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = UIHelper.GetFont(9.5f)
        };
        cboFilterGiaTriChenhLech.Items.AddRange(new object[] { "Tất cả mức tiền", "|CL| > 1 triệu", "|CL| > 5 triệu", "|CL| > 20 triệu", "🔻 Chỉ mục GIẢM (TĐ < ĐT)", "🔺 Chỉ mục TĂNG (TĐ > ĐT)" });
        cboFilterGiaTriChenhLech.SelectedIndex = 0;
        cboFilterGiaTriChenhLech.SelectedIndexChanged += (s, e) => NapDuLieuLenGrid();

        var lblLoaiLoi = new Label
        {
            Text = "Sai khác:",
            AutoSize = true,
            Margin = new Padding(3, 7, 3, 0),
            Font = UIHelper.GetFont(9.5f),
            ForeColor = Color.FromArgb(50, 50, 50)
        };

        cboFilterLoaiLoi = new ComboBox
        {
            Width = 145,
            Margin = new Padding(0, 4, 10, 0),
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = UIHelper.GetFont(9.5f)
        };
        cboFilterLoaiLoi.Items.AddRange(new object[] { "Tất cả sai khác", "Sai định mức", "Sai đơn giá", "Hao phí thừa", "Thiếu hao phí", "Công tác tạm tính" });
        cboFilterLoaiLoi.SelectedIndex = 0;
        cboFilterLoaiLoi.SelectedIndexChanged += (s, e) => NapDuLieuLenGrid();

        var lblLoaiHP = new Label
        {
            Text = "Chi phí:",
            AutoSize = true,
            Margin = new Padding(3, 7, 3, 0),
            Font = UIHelper.GetFont(9.5f),
            ForeColor = Color.FromArgb(50, 50, 50)
        };

        cboFilterLoaiHP = new ComboBox
        {
            Width = 120,
            Margin = new Padding(0, 4, 10, 0),
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = UIHelper.GetFont(9.5f)
        };
        cboFilterLoaiHP.Items.AddRange(new object[] { "Tất cả", "Vật liệu (VL)", "Nhân công (NC)", "Máy thi công (M)" });
        cboFilterLoaiHP.SelectedIndex = 0;
        cboFilterLoaiHP.SelectedIndexChanged += (s, e) => NapDuLieuLenGrid();

        var lblLuaChon = new Label
        {
            Text = "Lựa chọn:",
            AutoSize = true,
            Margin = new Padding(3, 7, 3, 0),
            Font = UIHelper.GetFont(9.5f),
            ForeColor = Color.FromArgb(50, 50, 50)
        };

        cboFilterLuaChon = new ComboBox
        {
            Width = 140,
            Margin = new Padding(0, 4, 12, 0),
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = UIHelper.GetFont(9.5f)
        };
        cboFilterLuaChon.Items.AddRange(new object[] { "Tất cả trạng thái", "Theo Thẩm định", "Theo Dự toán" });
        cboFilterLuaChon.SelectedIndex = 0;
        cboFilterLuaChon.SelectedIndexChanged += (s, e) => NapDuLieuLenGrid();

        btnMoRongTatCa = new Button
        {
            Text = "➕ Mở rộng tất cả",
            AutoSize = true,
            Height = 32,
            Padding = new Padding(8, 0, 8, 0),
            Margin = new Padding(0, 3, 6, 0),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(0, 102, 204),
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnMoRongTatCa.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 215);
        btnMoRongTatCa.Click += (s, e) =>
        {
            foreach (var g in _allGroups) g.IsExpanded = true;
            NapDuLieuLenGrid();
        };

        btnThuGonTatCa = new Button
        {
            Text = "➖ Thu gọn tất cả",
            AutoSize = true,
            Height = 32,
            Padding = new Padding(8, 0, 8, 0),
            Margin = new Padding(0, 3, 10, 0),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(80, 80, 80),
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnThuGonTatCa.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180);
        btnThuGonTatCa.Click += (s, e) =>
        {
            foreach (var g in _allGroups) g.IsExpanded = false;
            NapDuLieuLenGrid();
        };

        btnApDungTatCaTD = new Button
        {
            Text = "✓ Chọn tất cả Thẩm định",
            AutoSize = true,
            Height = 32,
            Padding = new Padding(10, 0, 10, 0),
            Margin = new Padding(0, 3, 8, 0),
            BackColor = Color.FromArgb(235, 245, 255),
            ForeColor = Color.FromArgb(0, 102, 204),
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnApDungTatCaTD.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 215);
        btnApDungTatCaTD.Click += (s, e) => ApDungTatCa(theoThamDinh: true);

        btnApDungTatCaDT = new Button
        {
            Text = "↺ Chọn tất cả Dự toán",
            AutoSize = true,
            Height = 32,
            Padding = new Padding(10, 0, 10, 0),
            Margin = new Padding(0, 3, 12, 0),
            BackColor = Color.FromArgb(255, 248, 235),
            ForeColor = Color.FromArgb(180, 90, 0),
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnApDungTatCaDT.FlatAppearance.BorderColor = Color.FromArgb(230, 140, 0);
        btnApDungTatCaDT.Click += (s, e) => ApDungTatCa(theoThamDinh: false);

        btnChiDinhKL = new Button
        {
            Text = "📐 Chỉ định Cột Khối Lượng",
            AutoSize = true,
            Height = 32,
            Padding = new Padding(12, 0, 12, 0),
            Margin = new Padding(0, 3, 10, 0),
            BackColor = Color.FromArgb(230, 246, 255),
            ForeColor = Color.FromArgb(0, 102, 204),
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnChiDinhKL.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 215);
        btnChiDinhKL.Click += BtnChiDinhKL_Click;

        btnXoaCongTac = new Button
        {
            Text = "🗑 Xóa công tác thừa",
            AutoSize = true,
            Height = 32,
            Padding = new Padding(10, 0, 10, 0),
            Margin = new Padding(0, 3, 10, 0),
            BackColor = Color.FromArgb(255, 242, 242),
            ForeColor = Color.FromArgb(190, 30, 30),
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnXoaCongTac.FlatAppearance.BorderColor = Color.FromArgb(230, 120, 120);
        btnXoaCongTac.Click += BtnXoaCongTac_Click;

        pnlFilter.Controls.Add(lblSearch);
        pnlFilter.Controls.Add(txtSearch);
        pnlFilter.Controls.Add(lblTyLe);
        pnlFilter.Controls.Add(cboFilterTyLeChenhLech);
        pnlFilter.Controls.Add(lblGiaTriTien);
        pnlFilter.Controls.Add(cboFilterGiaTriChenhLech);
        pnlFilter.Controls.Add(lblLoaiLoi);
        pnlFilter.Controls.Add(cboFilterLoaiLoi);
        pnlFilter.Controls.Add(lblLoaiHP);
        pnlFilter.Controls.Add(cboFilterLoaiHP);
        pnlFilter.Controls.Add(lblLuaChon);
        pnlFilter.Controls.Add(cboFilterLuaChon);
        pnlFilter.Controls.Add(btnMoRongTatCa);
        pnlFilter.Controls.Add(btnThuGonTatCa);
        pnlFilter.Controls.Add(btnApDungTatCaTD);
        pnlFilter.Controls.Add(btnApDungTatCaDT);
        pnlFilter.Controls.Add(btnChiDinhKL);
        pnlFilter.Controls.Add(btnXoaCongTac);

        pnlFilterArea.Controls.Add(pnlFilter);
        pnlFilterArea.Controls.Add(pnlPresets);

        // =========================================================================
        // 3. BOTTOM PANEL: TỔNG HỢP CHI PHÍ TRỰC TIẾP & DUY NHẤT 1 NÚT XÁC NHẬN CHÍNH
        // =========================================================================
        var pnlBottom = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 130,
            BackColor = Color.White,
            Padding = new Padding(15, 8, 15, 8)
        };

        lblCountInfo = new Label
        {
            Text = "Thống kê: 0 công tác",
            AutoSize = true,
            Location = new Point(15, 6),
            Font = UIHelper.GetFont(9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(70, 70, 70)
        };
        pnlBottom.Controls.Add(lblCountInfo);

        int cardY = 28;
        int cardH = 58;
        int cardW = 255;

        var cardVL = TaoCardThongKe("CHI PHÍ VẬT LIỆU (VL_TĐ)", out lblTongVL, new Point(15, cardY), cardW, cardH, Color.FromArgb(0, 102, 204));
        var cardNC = TaoCardThongKe("CHI PHÍ NHÂN CÔNG (NC_TĐ)", out lblTongNC, new Point(280, cardY), cardW, cardH, Color.FromArgb(0, 102, 204));
        var cardMay = TaoCardThongKe("CHI PHÍ MÁY TC (M_TĐ)", out lblTongMay, new Point(545, cardY), cardW, cardH, Color.FromArgb(0, 102, 204));
        var cardTong = TaoCardThongKe("TỔNG TRỰC TIẾP THẨM ĐỊNH (T_TĐ)", out lblTongT, new Point(810, cardY), 310, cardH, Color.FromArgb(0, 128, 64), isLarge: true);

        pnlBottom.Controls.Add(cardVL);
        pnlBottom.Controls.Add(cardNC);
        pnlBottom.Controls.Add(cardMay);
        pnlBottom.Controls.Add(cardTong);

        lblChenhLechSoVoiDT = new Label
        {
            Text = "So với Dự toán gốc: Chênh lệch 0 đ",
            AutoSize = true,
            Location = new Point(1135, 45),
            Font = UIHelper.GetFont(10f, FontStyle.Bold),
            ForeColor = Color.FromArgb(180, 50, 50)
        };
        pnlBottom.Controls.Add(lblChenhLechSoVoiDT);

        btnTiepTuc = new Button
        {
            Text = "➔  TIẾP TỤC: ÁP ĐƠN GIÁ VÀO BẢNG DỰ TOÁN CHI TIẾT\n(TIÊN LƯỢNG KHỐI LƯỢNG)",
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            Location = new Point(pnlBottom.Width - 660, 72),
            Width = 540,
            Height = 48,
            BackColor = Color.FromArgb(16, 124, 65),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9.5f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnTiepTuc.FlatAppearance.BorderSize = 0;
        btnTiepTuc.Click += BtnTiepTuc_Click;

        btnDong = new Button
        {
            Text = "Đóng",
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            Location = new Point(pnlBottom.Width - 110, 72),
            Width = 95,
            Height = 48,
            BackColor = Color.FromArgb(240, 240, 240),
            ForeColor = Color.Black,
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(10f),
            Cursor = Cursors.Hand
        };
        btnDong.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
        btnDong.Click += (s, e) => this.Close();

        btnXuatDoiChieu = new Button
        {
            Text = "📥 Xuất Bảng đối chiếu chênh lệch",
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
            Location = new Point(20, 72),
            Width = 260,
            Height = 48,
            BackColor = Color.FromArgb(0, 102, 204),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9.5f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnXuatDoiChieu.FlatAppearance.BorderSize = 0;
        btnXuatDoiChieu.Click += (s, e) => XuatBangDoiChieuChenhLechExcel();

        pnlBottom.SizeChanged += (s, e) =>
        {
            btnTiepTuc.Location = new Point(pnlBottom.Width - 660, 72);
            btnDong.Location = new Point(pnlBottom.Width - 110, 72);
        };

        pnlBottom.Controls.Add(btnXuatDoiChieu);
        pnlBottom.Controls.Add(btnTiepTuc);
        pnlBottom.Controls.Add(btnDong);

        // =========================================================================
        // 4. MAIN DATAGRIDVIEW VỚI CỘT CỐ ĐỊNH (FROZEN COLUMNS) & AUTO CO GIÃN
        // =========================================================================
        dgvSaiKhac = new DataGridView
        {
            Dock = DockStyle.Fill,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.CellSelect,
            MultiSelect = false,
            AutoGenerateColumns = false
        };

        UIHelper.ApplyStyle(dgvSaiKhac);
        KhoiTaoCotGrid();

        dgvSaiKhac.CellValueChanged += DgvSaiKhac_CellValueChanged;
        dgvSaiKhac.CurrentCellDirtyStateChanged += (s, e) =>
        {
            if (dgvSaiKhac.IsCurrentCellDirty)
            {
                dgvSaiKhac.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        };
        dgvSaiKhac.DataError += (s, e) => { e.ThrowException = false; };
        dgvSaiKhac.SizeChanged += (s, e) => CapNhatDoRongCotTen();
        dgvSaiKhac.CellPainting += DgvSaiKhac_CellPainting;
        dgvSaiKhac.CellClick += DgvSaiKhac_CellClick;
        dgvSaiKhac.Scroll += (s, e) =>
        {
            Rectangle rtHeader = dgvSaiKhac.DisplayRectangle;
            rtHeader.Height = dgvSaiKhac.ColumnHeadersHeight;
            dgvSaiKhac.Invalidate(rtHeader);
        };
        dgvSaiKhac.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Delete)
            {
                XoaCongTacDuocChon();
            }
        };

        this.Controls.Add(dgvSaiKhac);
        this.Controls.Add(pnlBottom);
        this.Controls.Add(pnlFilterArea);
        this.Controls.Add(pnlHeader);
    }

    private Panel TaoCardThongKe(string tieuDe, out Label lblGiaTri, Point pos, int width, int height, Color color, bool isLarge = false)
    {
        var panel = new Panel
        {
            Location = pos,
            Size = new Size(width, height),
            BackColor = Color.FromArgb(248, 250, 254),
            BorderStyle = BorderStyle.FixedSingle
        };

        var lblTitle = new Label
        {
            Text = tieuDe,
            Font = UIHelper.GetFont(8.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(90, 100, 110),
            AutoSize = true,
            Location = new Point(8, 6)
        };

        lblGiaTri = new Label
        {
            Text = "0 đ",
            Font = UIHelper.GetFont(isLarge ? 12.5f : 11f, FontStyle.Bold),
            ForeColor = color,
            AutoSize = true,
            Location = new Point(8, 27)
        };

        panel.Controls.Add(lblTitle);
        panel.Controls.Add(lblGiaTri);
        return panel;
    }

    private void KhoiTaoCotGrid()
    {
        dgvSaiKhac.Columns.Clear();
        dgvSaiKhac.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        dgvSaiKhac.ColumnHeadersHeight = 54; // Chiều cao 2 tầng header: tầng trên tiêu đề nhóm, tầng dưới Dự toán / Thẩm định
        dgvSaiKhac.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
        dgvSaiKhac.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        dgvSaiKhac.ColumnHeadersDefaultCellStyle.Font = UIHelper.GetFont(9.5f, FontStyle.Bold);

        dgvSaiKhac.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
        dgvSaiKhac.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;

        // 0. STT (CỐ ĐỊNH)
        var colSTT = new DataGridViewTextBoxColumn
        {
            Name = "colSTT",
            HeaderText = "STT",
            Width = 50,
            ReadOnly = true,
            Frozen = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter, Font = UIHelper.GetFont(9.5f, FontStyle.Bold) }
        };
        dgvSaiKhac.Columns.Add(colSTT);

        // 1. Mã hiệu công tác / hao phí (CỐ ĐỊNH)
        var colMa = new DataGridViewTextBoxColumn
        {
            Name = "colMa",
            HeaderText = "Mã hiệu",
            Width = 90,
            ReadOnly = true,
            Frozen = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
        };
        dgvSaiKhac.Columns.Add(colMa);

        // 2. Tên công tác xây dựng / Hao phí chi tiết (Tự co giãn theo màn hình)
        var colTen = new DataGridViewTextBoxColumn
        {
            Name = "colTen",
            HeaderText = "Tên công tác xây dựng /\nHao phí chi tiết",
            Width = 340,
            MinimumWidth = 280,
            ReadOnly = true,
            Frozen = false,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft, WrapMode = DataGridViewTriState.True }
        };
        dgvSaiKhac.Columns.Add(colTen);

        // 3. Loại chi phí
        dgvSaiKhac.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colLoaiHP",
            HeaderText = "Phân\nloại",
            Width = 60,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
        });

        // 4. ĐVT
        dgvSaiKhac.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colDonVi",
            HeaderText = "ĐVT",
            Width = 50,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
        });

        // 5. Khối lượng toàn bộ
        dgvSaiKhac.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colKhoiLuong",
            HeaderText = "Khối lượng\ntoàn bộ",
            Width = 95,
            ReadOnly = false,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "#,##0.000" }
        });

        // 6. Loại sai khác
        dgvSaiKhac.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colLoaiSaiKhac",
            HeaderText = "Loại sai khác",
            Width = 140,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft, ForeColor = Color.FromArgb(160, 40, 40), WrapMode = DataGridViewTriState.True }
        });

        // 7. Định mức Dự toán (Nhóm: Định mức)
        dgvSaiKhac.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colDinhMucDT",
            HeaderText = "Dự toán",
            Width = 80,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "0.0000" }
        });

        // 8. Định mức Thẩm định (Nhóm: Định mức)
        dgvSaiKhac.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colDinhMucTD",
            HeaderText = "Thẩm định",
            Width = 80,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "0.0000", ForeColor = Color.FromArgb(0, 102, 204) }
        });

        // 9. Đơn giá VT Dự toán (Nhóm: Đơn giá vật tư)
        dgvSaiKhac.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colDonGiaDT",
            HeaderText = "Dự toán",
            Width = 95,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "#,##0" }
        });

        // 10. Đơn giá VT Thẩm định (Nhóm: Đơn giá vật tư)
        dgvSaiKhac.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colDonGiaTD",
            HeaderText = "Thẩm định",
            Width = 95,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "#,##0", ForeColor = Color.FromArgb(0, 102, 204) }
        });

        // 11. ĐG trong ĐM Dự toán (Nhóm: ĐG trong ĐM)
        dgvSaiKhac.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colDonGiaDmDT",
            HeaderText = "Dự toán",
            Width = 95,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "#,##0" }
        });

        // 12. ĐG trong ĐM Thẩm định (Nhóm: ĐG trong ĐM)
        dgvSaiKhac.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colDonGiaDmTD",
            HeaderText = "Thẩm định",
            Width = 95,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "#,##0", ForeColor = Color.FromArgb(0, 102, 204) }
        });

        // 13. Thành tiền Dự toán (Nhóm: Thành tiền)
        dgvSaiKhac.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colThanhTienDT",
            HeaderText = "Dự toán",
            Width = 110,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "#,##0" }
        });

        // 14. Thành tiền Thẩm định (Nhóm: Thành tiền)
        dgvSaiKhac.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colThanhTienTD",
            HeaderText = "Thẩm định",
            Width = 110,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "#,##0" }
        });

        // 15. Chênh lệch toàn bộ
        dgvSaiKhac.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colChenhLech",
            HeaderText = "Chênh lệch\n(đ)",
            Width = 100,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "#,##0" }
        });

        // 16. Lựa chọn áp dụng
        var colLuaChon = new DataGridViewComboBoxColumn
        {
            Name = "colLuaChon",
            HeaderText = "Lựa chọn\náp dụng",
            Width = 125,
            FlatStyle = FlatStyle.Flat,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
        };
        colLuaChon.Items.Add("Theo Thẩm định");
        colLuaChon.Items.Add("Theo Dự toán");
        colLuaChon.Items.Add("Tùy biến");
        dgvSaiKhac.Columns.Add(colLuaChon);
    }

    private void NapDuLieuLenGrid()
    {
        _isUpdatingGrid = true;
        try
        {
            string keyword = txtSearch?.Text?.Trim()?.ToLower() ?? "";
            string filterLoi = cboFilterLoaiLoi?.SelectedItem?.ToString() ?? "Tất cả sai khác";
            string filterHP = cboFilterLoaiHP?.SelectedItem?.ToString() ?? "Tất cả";
            string filterLC = cboFilterLuaChon?.SelectedItem?.ToString() ?? "Tất cả trạng thái";
            string filterTyLe = cboFilterTyLeChenhLech?.SelectedItem?.ToString() ?? "Tất cả tỷ lệ";
            string filterTien = cboFilterGiaTriChenhLech?.SelectedItem?.ToString() ?? "Tất cả mức tiền";

            _displayGroups = _allGroups.Where(g =>
            {
                // Nếu tìm kiếm
                if (!string.IsNullOrEmpty(keyword))
                {
                    bool matchGroup = g.MaCongTac.ToLower().Contains(keyword) || g.TenCongTac.ToLower().Contains(keyword);
                    bool matchChild = g.DanhSachHaoPhi.Any(h => h.MaHaoPhi.ToLower().Contains(keyword) || h.TenHaoPhi.ToLower().Contains(keyword) || h.LoaiSaiKhac.ToLower().Contains(keyword));
                    if (!matchGroup && !matchChild) return false;
                }

                // Lọc theo loại sai khác
                if (filterLoi != "Tất cả sai khác")
                {
                    if (filterLoi == "Công tác tạm tính" && !g.IsTamTinh) return false;
                    if (filterLoi == "Sai định mức" && !g.DanhSachHaoPhi.Any(h => h.LoaiSaiKhac.Contains("định mức") || h.LoaiSaiKhac.Contains("thừa") || h.LoaiSaiKhac.Contains("thiếu"))) return false;
                    if (filterLoi == "Sai đơn giá" && !g.DanhSachHaoPhi.Any(h => h.LoaiSaiKhac.Contains("đơn giá"))) return false;
                    if (filterLoi == "Hao phí thừa" && !g.DanhSachHaoPhi.Any(h => h.LoaiSaiKhac.Contains("thừa"))) return false;
                    if (filterLoi == "Thiếu hao phí" && !g.DanhSachHaoPhi.Any(h => h.LoaiSaiKhac.Contains("thiếu"))) return false;
                }

                // Lọc theo chi phí
                if (filterHP != "Tất cả")
                {
                    if (filterHP.StartsWith("Vật liệu") && !g.DanhSachHaoPhi.Any(h => h.LoaiHP == LoaiHaoPhi.VL)) return false;
                    if (filterHP.StartsWith("Nhân công") && !g.DanhSachHaoPhi.Any(h => h.LoaiHP == LoaiHaoPhi.NC)) return false;
                    if (filterHP.StartsWith("Máy") && !g.DanhSachHaoPhi.Any(h => h.LoaiHP == LoaiHaoPhi.MAY)) return false;
                }

                // Lọc theo lựa chọn
                if (filterLC != "Tất cả trạng thái")
                {
                    if (g.LuaChon != filterLC && !g.DanhSachHaoPhi.Any(h => h.LuaChon == filterLC)) return false;
                }

                // Lọc theo tỷ lệ chênh lệch %
                decimal pctDiff = g.ThanhTienDT != 0
                    ? Math.Abs(g.ChenhLech) / Math.Abs(g.ThanhTienDT) * 100m
                    : (g.ChenhLech != 0 ? 100m : 0m);

                if (filterTyLe == "Chênh lệch > 5%" && pctDiff <= 5m) return false;
                if (filterTyLe == "Chênh lệch > 10%" && pctDiff <= 10m) return false;
                if (filterTyLe == "Chênh lệch > 20%" && pctDiff <= 20m) return false;
                if (filterTyLe == "Chênh lệch > 50%" && pctDiff <= 50m) return false;
                if (filterTyLe == "Khớp (0%)" && Math.Abs(g.ChenhLech) > 1m) return false;

                // Lọc theo mức tiền chênh lệch
                decimal absCL = Math.Abs(g.ChenhLech);
                if (filterTien == "|CL| > 1 triệu" && absCL <= 1_000_000m) return false;
                if (filterTien == "|CL| > 5 triệu" && absCL <= 5_000_000m) return false;
                if (filterTien == "|CL| > 20 triệu" && absCL <= 20_000_000m) return false;
                if (filterTien == "🔻 Chỉ mục GIẢM (TĐ < ĐT)" && g.ChenhLech >= -1m) return false;
                if (filterTien == "🔺 Chỉ mục TĂNG (TĐ > ĐT)" && g.ChenhLech <= 1m) return false;

                return true;
            }).ToList();

            // Cập nhật trạng thái thống kê bộ lọc
            decimal tongClLoc = _displayGroups.Sum(x => x.ChenhLech);
            if (lblFilterStatus != null)
            {
                lblFilterStatus.Text = $"📊 Đang lọc: {_displayGroups.Count}/{_allGroups.Count} công tác | Tổng CL: {(tongClLoc >= 0 ? "+" : "")}{tongClLoc:N0} đ";
                lblFilterStatus.ForeColor = tongClLoc < -1m ? Color.FromArgb(190, 30, 30) : (tongClLoc > 1m ? Color.FromArgb(0, 140, 60) : Color.FromArgb(0, 102, 204));
            }

            dgvSaiKhac.Rows.Clear();
            string? lastHangMuc = null;

            for (int i = 0; i < _displayGroups.Count; i++)
            {
                var g = _displayGroups[i];

                // Chèn dòng Banner Hạng mục khi chuyển sang Hạng mục mới
                if (!string.IsNullOrEmpty(g.TenHangMuc) && g.TenHangMuc != lastHangMuc)
                {
                    lastHangMuc = g.TenHangMuc;
                    string hmKey = lastHangMuc;
                    int countHm = _displayGroups.Count(x => x.TenHangMuc == hmKey);
                    decimal hmTtDt = _displayGroups.Where(x => x.TenHangMuc == hmKey).Sum(x => x.ThanhTienDT);
                    decimal hmTtTd = _displayGroups.Where(x => x.TenHangMuc == hmKey).Sum(x => x.ThanhTienTD);
                    decimal hmCl = hmTtTd - hmTtDt;

                    int hmIdx = dgvSaiKhac.Rows.Add(
                        "",
                        "",
                        $"📁 HẠNG MỤC: {hmKey.ToUpper()}",
                        "[HẠNG MỤC]",
                        "",
                        (object)"",
                        $"({countHm} công tác)",
                        "", "", "", "", "", "",
                        hmTtDt,
                        hmTtTd,
                        hmCl,
                        ""
                    );

                    var hmRow = dgvSaiKhac.Rows[hmIdx];
                    hmRow.Tag = "HANG_MUC_HEADER";
                    hmRow.MinimumHeight = 34;
                    hmRow.DefaultCellStyle.BackColor = Color.FromArgb(220, 235, 252);
                    hmRow.DefaultCellStyle.Font = UIHelper.GetFont(10f, FontStyle.Bold);
                    hmRow.Cells["colTen"].Style.ForeColor = Color.FromArgb(0, 51, 102);
                    hmRow.Cells["colKhoiLuong"].ReadOnly = true;
                    hmRow.Cells["colLuaChon"].ReadOnly = true;
                }

                // 1. DÒNG CÔNG TÁC (PARENT / GROUP HEADER ROW)
                string prefixToggle = g.DanhSachHaoPhi.Count > 0 ? (g.IsExpanded ? "▼ " : "▶ ") : "";
                string summarySaiKhac = g.IsTamTinh
                    ? "Công tác tạm tính / Không có trong TT38"
                    : (g.DanhSachHaoPhi.Count > 0
                        ? $"📌 {g.DanhSachHaoPhi.Count} sai khác ({g.DanhSachHaoPhi.Count(x => x.LoaiHP == LoaiHaoPhi.VL)} VL, {g.DanhSachHaoPhi.Count(x => x.LoaiHP == LoaiHaoPhi.NC)} NC, {g.DanhSachHaoPhi.Count(x => x.LoaiHP == LoaiHaoPhi.MAY)} Máy)" + (g.IsExpanded ? "" : "  [Click để xem]")
                        : "✓ Chuẩn TT38 (Định mức & Đơn giá khớp)");

                decimal dgDmDT = g.DonGiaDT;
                decimal dgDmTD = g.DonGiaTDHienThi;

                int pIdx = dgvSaiKhac.Rows.Add(
                    prefixToggle + (i + 1),
                    g.MaCongTac,
                    g.TenCongTac,
                    "[CÔNG TÁC]",
                    g.DonVi,
                    g.KhoiLuong,
                    summarySaiKhac,
                    "", // ĐM DT trống ở dòng công tác
                    "", // ĐM TĐ trống ở dòng công tác
                    "", // Đơn giá VT DT trống ở dòng công tác
                    "", // Đơn giá VT TĐ trống ở dòng công tác
                    dgDmDT > 0 ? (object)dgDmDT : "",
                    dgDmTD > 0 ? (object)dgDmTD : "",
                    g.ThanhTienDT,
                    g.ThanhTienTD,
                    g.ChenhLech,
                    g.LuaChon
                );

                var pRow = dgvSaiKhac.Rows[pIdx];
                pRow.Tag = g;
                pRow.MinimumHeight = 36;
                pRow.DefaultCellStyle.BackColor = Color.FromArgb(232, 241, 252);
                pRow.DefaultCellStyle.Font = UIHelper.GetFont(9.5f, FontStyle.Bold);
                pRow.Cells["colMa"].Style.ForeColor = Color.FromArgb(0, 51, 102);
                pRow.Cells["colTen"].Style.ForeColor = Color.FromArgb(0, 51, 102);
                pRow.Cells["colKhoiLuong"].ReadOnly = false; // Dòng công tác cho phép sửa khối lượng
                pRow.Cells["colKhoiLuong"].Style.BackColor = Color.FromArgb(255, 255, 230); // Màu vàng nhạt nhắc nhở có thể sửa

                CapNhatMauLuaChon(pRow, g.LuaChon);

                // 2. CÁC DÒNG HAO PHÍ CON (CHỈ HIỂN THỊ KHI ĐƯỢC MỞ RỘNG)
                if (g.IsExpanded)
                {
                    for (int c = 0; c < g.DanhSachHaoPhi.Count; c++)
                    {
                        var h = g.DanhSachHaoPhi[c];
                        int cIdx = dgvSaiKhac.Rows.Add(
                            "", // STT để trống
                            h.MaHaoPhi,
                            "   ↳ " + h.TenHaoPhi,
                            h.LoaiHP?.ToString() ?? "",
                            h.DonVi,
                            h.KhoiLuongHaoPhiTD > 0 ? h.KhoiLuongHaoPhiTD : h.KhoiLuongHaoPhiDT,
                            h.LoaiSaiKhac,
                            h.DinhMucDT > 0 ? (object)h.DinhMucDT : "",
                            h.DinhMucTD > 0 ? (object)h.DinhMucTD : "",
                            h.DonGiaDT > 0 ? (object)h.DonGiaDT : "",
                            h.DonGiaTD > 0 ? (object)h.DonGiaTD : "",
                            h.DonGiaTheoDinhMucDT > 0 ? (object)h.DonGiaTheoDinhMucDT : "",
                            h.DonGiaTheoDinhMucTD > 0 ? (object)h.DonGiaTheoDinhMucTD : "",
                            h.ThanhTienDT,
                            h.ThanhTienTD,
                            h.ChenhLech,
                            h.LuaChon
                        );

                        var cRow = dgvSaiKhac.Rows[cIdx];
                        cRow.Tag = h;
                        cRow.MinimumHeight = 30;
                        cRow.DefaultCellStyle.Font = UIHelper.GetFont(9f, FontStyle.Regular);
                        cRow.Cells["colKhoiLuong"].ReadOnly = true; // Dòng hao phí tính theo công thức, không sửa trực tiếp
                        CapNhatMauLuaChon(cRow, h.LuaChon);
                    }
                }
            }

            // Tự động giãn chiều cao hàng theo nội dung text nhiều dòng & co giãn cột tên công tác
            dgvSaiKhac.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders);
            CapNhatDoRongCotTen();

            int countTD = _allGroups.Count(x => x.LuaChon == "Theo Thẩm định");
            int countDT = _allGroups.Count(x => x.LuaChon == "Theo Dự toán");
            lblCountInfo.Text = $"Hiển thị: {_displayGroups.Count}/{_allGroups.Count} công tác có sai khác   |   Theo Thẩm định: {countTD}   |   Theo Dự toán: {countDT}";
        }
        finally
        {
            _isUpdatingGrid = false;
        }
    }

    private void CapNhatMauLuaChon(DataGridViewRow row, string luaChon)
    {
        if (row == null) return;

        bool isGroup = row.Tag is CongTacSaiKhacGroup;

        if (luaChon == "Theo Dự toán")
        {
            row.DefaultCellStyle.BackColor = isGroup ? Color.FromArgb(254, 245, 225) : Color.FromArgb(255, 250, 238);
            row.Cells["colLuaChon"].Style.ForeColor = Color.FromArgb(180, 90, 0);
            row.Cells["colLuaChon"].Style.BackColor = Color.FromArgb(255, 240, 210);
        }
        else if (luaChon == "Tùy biến")
        {
            row.DefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
            row.Cells["colLuaChon"].Style.ForeColor = Color.FromArgb(120, 60, 160);
            row.Cells["colLuaChon"].Style.BackColor = Color.FromArgb(240, 230, 250);
        }
        else
        {
            row.DefaultCellStyle.BackColor = isGroup ? Color.FromArgb(232, 241, 252) : Color.White;
            row.Cells["colLuaChon"].Style.ForeColor = Color.FromArgb(0, 102, 204);
            row.Cells["colLuaChon"].Style.BackColor = Color.FromArgb(235, 245, 255);
        }
    }

    private void DgvSaiKhac_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
    {
        GridHelper.PaintMultiMergedHeaders(sender, e, dgvSaiKhac,
            new GridHelper.MergedHeaderGroup(7, 8, "Định mức"),
            new GridHelper.MergedHeaderGroup(9, 10, "Đơn giá vật tư (đ)"),
            new GridHelper.MergedHeaderGroup(11, 12, "ĐG trong ĐM (đ)"),
            new GridHelper.MergedHeaderGroup(13, 14, "Thành tiền (đ)")
        );
    }

    private void DgvSaiKhac_CellClick(object sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
        string colName = dgvSaiKhac.Columns[e.ColumnIndex].Name;

        // Không mở rộng/thu gọn nếu người dùng thao tác ở cột sửa Khối lượng hoặc Lựa chọn áp dụng
        if (colName == "colKhoiLuong" || colName == "colLuaChon") return;

        var row = dgvSaiKhac.Rows[e.RowIndex];
        if (row.Tag is CongTacSaiKhacGroup group && group.DanhSachHaoPhi.Count > 0)
        {
            group.IsExpanded = !group.IsExpanded;
            NapDuLieuLenGrid();

            // Khôi phục con trỏ về dòng công tác vừa click
            for (int r = 0; r < dgvSaiKhac.Rows.Count; r++)
            {
                if (dgvSaiKhac.Rows[r].Tag == group)
                {
                    dgvSaiKhac.CurrentCell = dgvSaiKhac.Rows[r].Cells[e.ColumnIndex];
                    break;
                }
            }
        }
    }

    private void DgvSaiKhac_CellValueChanged(object sender, DataGridViewCellEventArgs e)
    {
        if (_isUpdatingGrid || e.RowIndex < 0 || e.ColumnIndex < 0) return;

        var row = dgvSaiKhac.Rows[e.RowIndex];
        string colName = dgvSaiKhac.Columns[e.ColumnIndex].Name;

        // 1. NGƯỜI DÙNG CHỈNH SỬA KHỐI LƯỢNG TOÀN BỘ TRÊN DÒNG CÔNG TÁC
        if (colName == "colKhoiLuong")
        {
            if (row.Tag is CongTacSaiKhacGroup group)
            {
                object val = row.Cells["colKhoiLuong"].Value;
                if (val != null && decimal.TryParse(val.ToString().Replace(" ", "").Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal newKl) && newKl > 0)
                {
                    group.KhoiLuong = newKl;
                    // Cập nhật lại toàn bộ grid và tổng tiền
                    NapDuLieuLenGrid();
                    CapNhatTongHopChiPhiTrucTiep();
                }
            }
        }
        // 2. NGƯỜI DÙNG THAY ĐỔI LỰA CHỌN ÁP DỤNG (THEO THẨM ĐỊNH / THEO DỰ TOÁN)
        else if (colName == "colLuaChon")
        {
            string val = row.Cells["colLuaChon"].Value?.ToString() ?? "Theo Thẩm định";

            // Nếu người dùng chọn trên dòng CÔNG TÁC (Parent): Áp dụng đồng loạt cho tất cả các hao phí con!
            if (row.Tag is CongTacSaiKhacGroup group)
            {
                if (val == "Theo Thẩm định" || val == "Theo Dự toán")
                {
                    group.LuaChon = val;
                    foreach (var h in group.DanhSachHaoPhi)
                    {
                        h.LuaChon = val;
                    }
                    NapDuLieuLenGrid();
                    CapNhatTongHopChiPhiTrucTiep();
                }
            }
            // Nếu người dùng chọn trên dòng HAO PHÍ CON (Child)
            else if (row.Tag is HaoPhiSaiKhacItem child)
            {
                child.LuaChon = val;
                var p = child.ParentGroup;
                if (p != null)
                {
                    int tdCount = p.DanhSachHaoPhi.Count(x => x.LuaChon == "Theo Thẩm định");
                    int dtCount = p.DanhSachHaoPhi.Count(x => x.LuaChon == "Theo Dự toán");

                    if (tdCount == p.DanhSachHaoPhi.Count) p.LuaChon = "Theo Thẩm định";
                    else if (dtCount == p.DanhSachHaoPhi.Count) p.LuaChon = "Theo Dự toán";
                    else p.LuaChon = "Tùy biến";
                }

                CapNhatMauLuaChon(row, val);
                CapNhatTongHopChiPhiTrucTiep();
            }
        }
    }

    private void ApDungTatCa(bool theoThamDinh)
    {
        string target = theoThamDinh ? "Theo Thẩm định" : "Theo Dự toán";

        foreach (var g in _allGroups)
        {
            g.LuaChon = target;
            foreach (var h in g.DanhSachHaoPhi)
            {
                h.LuaChon = target;
            }
        }

        NapDuLieuLenGrid();
        CapNhatTongHopChiPhiTrucTiep();
    }

    /// <summary>
    /// Tính toán Chi phí trực tiếp thẩm định toàn bộ công trình
    /// Duyệt qua từng công tác độc lập của từng Hạng mục (tuyệt đối không cộng dồn khối lượng)
    /// Đảm bảo tính toán chính xác 100% khớp từng đồng với cột Thành tiền Thẩm định và Thành tiền Dự toán trên bảng!
    /// </summary>
    public void TinhChiPhiTrucTiepToanBo(
        out decimal vlTD, out decimal ncTD, out decimal mayTD,
        out decimal vlDT, out decimal ncDT, out decimal mayDT)
    {
        vlTD = 0m; ncTD = 0m; mayTD = 0m;
        vlDT = 0m; ncDT = 0m; mayDT = 0m;

        foreach (var grp in _allGroups)
        {
            decimal kl = grp.KhoiLuong;
            var kq = grp.KetQuaCongTac;
            decimal ttItemDT = grp.ThanhTienDT; // Luôn bằng Round(KhoiLuong * DonGiaDT, 0)
            decimal ttItemTD = grp.ThanhTienTD; // Luôn bằng ThanhTienDT + ChenhLech

            decimal itemVlDT = 0m, itemNcDT = 0m, itemMayDT = 0m;

            // 1. Phân bổ thành phần VL, NC, Máy cho Dự toán (itemVlDT + itemNcDT + itemMayDT = ttItemDT)
            if (kq?.DuToan?.DanhSachHaoPhi != null && kq.DuToan.DanhSachHaoPhi.Count > 0)
            {
                decimal rawVl = 0m, rawNc = 0m, rawMay = 0m;
                foreach (var hp in kq.DuToan.DanhSachHaoPhi)
                {
                    decimal tt = Math.Round(kl * hp.DinhMuc * hp.DonGia, 0);
                    if (hp.Loai == LoaiHaoPhi.VL) rawVl += tt;
                    else if (hp.Loai == LoaiHaoPhi.NC) rawNc += tt;
                    else if (hp.Loai == LoaiHaoPhi.MAY) rawMay += tt;
                }
                decimal rawTotal = rawVl + rawNc + rawMay;
                if (rawTotal > 0)
                {
                    itemVlDT = Math.Round(ttItemDT * (rawVl / rawTotal), 0);
                    itemNcDT = Math.Round(ttItemDT * (rawNc / rawTotal), 0);
                    itemMayDT = ttItemDT - itemVlDT - itemNcDT;
                }
                else
                {
                    itemVlDT = Math.Round(ttItemDT * 0.60m, 0);
                    itemNcDT = Math.Round(ttItemDT * 0.25m, 0);
                    itemMayDT = ttItemDT - itemVlDT - itemNcDT;
                }
            }
            else if (kq?.DanhSachSaiLech != null && kq.DanhSachSaiLech.Count > 0)
            {
                decimal rawVl = 0m, rawNc = 0m, rawMay = 0m;
                foreach (var s in kq.DanhSachSaiLech)
                {
                    decimal dm = s.HaoPhiChuan?.DinhMuc ?? s.HaoPhiDuToan?.DinhMuc ?? 0m;
                    decimal dg = s.DonGiaChuan ?? s.HaoPhiDuToan?.DonGia ?? 1000m;
                    decimal tt = Math.Round(kl * dm * dg, 0);
                    LoaiHaoPhi loai = s.LoaiHP ?? s.HaoPhiChuan?.LoaiHaoPhi ?? s.HaoPhiDuToan?.Loai ?? LoaiHaoPhi.VL;
                    if (loai == LoaiHaoPhi.VL) rawVl += tt;
                    else if (loai == LoaiHaoPhi.NC) rawNc += tt;
                    else if (loai == LoaiHaoPhi.MAY) rawMay += tt;
                }
                decimal rawTotal = rawVl + rawNc + rawMay;
                if (rawTotal > 0)
                {
                    itemVlDT = Math.Round(ttItemDT * (rawVl / rawTotal), 0);
                    itemNcDT = Math.Round(ttItemDT * (rawNc / rawTotal), 0);
                    itemMayDT = ttItemDT - itemVlDT - itemNcDT;
                }
                else
                {
                    itemVlDT = Math.Round(ttItemDT * 0.60m, 0);
                    itemNcDT = Math.Round(ttItemDT * 0.25m, 0);
                    itemMayDT = ttItemDT - itemVlDT - itemNcDT;
                }
            }
            else
            {
                // Công tác tạm tính hoặc không có định mức
                itemVlDT = Math.Round(ttItemDT * 0.60m, 0);
                itemNcDT = Math.Round(ttItemDT * 0.25m, 0);
                itemMayDT = ttItemDT - itemVlDT - itemNcDT;
            }

            vlDT += itemVlDT;
            ncDT += itemNcDT;
            mayDT += itemMayDT;

            // 2. Tính thành phần VL, NC, Máy cho Thẩm định (itemVlTD + itemNcTD + itemMayTD = ttItemTD)
            if (grp.LuaChon == "Theo Dự toán")
            {
                vlTD += itemVlDT;
                ncTD += itemNcDT;
                mayTD += itemMayDT;
            }
            else if (grp.DanhSachHaoPhi.Count > 0)
            {
                // Cộng chênh lệch thẩm định từng thành phần của các hao phí có áp dụng thẩm định
                decimal clVl = grp.DanhSachHaoPhi.Where(h => h.LoaiHP == LoaiHaoPhi.VL && h.LuaChon != "Theo Dự toán").Sum(h => h.ChenhLech);
                decimal clNc = grp.DanhSachHaoPhi.Where(h => h.LoaiHP == LoaiHaoPhi.NC && h.LuaChon != "Theo Dự toán").Sum(h => h.ChenhLech);
                decimal clMay = grp.DanhSachHaoPhi.Where(h => h.LoaiHP == LoaiHaoPhi.MAY && h.LuaChon != "Theo Dự toán").Sum(h => h.ChenhLech);

                decimal itemVlTD = itemVlDT + clVl;
                decimal itemNcTD = itemNcDT + clNc;
                decimal itemMayTD = itemMayDT + clMay;

                // Khử sai số làm tròn nếu có để itemVlTD + itemNcTD + itemMayTD == ttItemTD
                decimal diff = ttItemTD - (itemVlTD + itemNcTD + itemMayTD);
                if (diff != 0) itemMayTD += diff;

                vlTD += itemVlTD;
                ncTD += itemNcTD;
                mayTD += itemMayTD;
            }
            else
            {
                // Công tác không có chi tiết hao phí hoặc tạm tính: phân bổ theo tỷ lệ
                decimal itemVlTD = Math.Round(ttItemTD * 0.60m, 0);
                decimal itemNcTD = Math.Round(ttItemTD * 0.25m, 0);
                decimal itemMayTD = ttItemTD - itemVlTD - itemNcTD;

                vlTD += itemVlTD;
                ncTD += itemNcTD;
                mayTD += itemMayTD;
            }
        }
    }

    private void CapNhatTongHopChiPhiTrucTiep()
    {
        TinhChiPhiTrucTiepToanBo(
            out decimal vlTD, out decimal ncTD, out decimal mayTD,
            out decimal vlDT, out decimal ncDT, out decimal mayDT);

        decimal tongTTD = vlTD + ncTD + mayTD;
        decimal tongTDT = vlDT + ncDT + mayDT;
        decimal chenhLech = tongTTD - tongTDT;

        lblTongVL.Text = UIHelper.FormatTien(vlTD) + " đ";
        lblTongNC.Text = UIHelper.FormatTien(ncTD) + " đ";
        lblTongMay.Text = UIHelper.FormatTien(mayTD) + " đ";
        lblTongT.Text = UIHelper.FormatTien(tongTTD) + " đ";

        string clStr = (chenhLech >= 0 ? "+" : "") + UIHelper.FormatTien(chenhLech) + " đ";
        lblChenhLechSoVoiDT.Text = $"So với Dự toán gốc: {clStr}";
        lblChenhLechSoVoiDT.ForeColor = chenhLech < 0 ? Color.FromArgb(0, 128, 64) : (chenhLech > 0 ? Color.FromArgb(180, 50, 50) : Color.FromArgb(70, 70, 70));
    }

    private void BtnTiepTuc_Click(object sender, EventArgs e)
    {
        try
        {
            // 1. Tổng hợp từ điển Đơn giá Thẩm định chuẩn (cho 1 đơn vị công tác)
            var donGiaMap = new Dictionary<string, DonGiaThamDinhItem>(StringComparer.OrdinalIgnoreCase);
            var mapGroups = _allGroups
                .Where(g => g.KetQuaCongTac != null)
                .GroupBy(g => g.KetQuaCongTac)
                .ToDictionary(g => g.Key, g => g.First());

            foreach (var kq in _results)
            {
                string cleanMa = kq.DuToan.MaHieu.Trim().ToUpper();
                CongTacSaiKhacGroup? grp = null;
                mapGroups.TryGetValue(kq, out grp);

                var item = new DonGiaThamDinhItem
                {
                    MaHieu = kq.DuToan.MaHieu,
                    TenCongTac = kq.DuToan.TenCongTac,
                    DonVi = kq.DuToan.DonVi,
                    LuaChon = grp?.LuaChon ?? "Theo Thẩm định",
                    CoSaiKhac = grp != null
                };

                // Tính toán đơn giá cho 1 đơn vị công tác
                if (kq.DinhMucChuan == null)
                {
                    // Công tác tạm tính
                    decimal dg = kq.DuToan.DonGia > 0 ? kq.DuToan.DonGia : (kq.DuToan.KhoiLuong > 0 ? Math.Round(kq.DuToan.ThanhTien / kq.DuToan.KhoiLuong, 0) : kq.DuToan.ThanhTien);
                    item.DonGiaVL_DT = Math.Round(dg * 0.60m, 0);
                    item.DonGiaNC_DT = Math.Round(dg * 0.25m, 0);
                    item.DonGiaMay_DT = dg - item.DonGiaVL_DT - item.DonGiaNC_DT;

                    item.DonGiaVL_TD = item.DonGiaVL_DT;
                    item.DonGiaNC_TD = item.DonGiaNC_DT;
                    item.DonGiaMay_TD = item.DonGiaMay_DT;
                }
                else
                {
                    // 1. Đơn giá Dự toán gốc cho 1 đơn vị
                    decimal vlDt = 0m, ncDt = 0m, mayDt = 0m;
                    foreach (var hp in kq.DuToan.DanhSachHaoPhi)
                    {
                        decimal tt = hp.DinhMuc * hp.DonGia;
                        if (hp.Loai == LoaiHaoPhi.VL) vlDt += tt;
                        else if (hp.Loai == LoaiHaoPhi.NC) ncDt += tt;
                        else if (hp.Loai == LoaiHaoPhi.MAY) mayDt += tt;
                    }
                    item.DonGiaVL_DT = vlDt;
                    item.DonGiaNC_DT = ncDt;
                    item.DonGiaMay_DT = mayDt;

                    // 2. Đơn giá Thẩm định chuẩn cho 1 đơn vị
                    decimal vlTd = 0m, ncTd = 0m, mayTd = 0m;
                    var childMap = grp?.DanhSachHaoPhi.Where(x => x.SaiLech != null).ToDictionary(x => x.SaiLech!, x => x);

                    foreach (var s in kq.DanhSachSaiLech)
                    {
                        LoaiHaoPhi loai = s.LoaiHP ?? s.HaoPhiChuan?.LoaiHaoPhi ?? s.HaoPhiDuToan?.Loai ?? LoaiHaoPhi.VL;

                        if (childMap != null && childMap.TryGetValue(s, out var childItem))
                        {
                            if (childItem.LuaChon == "Theo Dự toán")
                            {
                                if (s.HaoPhiDuToan != null)
                                {
                                    decimal tt = s.HaoPhiDuToan.DinhMuc * s.HaoPhiDuToan.DonGia;
                                    if (loai == LoaiHaoPhi.VL) vlTd += tt;
                                    else if (loai == LoaiHaoPhi.NC) ncTd += tt;
                                    else if (loai == LoaiHaoPhi.MAY) mayTd += tt;
                                }
                            }
                            else
                            {
                                if (s.HaoPhiChuan != null)
                                {
                                    decimal dg = s.DonGiaChuan ?? s.HaoPhiDuToan?.DonGia ?? 0m;
                                    decimal tt = s.HaoPhiChuan.DinhMuc * dg;
                                    if (loai == LoaiHaoPhi.VL) vlTd += tt;
                                    else if (loai == LoaiHaoPhi.NC) ncTd += tt;
                                    else if (loai == LoaiHaoPhi.MAY) mayTd += tt;
                                }
                            }
                        }
                        else
                        {
                            decimal dm = s.HaoPhiChuan?.DinhMuc ?? s.HaoPhiDuToan?.DinhMuc ?? 0m;
                            decimal dg = s.DonGiaChuan ?? s.HaoPhiDuToan?.DonGia ?? 0m;
                            decimal tt = dm * dg;
                            if (loai == LoaiHaoPhi.VL) vlTd += tt;
                            else if (loai == LoaiHaoPhi.NC) ncTd += tt;
                            else if (loai == LoaiHaoPhi.MAY) mayTd += tt;
                        }
                    }

                    if (kq.DanhSachSaiLech.Count == 0)
                    {
                        vlTd = vlDt;
                        ncTd = ncDt;
                        mayTd = mayDt;
                    }

                    item.DonGiaVL_TD = vlTd;
                    item.DonGiaNC_TD = ncTd;
                    item.DonGiaMay_TD = mayTd;
                }

                donGiaMap[cleanMa] = item;
            }

            // 2. Chuyển sang Form Giai đoạn 2: Bảng Dự Toán Chi Tiết Thẩm Định (theo từng Hạng mục)
            var formChiTiet = new ThamDinhDuToanChiTietForm(_tenCongTrinh, _boDonGiaId, donGiaMap, _results, _selectedSheetName);
            formChiTiet.Show();
        }
        catch (Exception ex)
        {
            MessageBox.Show("Lỗi khi chuyển sang Bảng Dự Toán Chi Tiết: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private string NhanDienLoaiCongTrinhTuDong(string tenDuAn)
    {
        if (string.IsNullOrWhiteSpace(tenDuAn)) return "Công trình Dân dụng";

        string lower = tenDuAn.ToLower();

        // 1. Nông nghiệp & Môi trường
        string[] nnKw = { "mương", "kênh", "hồ chứa", "đập", "trạm bơm", "thủy lợi", "đê", "kè", "tưới tiêu", "nông nghiệp" };
        if (nnKw.Any(k => lower.Contains(k))) return "Công trình Nông nghiệp và Môi trường";

        // 2. Giao thông
        string[] gtKw = { "đường", "cầu", "tuyến", "giao thông", "hè phố", "ngã tư", "nút giao", "cống ngang", "bến" };
        if (gtKw.Any(k => lower.Contains(k))) return "Công trình Giao thông";

        // 3. Hạ tầng kỹ thuật
        string[] htKw = { "hạ tầng", "thoát nước", "cấp nước", "chiếu sáng", "cây xanh", "công viên", "xử lý nước thải", "rác thải", "nghĩa trang" };
        if (htKw.Any(k => lower.Contains(k))) return "Công trình Hạ tầng kỹ thuật";

        // 4. Công nghiệp
        string[] cnKw = { "nhà xưởng", "nhà máy", "công nghiệp", "khu chế xuất", "trạm biến áp", "đường dây", "kho xưởng", "lò" };
        if (cnKw.Any(k => lower.Contains(k))) return "Công trình Công nghiệp";

        return "Công trình Dân dụng";
    }

    private void CapNhatDoRongCotTen()
    {
        if (dgvSaiKhac?.Columns["colTen"] == null) return;
        int fixedWidth = 0;
        foreach (DataGridViewColumn col in dgvSaiKhac.Columns)
        {
            if (col.Name != "colTen" && col.Visible)
                fixedWidth += col.Width;
        }

        int available = dgvSaiKhac.ClientSize.Width - fixedWidth - SystemInformation.VerticalScrollBarWidth;
        if (available > 280)
        {
            dgvSaiKhac.Columns["colTen"].Width = available;
        }
        else
        {
            dgvSaiKhac.Columns["colTen"].Width = 280;
        }
    }

    private void BtnChiDinhKL_Click(object sender, EventArgs e)
    {
        try
        {
            using (var dlg = new ChonCotKhoiLuongDialog(_allGroups, _results, _boDonGiaId))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    CapNhatDanhSachCongTacTheoTungHangMuc(dlg.SelectedMappingItems, dlg.SelectedSheetName);
                    NapDuLieuLenGrid();
                    CapNhatTongHopChiPhiTrucTiep();

                    int soHM = _allGroups.Where(x => !string.IsNullOrEmpty(x.TenHangMuc)).Select(x => x.TenHangMuc).Distinct().Count();
                    MessageBox.Show(
                        $"Đã cập nhật thành công {dlg.AppliedCount} công tác độc lập thuộc {soHM} hạng mục từ sheet '{dlg.SelectedSheetName}'!\n\nĐịnh mức và đơn giá thẩm định đã được gán chính xác cho từng dòng mã hiệu của từng hạng mục.",
                        "Thành công",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Lỗi khi chỉ định cột khối lượng: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void BtnXoaCongTac_Click(object sender, EventArgs e)
    {
        XoaCongTacDuocChon();
    }

    private void XoaCongTacDuocChon()
    {
        try
        {
            if (dgvSaiKhac.CurrentRow == null)
            {
                MessageBox.Show("Vui lòng chọn công tác hoặc hao phí cần xóa trong bảng!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            CongTacSaiKhacGroup? targetGroup = null;
            if (dgvSaiKhac.CurrentRow.Tag is CongTacSaiKhacGroup grp)
            {
                targetGroup = grp;
            }
            else if (dgvSaiKhac.CurrentRow.Tag is HaoPhiSaiKhacItem hp && hp.ParentGroup != null)
            {
                targetGroup = hp.ParentGroup;
            }

            if (targetGroup == null)
            {
                MessageBox.Show("Không tìm thấy thông tin công tác tương ứng với dòng đang chọn!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var dr = MessageBox.Show(
                $"Bạn có chắc chắn muốn xóa công tác [{targetGroup.MaCongTac}] \"{targetGroup.TenCongTac}\" khỏi danh sách thẩm định không?\n\n(Lưu ý: Công tác này sẽ bị loại trừ khỏi bảng thẩm định và Bảng Tổng hợp kinh phí)",
                "Xác nhận xóa công tác",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (dr == DialogResult.Yes)
            {
                _allGroups.Remove(targetGroup);
                if (targetGroup.KetQuaCongTac != null)
                {
                    _results.Remove(targetGroup.KetQuaCongTac);
                }

                // Đánh lại số thứ tự STT
                for (int i = 0; i < _allGroups.Count; i++)
                {
                    _allGroups[i].STT = i + 1;
                }

                // Cập nhật lại nhãn thông tin dự án ở header
                int tongSoLoi = _allGroups.Sum(g => g.DanhSachHaoPhi.Count > 0 ? g.DanhSachHaoPhi.Count : 1);
                lblThongTin.Text = $"Dự án: {_tenCongTrinh}   |   Tổng số công tác: {_results.Count}   |   Số công tác có sai khác: {_allGroups.Count}   |   Tổng mục sai khác: {tongSoLoi}";

                NapDuLieuLenGrid();
                CapNhatTongHopChiPhiTrucTiep();

                MessageBox.Show($"Đã xóa thành công công tác [{targetGroup.MaCongTac}]!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Lỗi khi xóa công tác: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static string BoDau(string text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        string normalized = text.Normalize(System.Text.NormalizationForm.FormD);
        var sb = new System.Text.StringBuilder();
        foreach (char c in normalized)
        {
            var uc = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c);
            if (uc != System.Globalization.UnicodeCategory.NonSpacingMark)
                sb.Append(c);
        }
        return sb.ToString().Normalize(System.Text.NormalizationForm.FormC).ToLower().Trim();
    }

    /// <summary>
    /// Xuất danh sách đối chiếu chênh lệch Dự toán - Thẩm định sang một Sheet Excel chuyên dụng.
    /// Bao gồm đầy đủ khối lượng, đơn giá, thành tiền, mức chênh lệch và phương án điều chỉnh.
    /// </summary>
    private void XuatBangDoiChieuChenhLechExcel()
    {
        try
        {
            var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            var wb = app?.ActiveWorkbook;
            if (wb == null)
            {
                MessageBox.Show("Không tìm thấy file Excel đang mở!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Tạo hoặc lấy sheet DoiChieu_SaiKhac
            ExcelWorksheet? ws = null;
            foreach (ExcelWorksheet s in wb.Sheets)
            {
                if (string.Equals(s.Name, "DoiChieu_SaiKhac", StringComparison.OrdinalIgnoreCase))
                {
                    ws = s;
                    break;
                }
            }
            if (ws == null)
            {
                ws = (ExcelWorksheet)wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
                ws.Name = "DoiChieu_SaiKhac";
            }
            else
            {
                ws.Cells.Clear();
            }

            ws.Cells.Font.Name = "Times New Roman";
            ws.Cells.Font.Size = 11;

            // Tiêu đề
            ws.Cells[1, 1] = "BẢNG ĐỐI CHIẾU CHÊNH LỆCH DỰ TOÁN VÀ THẨM ĐỊNH";
            var rngTitle = ws.Range["A1:N1"];
            rngTitle.Merge();
            rngTitle.Font.Bold = true;
            rngTitle.Font.Size = 14;
            rngTitle.HorizontalAlignment = ExcelHAlign.xlHAlignCenter;

            // Thông tin công trình
            ws.Cells[2, 1] = $"Công trình: {_tenCongTrinh} | Sheet gốc: {_selectedSheetName} | Ngày đối chiếu: {DateTime.Now:dd/MM/yyyy HH:mm}";
            var rngSub = ws.Range["A2:N2"];
            rngSub.Merge();
            rngSub.Font.Italic = true;
            rngSub.Font.Size = 10;
            rngSub.HorizontalAlignment = ExcelHAlign.xlHAlignCenter;

            // Headers
            string[] headers = new string[]
            {
                "STT", "Hạng mục", "Mã hiệu", "Nội dung công việc / Hao phí", "ĐVT",
                "Khối lượng", "Đơn giá DT", "Đơn giá TĐ", "Thành tiền DT", "Thành tiền TĐ",
                "Chênh lệch (đ)", "Tỷ lệ (%)", "Phương án chọn", "Ghi chú sai khác"
            };

            for (int col = 0; col < headers.Length; col++)
            {
                ws.Cells[4, col + 1] = headers[col];
            }

            var rngHeader = ws.Range[ws.Cells[4, 1], ws.Cells[4, headers.Length]];
            rngHeader.Font.Bold = true;
            rngHeader.HorizontalAlignment = ExcelHAlign.xlHAlignCenter;
            rngHeader.VerticalAlignment = ExcelVAlign.xlVAlignCenter;
            rngHeader.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(220, 230, 242));
            rngHeader.RowHeight = 28;

            int row = 5;
            int stt = 1;
            var groupsToExport = _displayGroups.Count > 0 ? _displayGroups : _allGroups;

            foreach (var g in groupsToExport)
            {
                ws.Cells[row, 1] = stt++;
                ws.Cells[row, 2] = g.TenHangMuc;
                ws.Cells[row, 3] = g.MaCongTac;
                ws.Cells[row, 4] = g.TenCongTac;
                ws.Cells[row, 5] = g.DonVi;
                ws.Cells[row, 6] = (double)g.KhoiLuong;
                ws.Cells[row, 7] = (double)(g.DonVi.Length > 0 ? g.DonGiaDT : 0);
                ws.Cells[row, 8] = (double)(g.DonVi.Length > 0 ? g.DonGiaTDHienThi : 0);
                ws.Cells[row, 9] = (double)g.ThanhTienDT;
                ws.Cells[row, 10] = (double)g.ThanhTienTD;
                ws.Cells[row, 11] = (double)g.ChenhLech;
                ws.Cells[row, 12] = g.ThanhTienDT != 0 ? (double)Math.Round((g.ChenhLech / g.ThanhTienDT) * 100m, 2) : 0;
                ws.Cells[row, 13] = g.LuaChon;
                ws.Cells[row, 14] = g.DanhSachHaoPhi.Count > 0 ? $"{g.DanhSachHaoPhi.Count} mục sai khác" : "Sai lệch đơn giá";

                var rowRng = ws.Range[ws.Cells[row, 1], ws.Cells[row, headers.Length]];
                rowRng.Font.Bold = true;
                rowRng.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(245, 248, 252));
                row++;

                // Xuất từng hao phí chi tiết của công tác
                foreach (var hp in g.DanhSachHaoPhi)
                {
                    ws.Cells[row, 1] = "";
                    ws.Cells[row, 2] = "";
                    ws.Cells[row, 3] = hp.MaHaoPhi;
                    ws.Cells[row, 4] = "  - " + hp.TenHaoPhi;
                    ws.Cells[row, 5] = hp.DonVi;
                    ws.Cells[row, 6] = (double)hp.KhoiLuongHaoPhiTD;
                    ws.Cells[row, 7] = (double)hp.DonGiaDT;
                    ws.Cells[row, 8] = (double)hp.DonGiaTD;
                    ws.Cells[row, 9] = (double)hp.ThanhTienDT;
                    ws.Cells[row, 10] = (double)hp.ThanhTienTD;
                    ws.Cells[row, 11] = (double)hp.ChenhLech;
                    ws.Cells[row, 12] = hp.ThanhTienDT != 0 ? (double)Math.Round((hp.ChenhLech / hp.ThanhTienDT) * 100m, 2) : 0;
                    ws.Cells[row, 13] = hp.LuaChon;
                    ws.Cells[row, 14] = hp.LoaiSaiKhac;
                    row++;
                }
            }

            // Dòng Tổng cộng
            ws.Cells[row, 4] = "TỔNG CỘNG CHÊNH LỆCH";
            ws.Cells[row, 4].Font.Bold = true;
            ws.Cells[row, 9].Formula = $"=SUM(I5:I{row - 1})";
            ws.Cells[row, 10].Formula = $"=SUM(J5:J{row - 1})";
            ws.Cells[row, 11].Formula = $"=J{row}-I{row}";
            var totalRange = ws.Range[ws.Cells[row, 1], ws.Cells[row, headers.Length]];
            totalRange.Font.Bold = true;
            totalRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(254, 243, 199));

            // Định dạng số
            ExcelFormatHelper.ApplyRateFormat(ws.Range[$"F5:F{row}"], 3);
            ExcelFormatHelper.ApplyIntegerFormat(ws.Range[$"G5:K{row}"]);
            ExcelFormatHelper.ApplyRateFormat(ws.Range[$"L5:L{row}"], 2);

            // Kẻ bảng
            var tableRange = ws.Range[ws.Cells[4, 1], ws.Cells[row, headers.Length]];
            tableRange.Borders.LineStyle = ExcelLineStyle.xlContinuous;
            tableRange.Borders.Weight = ExcelBorderWeight.xlThin;

            ws.Columns.AutoFit();
            XuatBangBieuService.ThietLapTrangInA4(ws, Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape, "$4:$4");
            ws.Activate();

            MessageBox.Show($"Đã xuất thành công {groupsToExport.Count} công tác đối chiếu sang sheet 'DoiChieu_SaiKhac'!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi khi xuất bảng đối chiếu chênh lệch: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
