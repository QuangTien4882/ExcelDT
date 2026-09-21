using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AIE.Core.Models;
using AIE.ExcelAddIn.Helpers;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Button = System.Windows.Forms.Button;
using Label = System.Windows.Forms.Label;
using Point = System.Drawing.Point;

namespace AIE.ExcelAddIn.Forms;

/// <summary>
/// Đại diện cho một dòng ánh xạ khối lượng giữa Thẩm định và Excel
/// </summary>
public class KhoiLuongMappingItem
{
    public bool Selected { get; set; } = true;
    public KetQuaCongTacThamDinh? MatchedResult { get; set; }
    public CongTacSaiKhacGroup? TargetGroup { get; set; }
    public int STT { get; set; }
    public string TenHangMuc { get; set; } = string.Empty;
    public string MaCongTac { get; set; } = string.Empty;
    public string TenCongTac { get; set; } = string.Empty;
    public string DonVi { get; set; } = string.Empty;
    public decimal KhoiLuongHienTai { get; set; }
    public decimal KhoiLuongQuet { get; set; }
    public decimal DonGiaDT { get; set; }
    public string TenExcel { get; set; } = string.Empty;
    public int DongExcel { get; set; }
    public string TrangThai { get; set; } = string.Empty;
    public bool IsHangMucHeader { get; set; } = false;
}

/// <summary>
/// Hộp thoại cho phép người dùng chỉ định chính xác Sheet và Cột Khối lượng từ Excel.
/// Nhận diện nguyên vẹn cấu trúc từng Hạng mục, gán định mức đơn giá vào từng dòng mã hiệu
/// của từng hạng mục (tuyệt đối không cộng dồn khối lượng) để đối chiếu đồng bộ 1:1 với dự toán.
/// Sử dụng cơ chế đọc mảng 2D (Single COM call) siêu tốc, chống crash Excel triệt để.
/// </summary>
public class ChonCotKhoiLuongDialog : Form
{
    private readonly List<CongTacSaiKhacGroup> _targetGroups;
    private readonly List<KetQuaCongTacThamDinh> _auditResults;
    private List<KhoiLuongMappingItem> _mappingList = new();

    private Panel pnlConfig;
    private ComboBox cboSheets;
    private ComboBox cboColMa;
    private ComboBox cboColTen;
    private ComboBox cboColKl;
    private NumericUpDown numDongBatDau;
    private Button btnQuetLai;

    private Button btnChonTatCa;
    private Button btnBoChonTatCa;
    private Button btnXoaDong;

    private DataGridView dgvPreview;
    private Label lblPreviewSummary;
    private Button btnApDung;
    private Button btnDong;

    private bool _isLoading = true;

    /// <summary>
    /// Số lượng công tác thẩm định được gán khối lượng thành công
    /// </summary>
    public int AppliedCount { get; private set; } = 0;

    /// <summary>
    /// Tổng số dòng Excel được tích chọn
    /// </summary>
    public int AppliedRowsCount { get; private set; } = 0;

    /// <summary>
    /// Tên sheet đã chọn để áp dụng
    /// </summary>
    public string SelectedSheetName { get; private set; } = string.Empty;

    /// <summary>
    /// Danh sách các dòng được chọn để tạo các công tác độc lập theo từng hạng mục
    /// </summary>
    public List<KhoiLuongMappingItem> SelectedMappingItems => _mappingList.Where(x => x.Selected && !x.IsHangMucHeader).ToList();

    /// <summary>
    /// Từ điển khối lượng thu thập được (hỗ trợ tương thích ngược)
    /// </summary>
    public Dictionary<string, decimal> DuLieuKhoiLuong { get; private set; } = new(StringComparer.OrdinalIgnoreCase);

    private readonly int? _boDonGiaId;

    public ChonCotKhoiLuongDialog(List<CongTacSaiKhacGroup>? targetGroups = null, List<KetQuaCongTacThamDinh>? auditResults = null, int? boDonGiaId = null)
    {
        _targetGroups = targetGroups ?? new List<CongTacSaiKhacGroup>();
        _boDonGiaId = boDonGiaId;
        if (auditResults != null)
        {
            _auditResults = auditResults;
        }
        else
        {
            _auditResults = _targetGroups.Where(g => g.KetQuaCongTac != null).Select(g => g.KetQuaCongTac).ToList();
        }

        InitializeComponent();
        FormStateHelper.Attach(this, "ChonCotKhoiLuongDialog");
        this.Load += (s, e) =>
        {
            this.WindowState = FormWindowState.Maximized;
        };
        LoadExcelSheetsAndColumns();
    }

    private void InitializeComponent()
    {
        this.Text = "Chỉ Định Sheet & Cột Khối Lượng Dự Toán Từ Excel";

        // Kích thước mặc định rộng rãi, mở toàn màn hình để chống che dữ liệu
        var screen = Screen.FromControl(this) ?? Screen.PrimaryScreen;
        var workArea = screen.WorkingArea;
        int formWidth = Math.Min(1450, Math.Max(1100, workArea.Width - 60));
        int formHeight = Math.Min(880, Math.Max(680, workArea.Height - 60));
        this.Size = new Size(formWidth, formHeight);
        this.MinimumSize = new Size(1020, 620);
        this.WindowState = FormWindowState.Maximized;
        this.MaximizeBox = true;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.ShowInTaskbar = false;
        this.Font = UIHelper.GetFont(9.5f);
        this.BackColor = Color.FromArgb(246, 248, 252);

        // 1. HEADER PANEL
        var pnlHeader = new Panel
        {
            Dock = DockStyle.Top,
            Height = 74,
            BackColor = Color.FromArgb(0, 51, 102),
            Padding = new Padding(18, 10, 18, 10)
        };

        var lblTitle = new Label
        {
            Text = "📐 CHỈ ĐỊNH SHEET & CỘT KHỐI LƯỢNG TỪ FILE EXCEL",
            Font = UIHelper.GetFont(12f, FontStyle.Bold),
            ForeColor = Color.White,
            AutoSize = true,
            Location = new Point(16, 10)
        };

        var lblSubTitle = new Label
        {
            Text = "Giữ nguyên cấu trúc từng Hạng mục & gán định mức, đơn giá thẩm định vào từng dòng mã hiệu (không cộng dồn để đối chiếu chuẩn xác 1:1).",
            Font = UIHelper.GetFont(9f),
            ForeColor = Color.FromArgb(210, 230, 255),
            AutoSize = true,
            Location = new Point(17, 40)
        };

        pnlHeader.Controls.Add(lblTitle);
        pnlHeader.Controls.Add(lblSubTitle);

        // 2. CONFIGURATION PANEL (BỐ CỤC ĐỘNG THÔNG MINH, KHÔNG BAO GIỜ BỊ ĐÈ CHỮ)
        pnlConfig = new Panel
        {
            Dock = DockStyle.Top,
            Height = 106,
            BackColor = Color.White,
            Padding = new Padding(15, 8, 15, 8)
        };

        // Dòng 1: Chọn Sheet, Dòng bắt đầu, Nút quét lại
        var lblSheet = new Label
        {
            Text = "1. Sheet dữ liệu:",
            AutoSize = true,
            Location = new Point(15, 14),
            Font = UIHelper.GetFont(9.5f, FontStyle.Bold)
        };

        cboSheets = new ComboBox
        {
            Location = new Point(130, 10),
            Width = 320,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = UIHelper.GetFont(9.5f)
        };
        cboSheets.SelectedIndexChanged += (s, e) =>
        {
            if (!_isLoading)
            {
                CapNhatDanhSachCotTheoSheet();
                ThucHienQuetVaKhopDuLieu();
            }
        };

        var lblDong = new Label
        {
            Text = "Dòng bắt đầu:",
            AutoSize = true,
            Location = new Point(475, 14),
            Font = UIHelper.GetFont(9.5f)
        };

        numDongBatDau = new NumericUpDown
        {
            Location = new Point(575, 11),
            Width = 70,
            Minimum = 1,
            Maximum = 500,
            Value = 7,
            Font = UIHelper.GetFont(9.5f)
        };
        numDongBatDau.ValueChanged += (s, e) =>
        {
            if (!_isLoading) ThucHienQuetVaKhopDuLieu();
        };

        btnQuetLai = new Button
        {
            Text = "⚡ Quét lại dữ liệu",
            Location = new Point(665, 8),
            Width = 160,
            Height = 32,
            BackColor = Color.FromArgb(235, 245, 255),
            ForeColor = Color.FromArgb(0, 102, 204),
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnQuetLai.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 215);
        btnQuetLai.Click += (s, e) => ThucHienQuetVaKhopDuLieu();

        // Dòng 2: Cột Mã hiệu, Cột Tên, Cột Khối lượng (khoảng cách rộng rãi, chống đè chữ)
        var lblColMa = new Label
        {
            Text = "2. Cột Mã hiệu:",
            AutoSize = true,
            Location = new Point(15, 56),
            Font = UIHelper.GetFont(9.5f)
        };

        cboColMa = new ComboBox
        {
            Location = new Point(130, 52),
            Width = 190,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = UIHelper.GetFont(9f)
        };
        cboColMa.SelectedIndexChanged += (s, e) =>
        {
            if (!_isLoading) ThucHienQuetVaKhopDuLieu();
        };

        var lblColTen = new Label
        {
            Text = "3. Cột Tên CV:",
            AutoSize = true,
            Location = new Point(340, 56),
            Font = UIHelper.GetFont(9.5f)
        };

        cboColTen = new ComboBox
        {
            Location = new Point(440, 52),
            Width = 250,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = UIHelper.GetFont(9f)
        };
        cboColTen.SelectedIndexChanged += (s, e) =>
        {
            if (!_isLoading) ThucHienQuetVaKhopDuLieu();
        };

        var lblColKl = new Label
        {
            Text = "4. Cột Khối lượng:",
            AutoSize = true,
            Location = new Point(715, 56),
            Font = UIHelper.GetFont(9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 102, 204)
        };

        cboColKl = new ComboBox
        {
            Location = new Point(845, 52),
            Width = 250,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = UIHelper.GetFont(9.5f, FontStyle.Bold)
        };
        cboColKl.SelectedIndexChanged += (s, e) =>
        {
            if (!_isLoading) ThucHienQuetVaKhopDuLieu();
        };

        pnlConfig.Controls.Add(lblSheet);
        pnlConfig.Controls.Add(cboSheets);
        pnlConfig.Controls.Add(lblDong);
        pnlConfig.Controls.Add(numDongBatDau);
        pnlConfig.Controls.Add(btnQuetLai);
        pnlConfig.Controls.Add(lblColMa);
        pnlConfig.Controls.Add(cboColMa);
        pnlConfig.Controls.Add(lblColTen);
        pnlConfig.Controls.Add(cboColTen);
        pnlConfig.Controls.Add(lblColKl);
        pnlConfig.Controls.Add(cboColKl);

        // 3. TOOLBAR PHÍA TRÊN LƯỚI
        var pnlGridToolbar = new Panel
        {
            Dock = DockStyle.Top,
            Height = 40,
            BackColor = Color.FromArgb(240, 244, 250),
            Padding = new Padding(15, 6, 15, 6)
        };

        var lblListTitle = new Label
        {
            Text = "📋 Danh sách công tác theo từng Hạng mục (Gán định mức, đơn giá cho từng dòng):",
            AutoSize = true,
            Location = new Point(15, 11),
            Font = UIHelper.GetFont(9f, FontStyle.Bold),
            ForeColor = Color.FromArgb(50, 50, 50)
        };

        btnChonTatCa = new Button
        {
            Text = "✓ Chọn tất cả",
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            Location = new Point(pnlGridToolbar.Width - 430, 6),
            Width = 100,
            Height = 28,
            BackColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(8.5f),
            Cursor = Cursors.Hand
        };
        btnChonTatCa.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180);
        btnChonTatCa.Click += (s, e) => DatTrangThaiTatCa(true);

        btnBoChonTatCa = new Button
        {
            Text = "✕ Bỏ chọn",
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            Location = new Point(pnlGridToolbar.Width - 325, 6),
            Width = 90,
            Height = 28,
            BackColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(8.5f),
            Cursor = Cursors.Hand
        };
        btnBoChonTatCa.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180);
        btnBoChonTatCa.Click += (s, e) => DatTrangThaiTatCa(false);

        btnXoaDong = new Button
        {
            Text = "🗑 Xóa dòng đã chọn / thừa",
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            Location = new Point(pnlGridToolbar.Width - 230, 6),
            Width = 215,
            Height = 28,
            BackColor = Color.FromArgb(255, 240, 240),
            ForeColor = Color.FromArgb(180, 30, 30),
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(8.5f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnXoaDong.FlatAppearance.BorderColor = Color.FromArgb(220, 100, 100);
        btnXoaDong.Click += BtnXoaDong_Click;

        pnlGridToolbar.SizeChanged += (s, e) =>
        {
            btnChonTatCa.Location = new Point(pnlGridToolbar.Width - 430, 6);
            btnBoChonTatCa.Location = new Point(pnlGridToolbar.Width - 325, 6);
            btnXoaDong.Location = new Point(pnlGridToolbar.Width - 230, 6);
        };

        pnlGridToolbar.Controls.Add(lblListTitle);
        pnlGridToolbar.Controls.Add(btnChonTatCa);
        pnlGridToolbar.Controls.Add(btnBoChonTatCa);
        pnlGridToolbar.Controls.Add(btnXoaDong);

        // 4. BOTTOM PANEL (2 TẦNG: TẦNG 1 THÔNG BÁO TÓM TẮT, TẦNG 2 NÚT BẤM)
        var pnlBottom = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 90,
            BackColor = Color.White,
            Padding = new Padding(18, 8, 18, 10)
        };

        lblPreviewSummary = new Label
        {
            Text = "Đang quét dữ liệu...",
            AutoSize = true,
            Location = new Point(18, 12),
            Font = UIHelper.GetFont(9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(60, 60, 60)
        };

        btnApDung = new Button
        {
            Text = "✓  Áp dụng Khối Lượng vào Bảng Thẩm Định (Theo Từng Hạng Mục)",
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            Location = new Point(pnlBottom.Width - 520, 38),
            Width = 405,
            Height = 44,
            BackColor = Color.FromArgb(16, 124, 65),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(10f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnApDung.FlatAppearance.BorderSize = 0;
        btnApDung.Click += BtnApDung_Click;

        btnDong = new Button
        {
            Text = "Đóng",
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            Location = new Point(pnlBottom.Width - 105, 38),
            Width = 90,
            Height = 44,
            BackColor = Color.FromArgb(240, 240, 240),
            ForeColor = Color.Black,
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9.5f),
            Cursor = Cursors.Hand
        };
        btnDong.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
        btnDong.Click += (s, e) => this.Close();

        pnlBottom.SizeChanged += (s, e) =>
        {
            btnApDung.Location = new Point(pnlBottom.Width - 520, 38);
            btnDong.Location = new Point(pnlBottom.Width - 105, 38);
        };

        pnlBottom.Controls.Add(lblPreviewSummary);
        pnlBottom.Controls.Add(btnApDung);
        pnlBottom.Controls.Add(btnDong);

        // 5. MAIN PREVIEW GRID (TỰ CO GIÃN DÒNG & CỘT, KHÔNG CHE CHỮ)
        var pnlCenter = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(15, 5, 15, 5),
            BackColor = Color.FromArgb(246, 248, 252)
        };

        dgvPreview = new DataGridView
        {
            Dock = DockStyle.Fill,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = true,
            AutoGenerateColumns = false,
            RowHeadersVisible = false,
            ColumnHeadersHeight = 48,
            AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
            DefaultCellStyle = { WrapMode = DataGridViewTriState.True }
        };
        UIHelper.ApplyStyle(dgvPreview);

        // Cột 1: Checkbox Sử dụng
        var colCheck = new DataGridViewCheckBoxColumn
        {
            Name = "colCheck",
            HeaderText = "Áp dụng",
            Width = 65,
            ReadOnly = false
        };
        dgvPreview.Columns.Add(colCheck);

        // Cột 2: STT
        dgvPreview.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colSTT",
            HeaderText = "STT",
            Width = 50,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
        });

        // Cột 3: Hạng mục
        dgvPreview.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colHangMuc",
            HeaderText = "Hạng mục",
            Width = 150,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft, Font = UIHelper.GetFont(9f, FontStyle.Bold) }
        });

        // Cột 4: Mã hiệu
        dgvPreview.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colMa",
            HeaderText = "Mã hiệu",
            Width = 95,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter, Font = UIHelper.GetFont(9f, FontStyle.Bold) }
        });

        // Cột 5: Tên công tác thẩm định (TỰ ĐỘNG GIÃN CỘT VÀ XUỐNG DÒNG)
        dgvPreview.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colTen",
            HeaderText = "Tên công tác thẩm định",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 40,
            MinimumWidth = 240,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft, WrapMode = DataGridViewTriState.True }
        });

        // Cột 6: Khối lượng quét được từ Excel (CHO PHÉP SỬA TRỰC TIẾP)
        dgvPreview.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colKlQuet",
            HeaderText = "Khối lượng Excel",
            Width = 130,
            ReadOnly = false,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Font = UIHelper.GetFont(9.5f, FontStyle.Bold), ForeColor = Color.FromArgb(0, 102, 204), Format = "#,##0.000", BackColor = Color.FromArgb(255, 255, 235) }
        });

        // Cột 7: Tên công tác trong Excel đối chiếu (TỰ ĐỘNG GIÃN CỘT VÀ XUỐNG DÒNG)
        dgvPreview.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colTenExcel",
            HeaderText = "Tên công tác trong Excel (đối chiếu)",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 40,
            MinimumWidth = 240,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft, ForeColor = Color.FromArgb(70, 70, 70), WrapMode = DataGridViewTriState.True }
        });

        // Cột 8: Trạng thái khớp
        dgvPreview.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colTrangThai",
            HeaderText = "Trạng thái",
            Width = 160,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter, Font = UIHelper.GetFont(8.5f, FontStyle.Bold) }
        });

        dgvPreview.CellValueChanged += DgvPreview_CellValueChanged;
        dgvPreview.CurrentCellDirtyStateChanged += (s, e) =>
        {
            if (dgvPreview.IsCurrentCellDirty) dgvPreview.CommitEdit(DataGridViewDataErrorContexts.Commit);
        };
        dgvPreview.DataError += (s, e) => { e.ThrowException = false; };

        pnlCenter.Controls.Add(dgvPreview);

        this.Controls.Add(pnlCenter);
        this.Controls.Add(pnlGridToolbar);
        this.Controls.Add(pnlBottom);
        this.Controls.Add(pnlConfig);
        this.Controls.Add(pnlHeader);
    }

    private void DgvPreview_CellValueChanged(object sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.RowIndex >= _mappingList.Count) return;

        var item = _mappingList[e.RowIndex];
        string colName = dgvPreview.Columns[e.ColumnIndex].Name;

        if (colName == "colCheck")
        {
            item.Selected = Convert.ToBoolean(dgvPreview.Rows[e.RowIndex].Cells["colCheck"].Value);
            CapNhatTongKet();
        }
        else if (colName == "colKlQuet")
        {
            object val = dgvPreview.Rows[e.RowIndex].Cells["colKlQuet"].Value;
            if (val != null && decimal.TryParse(val.ToString().Replace(" ", "").Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal newKl))
            {
                item.KhoiLuongQuet = newKl;
            }
            CapNhatTongKet();
        }
    }

    private void DatTrangThaiTatCa(bool selected)
    {
        foreach (var item in _mappingList)
        {
            if (!item.IsHangMucHeader && item.KhoiLuongQuet > 0)
            {
                item.Selected = selected;
            }
        }
        NapDanhSachLenGrid();
        CapNhatTongKet();
    }

    private void BtnXoaDong_Click(object sender, EventArgs e)
    {
        var selectedRows = dgvPreview.SelectedRows;
        if (selectedRows.Count == 0)
        {
            MessageBox.Show("Vui lòng chọn các dòng cần xóa bỏ trong bảng.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var itemsToRemove = new List<KhoiLuongMappingItem>();
        foreach (DataGridViewRow row in selectedRows)
        {
            if (row.Tag is KhoiLuongMappingItem item)
            {
                itemsToRemove.Add(item);
            }
        }

        var dr = MessageBox.Show(
            $"Bạn có chắc chắn muốn xóa {itemsToRemove.Count} dòng này khỏi danh sách khối lượng áp dụng không?",
            "Xác nhận xóa",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (dr == DialogResult.Yes)
        {
            foreach (var item in itemsToRemove)
            {
                _mappingList.Remove(item);
            }
            NapDanhSachLenGrid();
            CapNhatTongKet();
        }
    }

    private void LoadExcelSheetsAndColumns()
    {
        _isLoading = true;
        try
        {
            var app = (Application)ExcelDnaUtil.Application;
            var wb = app?.ActiveWorkbook;

            if (wb == null)
            {
                lblPreviewSummary.Text = "Không tìm thấy file Excel đang mở.";
                btnApDung.Enabled = false;
                return;
            }

            cboSheets.Items.Clear();
            int selectedIdx = 0;
            int idx = 0;

            foreach (Worksheet ws in wb.Worksheets)
            {
                cboSheets.Items.Add(ws.Name);
                string cleanName = BoDau(ws.Name);
                if (cleanName.Contains("du toan") || cleanName.Contains("dutoan") || cleanName.Contains("dt") ||
                    cleanName.Contains("khoi luong") || cleanName.Contains("kl") || cleanName.Contains("tong hop"))
                {
                    selectedIdx = idx;
                }
                idx++;
            }

            if (cboSheets.Items.Count > 0)
            {
                cboSheets.SelectedIndex = selectedIdx;
                CapNhatDanhSachCotTheoSheet();
            }
        }
        catch (Exception ex)
        {
            lblPreviewSummary.Text = "Lỗi khi kết nối Excel: " + ex.Message;
        }
        finally
        {
            _isLoading = false;
            ThucHienQuetVaKhopDuLieu();
        }
    }

    private void CapNhatDanhSachCotTheoSheet()
    {
        string sheetName = cboSheets.SelectedItem?.ToString();
        if (string.IsNullOrEmpty(sheetName)) return;

        var app = (Application)ExcelDnaUtil.Application;
        var wb = app?.ActiveWorkbook;
        if (wb == null) return;

        Worksheet ws = null;
        try { ws = wb.Worksheets[sheetName] as Worksheet; } catch { }
        if (ws == null) return;

        var colList = new List<string>();
        for (char c = 'A'; c <= 'Z'; c++) colList.Add(c.ToString());
        colList.Add("AA"); colList.Add("AB"); colList.Add("AC"); colList.Add("AD"); colList.Add("AE"); colList.Add("AF");

        var colHeaders = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        Range headerRange = null;

        try
        {
            int maxColCheck = Math.Min(colList.Count, 32);
            string lastColLetter = colList[maxColCheck - 1];

            // ĐỌC TOÀN BỘ 6 DÒNG ĐẦU BẰNG 1 LẦN GỌI COM (SINGLE-CALL)
            headerRange = ws.Range["A1", $"{lastColLetter}6"];
            object[,] rawHeaders = headerRange.Value2 as object[,];

            if (rawHeaders != null)
            {
                int rowCount = rawHeaders.GetLength(0);
                int colCount = rawHeaders.GetLength(1);

                for (int colIdx = 1; colIdx <= colCount; colIdx++)
                {
                    string colLetter = colList[colIdx - 1];
                    string detectedHeader = "";

                    for (int r = 1; r <= rowCount; r++)
                    {
                        object val = rawHeaders[r, colIdx];
                        if (val == null) continue;

                        string s = val.ToString().Trim().Replace("\r", " ").Replace("\n", " ");
                        if (string.IsNullOrEmpty(s) || s.Length >= 50) continue;

                        // Bỏ qua nếu là số liệu thuần túy (e.g. 4.9762, 12345), tránh lấy nhầm số liệu vào combobox
                        if (double.TryParse(s.Replace(" ", "").Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _))
                            continue;

                        string clean = BoDau(s);
                        // Ưu tiên các từ khóa tiêu đề chuẩn
                        if (clean.Contains("ma hieu") || clean.Contains("ma cv") || clean.Contains("ten cong tac") || clean.Contains("ten cv") ||
                            clean.Contains("khoi luong") || clean.Contains("don gia") || clean.Contains("thanh tien") || clean == "kl" || clean == "ma" || clean == "dvt")
                        {
                            detectedHeader = s;
                            break; // Giữ tiêu đề chuẩn, không ghi đè
                        }

                        if (string.IsNullOrEmpty(detectedHeader))
                        {
                            detectedHeader = s;
                        }
                    }

                    if (!string.IsNullOrEmpty(detectedHeader))
                    {
                        colHeaders[colLetter] = detectedHeader;
                    }
                }
            }
        }
        catch { }
        finally
        {
            if (headerRange != null)
            {
                try { Marshal.FinalReleaseComObject(headerRange); } catch { }
            }
        }

        cboColMa.Items.Clear();
        cboColTen.Items.Clear();
        cboColKl.Items.Clear();

        int defMaIdx = -1, defTenIdx = -1, defKlIdx = -1;

        for (int i = 0; i < colList.Count; i++)
        {
            string col = colList[i];
            string label = col;
            if (colHeaders.TryGetValue(col, out string headerText))
            {
                label = $"{col} ({headerText})";
                string cleanH = BoDau(headerText);
                bool isPriceOrAmount = cleanH.Contains("don gia") || cleanH.Contains("dongia") || cleanH.Contains("vat lieu") || cleanH.Contains("thanh tien") || cleanH.Contains("nhan cong") || cleanH.Contains("may");

                if (defMaIdx == -1 && (cleanH.Contains("ma hieu") || cleanH.Contains("ma cv") || cleanH.Contains("ma dinh muc") || cleanH.Contains("so hieu") || cleanH == "ma"))
                    defMaIdx = i;

                if (defTenIdx == -1 && (cleanH.Contains("ten cong tac") || cleanH.Contains("ten cv") || cleanH.Contains("noi dung") || cleanH.Contains("danh muc") || cleanH == "ten"))
                    defTenIdx = i;

                if (defKlIdx == -1 && !isPriceOrAmount && (cleanH.Contains("khoi luong") || cleanH == "kl" || cleanH.Contains("khoiluong") || cleanH.Contains("so luong") || cleanH.Contains("thiet ke")))
                    defKlIdx = i;
            }

            cboColMa.Items.Add(new ComboboxItem { Text = label, Value = col });
            cboColTen.Items.Add(new ComboboxItem { Text = label, Value = col });
            cboColKl.Items.Add(new ComboboxItem { Text = label, Value = col });
        }

        cboColMa.SelectedIndex = defMaIdx != -1 ? defMaIdx : 1; // Cột B
        cboColTen.SelectedIndex = defTenIdx != -1 ? defTenIdx : 2; // Cột C
        cboColKl.SelectedIndex = defKlIdx != -1 ? defKlIdx : (colList.Count > 4 ? 4 : 0); // Cột E
    }

    private string GetSelectedCol(ComboBox cbo)
    {
        if (cbo?.SelectedItem is ComboboxItem item) return item.Value;
        return cbo?.SelectedItem?.ToString() ?? "A";
    }

    /// <summary>
    /// Quét toàn bộ dòng từ Excel bằng mảng 2D trong 1 lần gọi COM duy nhất.
    /// Nhận diện nguyên vẹn từng Hạng mục và gán định mức đơn giá vào từng dòng mã hiệu độc lập.
    /// </summary>
    private void ThucHienQuetVaKhopDuLieu()
    {
        _mappingList.Clear();
        dgvPreview.Rows.Clear();

        string sheetName = cboSheets.SelectedItem?.ToString();
        if (string.IsNullOrEmpty(sheetName)) return;

        var app = (Application)ExcelDnaUtil.Application;
        var wb = app?.ActiveWorkbook;
        if (wb == null) return;

        Worksheet ws = null;
        try { ws = wb.Worksheets[sheetName] as Worksheet; } catch { }
        if (ws == null) return;

        string colMa = GetSelectedCol(cboColMa);
        string colTen = GetSelectedCol(cboColTen);
        string colKl = GetSelectedCol(cboColKl);
        int startRow = (int)numDongBatDau.Value;

        Range dataRange = null;

        try
        {
            Range used = ws.UsedRange;
            int maxRow = used != null ? used.Rows.Count + used.Row - 1 : 1000;
            if (maxRow < startRow) maxRow = startRow + 500;
            maxRow = Math.Min(maxRow, 3000); // Giới hạn an toàn

            int colMaIdx = ColLetterToIndex(colMa);
            int colTenIdx = !string.IsNullOrEmpty(colTen) ? ColLetterToIndex(colTen) : -1;
            int colKlIdx = ColLetterToIndex(colKl);

            int maxColIdx = Math.Max(colMaIdx, Math.Max(colTenIdx > 0 ? colTenIdx : 1, colKlIdx));
            string maxColLetter = IndexToColLetter(maxColIdx);

            // =========================================================================
            // ĐỌC TOÀN BỘ VÙNG DỮ LIỆU TRONG ĐÚNG 1 LẦN GỌI COM (SINGLE-CALL)
            // =========================================================================
            dataRange = ws.Range[$"A{startRow}", $"{maxColLetter}{maxRow}"];
            object[,] rawValues = dataRange.Value2 as object[,];

            var excelRows = new List<ExcelRowData>();
            if (rawValues != null)
            {
                int rowCount = rawValues.GetLength(0);
                for (int i = 1; i <= rowCount; i++)
                {
                    int actualRow = startRow + i - 1;
                    object valMa = colMaIdx <= rawValues.GetLength(1) ? rawValues[i, colMaIdx] : null;
                    object valTen = (colTenIdx > 0 && colTenIdx <= rawValues.GetLength(1)) ? rawValues[i, colTenIdx] : null;
                    object valKl = colKlIdx <= rawValues.GetLength(1) ? rawValues[i, colKlIdx] : null;

                    string ma = valMa?.ToString()?.Trim() ?? "";
                    string ten = valTen?.ToString()?.Trim() ?? "";
                    decimal kl = ParseDecimal(valKl);

                    if (!string.IsNullOrEmpty(ma) || !string.IsNullOrEmpty(ten) || kl > 0)
                    {
                        excelRows.Add(new ExcelRowData
                        {
                            RowIndex = actualRow,
                            MaHieu = ma,
                            TenCongTac = ten,
                            KhoiLuong = kl
                        });
                    }
                }
            }

            // =========================================================================
            // BÓC TÁCH TỪNG HẠNG MỤC & KHỚP ĐƠN GIÁ VÀO TỪNG DÒNG ĐỘC LẬP
            // =========================================================================
            AIE.ExcelAddIn.Services.ThamDinhEngine? dbEngine = null;
            try
            {
                var db = new AIE.Data.DatabaseManager();
                var ctRepo = new AIE.Data.Repositories.CongTacRepository(db.Context);
                var vlRepo = new AIE.Data.Repositories.VatLieuRepository(db.Context);
                var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);
                var mayRepo = new AIE.Data.Repositories.MayThiCongRepository(db.Context);
                dbEngine = new AIE.ExcelAddIn.Services.ThamDinhEngine(ctRepo, vlRepo, ncRepo, mayRepo);
            }
            catch { }

            string currentHangMuc = "Hạng mục chung";
            int workItemCounter = 0;

            foreach (var er in excelRows)
            {
                string tenClean = BoDau(er.TenCongTac);
                string maClean = er.MaHieu.Trim();

                // 1. Nhận diện dòng TIÊU ĐỀ HẠNG MỤC (Ví dụ: I, II, III hoặc "Hạng mục...", "Tuyến...")
                bool isHM = false;
                if (IsRomanNumeral(maClean) && er.KhoiLuong == 0) isHM = true;
                else if (string.IsNullOrEmpty(maClean) && er.KhoiLuong == 0)
                {
                    if (tenClean.StartsWith("hang muc") || tenClean.StartsWith("tuyen ") || tenClean.StartsWith("doan ") ||
                        tenClean.StartsWith("cong ") || tenClean.StartsWith("phan ") || tenClean.StartsWith("goi thau"))
                    {
                        isHM = true;
                    }
                }

                if (isHM)
                {
                    currentHangMuc = !string.IsNullOrEmpty(maClean) ? $"{maClean}. {er.TenCongTac}" : er.TenCongTac;

                    _mappingList.Add(new KhoiLuongMappingItem
                    {
                        Selected = false,
                        IsHangMucHeader = true,
                        TargetGroup = null,
                        MatchedResult = null,
                        STT = 0,
                        TenHangMuc = currentHangMuc,
                        MaCongTac = maClean,
                        TenCongTac = $">> HẠNG MỤC: {currentHangMuc.ToUpper()}",
                        KhoiLuongHienTai = 0,
                        KhoiLuongQuet = 0,
                        TenExcel = er.TenCongTac,
                        DongExcel = er.RowIndex,
                        TrangThai = "📁 Tiêu đề Hạng mục"
                    });
                    continue;
                }

                // 2. DÒNG CÔNG TÁC (Có khối lượng > 0 hoặc có mã hiệu)
                if (er.KhoiLuong > 0 || !string.IsNullOrEmpty(maClean))
                {
                    workItemCounter++;

                    // 1. Tìm kết quả thẩm định đã có (để lấy định mức và đơn giá chuẩn)
                    KetQuaCongTacThamDinh? matchedKq = null;
                    string normMaEr = NormalizeMa(maClean);

                    if (!string.IsNullOrEmpty(normMaEr) && _auditResults != null)
                    {
                        matchedKq = _auditResults.FirstOrDefault(r => NormalizeMa(r.DuToan.MaHieu) == normMaEr);
                    }

                    // 2. Nếu không khớp mã, thử tìm theo độ tương đồng tên công tác
                    if (matchedKq == null && _auditResults != null && !string.IsNullOrEmpty(er.TenCongTac))
                    {
                        string normTenEr = BoDau(er.TenCongTac);
                        if (normTenEr.Length > 5)
                        {
                            matchedKq = _auditResults.FirstOrDefault(r =>
                            {
                                string normTenR = BoDau(r.DuToan.TenCongTac);
                                return normTenR.Length > 5 && (normTenR.Contains(normTenEr) || normTenEr.Contains(normTenR));
                            });
                        }
                    }

                    // 3. Nếu vẫn chưa khớp nhưng có mã hiệu và có dbEngine: tra cứu trực tiếp trong Bộ Đơn Giá / TT38 từ Database SQLite
                    if (matchedKq == null && !string.IsNullOrEmpty(normMaEr) && dbEngine != null)
                    {
                        try
                        {
                            var ct = new CongTacThamDinh
                            {
                                MaHieu = maClean,
                                TenCongTac = er.TenCongTac,
                                DonVi = "",
                                KhoiLuong = er.KhoiLuong,
                                DonGia = 0
                            };
                            var res = dbEngine.KiemTra(new List<CongTacThamDinh> { ct }, _boDonGiaId, AIE.Core.Enums.Vung.VungII);
                            if (res != null && res.Count > 0 && res[0].DinhMucChuan != null)
                            {
                                matchedKq = res[0];
                                _auditResults?.Add(matchedKq);
                            }
                        }
                        catch { }
                    }

                    // 4. Nếu vẫn không có trong Database (công tác tạm tính như VL, TT, v.v.): tạo công tác tạm tính hợp lệ
                    if (matchedKq == null)
                    {
                        var ct = new CongTacThamDinh
                        {
                            MaHieu = !string.IsNullOrEmpty(maClean) ? maClean : "TT",
                            TenCongTac = !string.IsNullOrEmpty(er.TenCongTac) ? er.TenCongTac : "(Công tác tạm tính)",
                            DonVi = "",
                            KhoiLuong = er.KhoiLuong,
                            DonGia = 0
                        };
                        matchedKq = new KetQuaCongTacThamDinh
                        {
                            DuToan = ct,
                            DinhMucChuan = null,
                            DanhSachSaiLech = new List<SaiLechDinhMuc>()
                        };
                        _auditResults?.Add(matchedKq);

                        _mappingList.Add(new KhoiLuongMappingItem
                        {
                            Selected = true,
                            IsHangMucHeader = false,
                            MatchedResult = matchedKq,
                            STT = workItemCounter,
                            TenHangMuc = currentHangMuc,
                            MaCongTac = ct.MaHieu,
                            TenCongTac = ct.TenCongTac,
                            DonVi = ct.DonVi,
                            KhoiLuongHienTai = er.KhoiLuong,
                            KhoiLuongQuet = er.KhoiLuong,
                            DonGiaDT = 0,
                            TenExcel = er.TenCongTac,
                            DongExcel = er.RowIndex,
                            TrangThai = $"✓ Tạm tính [{currentHangMuc}]"
                        });
                    }
                    else
                    {
                        _mappingList.Add(new KhoiLuongMappingItem
                        {
                            Selected = true,
                            IsHangMucHeader = false,
                            MatchedResult = matchedKq,
                            STT = workItemCounter,
                            TenHangMuc = currentHangMuc,
                            MaCongTac = matchedKq.DuToan.MaHieu,
                            TenCongTac = matchedKq.DuToan.TenCongTac,
                            DonVi = matchedKq.DuToan.DonVi,
                            KhoiLuongHienTai = er.KhoiLuong,
                            KhoiLuongQuet = er.KhoiLuong,
                            DonGiaDT = matchedKq.DuToan.DonGia,
                            TenExcel = er.TenCongTac,
                            DongExcel = er.RowIndex,
                            TrangThai = matchedKq.DinhMucChuan != null ? $"✓ Khớp mã [{currentHangMuc}]" : $"✓ Tạm tính [{currentHangMuc}]"
                        });
                    }
                }
            }

            NapDanhSachLenGrid();
            CapNhatTongKet();
        }
        catch (Exception ex)
        {
            lblPreviewSummary.Text = "Lỗi khi quét dữ liệu: " + ex.Message;
        }
        finally
        {
            if (dataRange != null)
            {
                try { Marshal.FinalReleaseComObject(dataRange); } catch { }
            }
        }
    }

    private void NapDanhSachLenGrid()
    {
        dgvPreview.Rows.Clear();

        for (int i = 0; i < _mappingList.Count; i++)
        {
            var item = _mappingList[i];

            if (item.IsHangMucHeader)
            {
                int rIdx = dgvPreview.Rows.Add(
                    false,
                    "",
                    item.TenHangMuc,
                    "",
                    item.TenCongTac,
                    (object)"",
                    item.TenExcel,
                    item.TrangThai
                );

                var row = dgvPreview.Rows[rIdx];
                row.Tag = item;
                row.DefaultCellStyle.BackColor = Color.FromArgb(225, 235, 250);
                row.DefaultCellStyle.Font = UIHelper.GetFont(9.5f, FontStyle.Bold);
                row.Cells["colTen"].Style.ForeColor = Color.FromArgb(0, 51, 102);
                row.Cells["colCheck"].ReadOnly = true;
                row.Cells["colKlQuet"].ReadOnly = true;
                continue;
            }

            int itemIdx = dgvPreview.Rows.Add(
                item.Selected,
                item.STT > 0 ? (object)item.STT : "",
                item.TenHangMuc,
                item.MaCongTac,
                item.TenCongTac,
                item.KhoiLuongQuet,
                item.TenExcel,
                item.TrangThai
            );

            var itemRow = dgvPreview.Rows[itemIdx];
            itemRow.Tag = item;

            if (item.TrangThai.StartsWith("✓"))
            {
                itemRow.Cells["colTrangThai"].Style.ForeColor = Color.FromArgb(0, 128, 64);
                itemRow.DefaultCellStyle.BackColor = Color.White;
            }
            else
            {
                itemRow.Cells["colTrangThai"].Style.ForeColor = Color.FromArgb(130, 130, 130);
                itemRow.DefaultCellStyle.BackColor = Color.FromArgb(248, 248, 248);
            }
        }

        dgvPreview.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
    }

    private void CapNhatTongKet()
    {
        int total = _mappingList.Count;
        int matchedRows = _mappingList.Count(x => x.MatchedResult != null);
        int selectedRows = _mappingList.Count(x => x.Selected && !x.IsHangMucHeader && x.KhoiLuongQuet > 0);
        int soHM = _mappingList.Where(x => !string.IsNullOrEmpty(x.TenHangMuc)).Select(x => x.TenHangMuc).Distinct().Count();

        lblPreviewSummary.Text = $"Đã quét: {total} dòng ({matchedRows} dòng công tác thuộc {soHM} hạng mục) | Đang chọn áp dụng: {selectedRows} dòng công tác độc lập";
        lblPreviewSummary.ForeColor = selectedRows > 0 ? Color.FromArgb(16, 124, 65) : Color.FromArgb(180, 50, 50);
        btnApDung.Enabled = selectedRows > 0;
    }

    private void BtnApDung_Click(object sender, EventArgs e)
    {
        try
        {
            string sheetName = cboSheets.SelectedItem?.ToString();
            SelectedSheetName = sheetName ?? "";

            var selectedList = SelectedMappingItems;

            if (selectedList.Count == 0)
            {
                MessageBox.Show("Chưa có công tác nào được tích chọn để áp dụng khối lượng.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            AppliedCount = selectedList.Count;
            AppliedRowsCount = selectedList.Count;

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show("Lỗi khi áp dụng khối lượng: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static decimal ParseDecimal(object v)
    {
        if (v == null) return 0;
        if (v is double d) return (decimal)d;
        if (v is float f) return (decimal)f;
        if (v is int i) return (decimal)i;
        if (v is long l) return (decimal)l;
        if (v is decimal m) return m;

        string s = v.ToString().Trim().Replace(" ", "");
        if (string.IsNullOrEmpty(s)) return 0;

        if (s.Contains(",") && s.Contains("."))
        {
            if (s.LastIndexOf(',') > s.LastIndexOf('.'))
                s = s.Replace(".", "").Replace(",", ".");
            else
                s = s.Replace(",", "");
        }
        else if (s.Contains(","))
        {
            s = s.Replace(",", ".");
        }

        if (decimal.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal val))
            return val;

        return 0;
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

    private static string NormalizeMa(string ma)
    {
        if (string.IsNullOrEmpty(ma)) return string.Empty;
        return ma.Replace(".", "").Replace(" ", "").Replace("-", "").ToUpper().Trim();
    }

    private static bool IsRomanNumeral(string s)
    {
        if (string.IsNullOrEmpty(s)) return false;
        string trimmed = s.Trim().ToUpper();
        string[] romans = { "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII", "XIV", "XV" };
        return romans.Contains(trimmed);
    }

    private static int ColLetterToIndex(string colLetter)
    {
        if (string.IsNullOrEmpty(colLetter)) return 1;
        int col = 0;
        foreach (char c in colLetter.Trim().ToUpper())
        {
            if (c >= 'A' && c <= 'Z')
                col = col * 26 + (c - 'A' + 1);
        }
        return col > 0 ? col : 1;
    }

    private static string IndexToColLetter(int colIndex)
    {
        string letter = "";
        while (colIndex > 0)
        {
            int rem = (colIndex - 1) % 26;
            letter = (char)('A' + rem) + letter;
            colIndex = (colIndex - rem) / 26;
        }
        return string.IsNullOrEmpty(letter) ? "A" : letter;
    }

    private class ComboboxItem
    {
        public string Text { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
        public override string ToString() => Text;
    }

    private class ExcelRowData
    {
        public int RowIndex { get; set; }
        public string MaHieu { get; set; } = string.Empty;
        public string TenCongTac { get; set; } = string.Empty;
        public decimal KhoiLuong { get; set; }
    }
}
