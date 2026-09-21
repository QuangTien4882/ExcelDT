using AIE.Core.Models;
using AIE.Data;
using AIE.Data.Repositories;
using AIE.ExcelAddIn.Helpers;
using Dapper;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace AIE.ExcelAddIn.Forms;

public class ThamDinhDonGiaForm : Form
{
    private TabControl tabControl;
    private DataGridView dgvVL, dgvNC, dgvMay;
    private TextBox txtTenBoDonGia;
    private TextBox txtGiaXang;
    private TextBox txtGiaDiezel;
    private TextBox txtGiaDien;
    
    private List<DgVatLieuModel> _vatLieuList;
    private List<DgNhanCongModel> _nhanCongList;
    private List<DgMayThiCongModel> _mayThiCongList;

    public int? SavedBoDonGiaId { get; private set; }
    private int? _loadedBoId = null;
    
    private readonly CultureInfo ViVn = new CultureInfo("vi-VN");
    private bool _suppressRecalc = false;
    private System.Windows.Forms.Timer _giaMayDebounce;
    private List<AIE.Core.Models.NhanCong> _cachedAllNC;
    private Dictionary<string, AIE.Core.Models.DinhMucCaMay_TT37> _cachedDmMay;

    private AIE.Core.Enums.Vung _vungApDung;

    public ThamDinhDonGiaForm(List<VatTuGiaModel> extractedItems, AIE.Core.Enums.Vung vungApDung, int? loadedBoId = null)
    {
        _vungApDung = vungApDung;
        _loadedBoId = loadedBoId;
        _suppressRecalc = true;
        _vatLieuList = new List<DgVatLieuModel>();
        _nhanCongList = new List<DgNhanCongModel>();
        _mayThiCongList = new List<DgMayThiCongModel>();
        InitializeComponent();
        FormStateHelper.Attach(this);
        PrepareData(extractedItems);
        LoadDataToGrids();
        _suppressRecalc = false;
        RecalculateMachineCosts();

        this.Shown += (s, e) =>
        {
            dgvVL?.AutoFit();
            dgvNC?.AutoFit();
            dgvMay?.AutoFit();
        };

        if (tabControl != null)
        {
            tabControl.SelectedIndexChanged += (s, e) =>
            {
                if (tabControl.SelectedIndex == 0) dgvVL?.AutoFit();
                else if (tabControl.SelectedIndex == 1) dgvNC?.AutoFit();
                else if (tabControl.SelectedIndex == 2) dgvMay?.AutoFit();
            };
        }
    }

    /// <summary>
    /// Constructor mở trực tiếp Bộ đơn giá đã lưu từ cơ sở dữ liệu (tính năng "Mở Bộ đơn giá").
    /// Nạp đầy đủ 100% dữ liệu đã lưu, kể cả cước vận chuyển, giá máy, giá nhiên liệu.
    /// </summary>
    public ThamDinhDonGiaForm(int loadedBoId, AIE.Core.Enums.Vung vungApDung)
    {
        _vungApDung = vungApDung;
        _loadedBoId = loadedBoId;
        _suppressRecalc = true;
        _vatLieuList = new List<DgVatLieuModel>();
        _nhanCongList = new List<DgNhanCongModel>();
        _mayThiCongList = new List<DgMayThiCongModel>();
        InitializeComponent();
        FormStateHelper.Attach(this);
        LoadFromDatabase(loadedBoId);
        _suppressRecalc = false;
        RecalculateMachineCosts();

        this.Shown += (s, e) =>
        {
            dgvVL?.AutoFit();
            dgvNC?.AutoFit();
            dgvMay?.AutoFit();
        };

        if (tabControl != null)
        {
            tabControl.SelectedIndexChanged += (s, e) =>
            {
                if (tabControl.SelectedIndex == 0) dgvVL?.AutoFit();
                else if (tabControl.SelectedIndex == 1) dgvNC?.AutoFit();
                else if (tabControl.SelectedIndex == 2) dgvMay?.AutoFit();
            };
        }
    }

    private void InitializeComponent()
    {
        this.Text = "Bộ Đơn Giá Thẩm Định";
        var workingArea = Screen.PrimaryScreen.WorkingArea;
        this.Size = new Size((int)(workingArea.Width * 0.9), (int)(workingArea.Height * 0.9));
        this.StartPosition = FormStartPosition.CenterScreen;
        this.MinimizeBox = true;
        this.MaximizeBox = true;
        this.ShowInTaskbar = true;
        this.Font = new Font("Be Vietnam Pro", 9.5f);

        // Header Panel (Tên bộ đơn giá & Nút lưu)
        var topPanel = new Panel { Dock = DockStyle.Top, Height = 50, Padding = new Padding(12, 10, 12, 5) };
        
        var leftPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Left,
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(0)
        };
        var lblTitle = new Label { Text = "Tên Bộ Đơn Giá:", AutoSize = true, Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold), Margin = new Padding(0, 5, 5, 0) };
        txtTenBoDonGia = new TextBox { Width = 400, Text = $"Đơn giá thẩm định - {DateTime.Now:dd/MM/yyyy HH:mm}", Font = new Font("Be Vietnam Pro", 10f), Margin = new Padding(0) };
        leftPanel.Controls.Add(lblTitle);
        leftPanel.Controls.Add(txtTenBoDonGia);

        var rightPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Right,
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            WrapContents = false,
            Padding = new Padding(0, 0, 10, 0)
        };
        var btnAutoFit = new Button
        {
            Text = "↕ Tự động giãn cột/dòng",
            AutoSize = true,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(240, 240, 240),
            Font = new Font("Be Vietnam Pro", 9f),
            Margin = new Padding(15, 0, 0, 0)
        };
        btnAutoFit.FlatAppearance.BorderColor = Color.LightGray;
        btnAutoFit.Click += (s, e) =>
        {
            DataGridView activeGrid = null;
            if (tabControl.SelectedTab == tabControl.TabPages[0]) activeGrid = dgvVL;
            else if (tabControl.SelectedTab == tabControl.TabPages[1]) activeGrid = dgvNC;
            else if (tabControl.SelectedTab == tabControl.TabPages[2]) activeGrid = dgvMay;

            if (activeGrid != null)
            {
                activeGrid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
                activeGrid.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
            }
        };

        var btnApGiaTuBoCu = new Button
        {
            Text = "📥 Áp giá từ Bộ Đơn Giá cũ",
            AutoSize = true,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(230, 240, 255),
            Font = new Font("Be Vietnam Pro", 9f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 100, 200),
            Margin = new Padding(15, 0, 10, 0)
        };
        btnApGiaTuBoCu.FlatAppearance.BorderColor = Color.FromArgb(180, 210, 255);
        btnApGiaTuBoCu.Click += BtnApGiaTuBoCu_Click;

        rightPanel.Controls.Add(btnAutoFit);
        rightPanel.Controls.Add(btnApGiaTuBoCu);

        topPanel.Controls.Add(leftPanel);
        topPanel.Controls.Add(rightPanel);

        // Bottom Panel
        var bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 55, Padding = new Padding(12, 8, 12, 8) };
        var btnTable = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 1
        };
        btnTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
        btnTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
        btnTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));

        var btnThoat = new Button { Text = "Đóng", Dock = DockStyle.Fill, Margin = new Padding(5, 0, 5, 0), FlatStyle = FlatStyle.Flat, Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold) };
        btnThoat.FlatAppearance.BorderColor = Color.LightGray;
        btnThoat.Click += (s, e) => this.Close();

        var btnLuu = new Button { Text = "💾  Lưu Đơn Giá Thẩm Định", Dock = DockStyle.Fill, Margin = new Padding(5, 0, 5, 0), BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold) };
        btnLuu.FlatAppearance.BorderSize = 0;
        btnLuu.Click += BtnLuu_Click;

        var btnCapNhat = new Button { Text = "🔄 Cập nhật vào Đơn giá gốc", Dock = DockStyle.Fill, Margin = new Padding(5, 0, 5, 0), BackColor = Color.FromArgb(40, 167, 69), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold) };
        btnCapNhat.FlatAppearance.BorderSize = 0;
        btnCapNhat.Click += BtnCapNhatGiaGoc_Click;

        btnTable.Controls.Add(btnThoat, 0, 0);
        btnTable.Controls.Add(btnLuu, 1, 0);
        btnTable.Controls.Add(btnCapNhat, 2, 0);
        bottomPanel.Controls.Add(btnTable);

        tabControl = new TabControl { Dock = DockStyle.Fill, Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold), DrawMode = TabDrawMode.OwnerDrawFixed };
        tabControl.DrawItem += TabControl_DrawItem;

        // Tab Vật Liệu
        var tabVL = new TabPage("Vật liệu");
        dgvVL = CreateGrid();
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaHieu", HeaderText = "Mã VL", DataPropertyName = "MaHieu", ReadOnly = true, Width = 90, Frozen = true });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ten", HeaderText = "Tên vật liệu", DataPropertyName = "Ten", ReadOnly = true, Width = 260, Frozen = true });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", DataPropertyName = "DonVi", ReadOnly = true, Width = 65, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "KhoiLuong", HeaderText = "Khối lượng", DataPropertyName = "KhoiLuong", ReadOnly = true, Width = 95, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaGoc", HeaderText = "Giá gốc", DataPropertyName = "GiaGoc", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.LightYellow } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiBocXep", HeaderText = "Chi phí bốc xếp", DataPropertyName = "ChiPhiBocXep", Width = 120, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.LightYellow } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "CuocVCOTo", HeaderText = "Vận chuyển ô tô", DataPropertyName = "CuocVCOTo", Width = 120, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.LightYellow } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "CuocVCBo", HeaderText = "Vận chuyển bộ", DataPropertyName = "CuocVCBo", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.LightYellow } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaHienTruong", HeaderText = "Giá hiện trường", DataPropertyName = "GiaHienTruong", ReadOnly = true, Width = 130, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, ForeColor = Color.Red, Font = new Font(dgvVL.Font, FontStyle.Bold) } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "ThanhTien", HeaderText = "Thành tiền", DataPropertyName = "ThanhTien", ReadOnly = true, Width = 140, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, ForeColor = Color.FromArgb(0, 102, 204), Font = new Font(dgvVL.Font, FontStyle.Bold) } });
        dgvVL.CellValueChanged += DgvVL_CellValueChanged;
        dgvVL.CellDoubleClick += DgvVL_CellDoubleClick;
        tabVL.Controls.Add(dgvVL);

        // Tab Nhân Công
        var tabNC = new TabPage("Nhân công");
        dgvNC = CreateGrid();
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaHieu", HeaderText = "Mã NC", DataPropertyName = "MaHieu", ReadOnly = true, Width = 100, Frozen = true });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ten", HeaderText = "Tên nhân công", DataPropertyName = "Ten", ReadOnly = true, Width = 250, DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True }, Frozen = true });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", DataPropertyName = "DonVi", ReadOnly = true, Width = 60, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });

        var nhomList = new System.Collections.Generic.List<System.Collections.Generic.KeyValuePair<int, string>>
        {
            new System.Collections.Generic.KeyValuePair<int, string>(1, "Nhóm I"),
            new System.Collections.Generic.KeyValuePair<int, string>(2, "Nhóm II"),
            new System.Collections.Generic.KeyValuePair<int, string>(3, "Nhóm III"),
            new System.Collections.Generic.KeyValuePair<int, string>(4, "Nhóm IV")
        };
        var colNhom = new DataGridViewComboBoxColumn 
        { 
            Name = "NhomNhanCong", 
            HeaderText = "Nhóm nhân công", 
            DataPropertyName = "NhomNhanCong", 
            DataSource = nhomList, 
            DisplayMember = "Value",
            ValueMember = "Key",
            ValueType = typeof(int), 
            Width = 140 
        };
        dgvNC.Columns.Add(colNhom);

        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaHienTruong", HeaderText = "Giá nhân công", DataPropertyName = "GiaHienTruong", Width = 150, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.LightYellow, ForeColor = Color.Red, Font = new Font(dgvNC.Font, FontStyle.Bold) } });
        
        dgvNC.CurrentCellDirtyStateChanged += DgvNC_CurrentCellDirtyStateChanged;
        dgvNC.CellValueChanged += DgvNC_CellValueChanged;

        var lblNCWarning = new Label { Text = "Ghi chú: Sửa trực tiếp chỉ áp dụng cho bảng này.", Dock = DockStyle.Top, Height = 30, ForeColor = Color.DimGray, Font = new Font("Be Vietnam Pro", 9f, FontStyle.Italic), TextAlign = ContentAlignment.MiddleLeft };
        tabNC.Controls.Add(dgvNC);
        tabNC.Controls.Add(lblNCWarning);

        // Tab Máy thi công
        var tabMay = new TabPage("Máy thi công");
        var fuelPanel = new Panel { Dock = DockStyle.Top, Height = 60, Padding = new Padding(10), BackColor = Color.FromArgb(240, 248, 255) };
        int px = 10;
        
        // Vùng áp dụng hiển thị tĩnh
        var lblVungLabel = new Label { Text = "Vùng áp dụng:", AutoSize = true, Location = new Point(px, 20), Font = new Font("Be Vietnam Pro", 9f) };
        fuelPanel.Controls.Add(lblVungLabel);
        px += lblVungLabel.PreferredWidth + 4;
        
        string vungText = _vungApDung switch {
            AIE.Core.Enums.Vung.VungII => "Vùng II",
            AIE.Core.Enums.Vung.VungIII => "Vùng III",
            AIE.Core.Enums.Vung.VungIV => "Vùng IV",
            AIE.Core.Enums.Vung.CuLaoCham => "Cù Lao Chàm",
            _ => "Vùng II"
        };
        var lblVungValue = new Label { Text = vungText, AutoSize = true, Location = new Point(px, 20), Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 215) };
        fuelPanel.Controls.Add(lblVungValue);
        px += lblVungValue.PreferredWidth + 30;
        
        txtGiaXang = AddFuelInput(fuelPanel, "Giá Xăng (đ/lít):", ref px, "");
        txtGiaDiezel = AddFuelInput(fuelPanel, "Giá Diezel (đ/lít):", ref px, "");
        txtGiaDien = AddFuelInput(fuelPanel, "Giá Điện (đ/kWh):", ref px, "");

        dgvMay = CreateGrid();
        dgvMay.ColumnHeadersHeight = 45;
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaHieu", HeaderText = "Mã Máy", DataPropertyName = "MaHieu", ReadOnly = true, Width = 100, Frozen = true });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ten", HeaderText = "Tên máy", DataPropertyName = "Ten", ReadOnly = true, Width = 250, DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True }, Frozen = true });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "SoCaNam", HeaderText = "Số ca/năm", DataPropertyName = "SoCaNam", ReadOnly = true, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "KhauHao", HeaderText = "ĐM Khấu hao", DataPropertyName = "TyLeKhauHao", ReadOnly = false, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, BackColor = Color.FromArgb(255, 255, 200) } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "SuaChua", HeaderText = "ĐM Sửa chữa", DataPropertyName = "TyLeSuaChua", ReadOnly = false, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, BackColor = Color.FromArgb(255, 255, 200) } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiKhac", HeaderText = "ĐM Khác (%)", DataPropertyName = "TyLeKhac", ReadOnly = false, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, BackColor = Color.FromArgb(255, 255, 200) } });
        
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "DinhMucNhienLieuDisplay", HeaderText = "Định mức tiêu hao NL", DataPropertyName = "DinhMucNhienLieuDisplay", ReadOnly = true, Width = 150, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, WrapMode = DataGridViewTriState.True } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "HeSoNhienLieuPhu", HeaderText = "Hệ số NL phụ", DataPropertyName = "HeSoNhienLieuPhu", ReadOnly = true, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Format = "0.##", Alignment = DataGridViewContentAlignment.MiddleCenter } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "NhanCongVanHanhDisplay", HeaderText = "Nhân công vận hành", DataPropertyName = "NhanCongVanHanhDisplay", ReadOnly = true, Width = 250, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleLeft, WrapMode = DataGridViewTriState.True } });
        
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "NguyenGia", HeaderText = "Nguyên giá", DataPropertyName = "NguyenGia", ReadOnly = false, Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.FromArgb(255, 255, 200) } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "KhauHaoGia", HeaderText = "CP Khấu hao", DataPropertyName = "KhauHao", ReadOnly = true, Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "SuaChuaGia", HeaderText = "CP Sửa chữa", DataPropertyName = "SuaChua", ReadOnly = true, Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiKhacGia", HeaderText = "CP Khác", DataPropertyName = "ChiPhiKhac", ReadOnly = true, Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "NhienLieu", HeaderText = "CP Nhiên liệu", DataPropertyName = "NhienLieu", ReadOnly = true, Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "LuongTho", HeaderText = "Lương thợ", DataPropertyName = "LuongTho", ReadOnly = true, Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaHienTruong", HeaderText = "ĐƠN GIÁ CA MÁY", DataPropertyName = "GiaHienTruong", ReadOnly = true, Width = 120, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, ForeColor = Color.Red, Font = new Font(dgvMay.Font, FontStyle.Bold) } });
        
        dgvMay.CellPainting += DgvMay_CellPainting;
        dgvMay.CellDoubleClick += DgvMay_CellDoubleClick;
        dgvMay.CellValueChanged += DgvMay_CellValueChanged;
        tabMay.Controls.Add(dgvMay);
        tabMay.Controls.Add(fuelPanel);

        tabControl.TabPages.Add(tabVL);
        tabControl.TabPages.Add(tabNC);
        tabControl.TabPages.Add(tabMay);

        this.Controls.Add(tabControl);
        this.Controls.Add(topPanel);
        this.Controls.Add(bottomPanel);
    }

    private void TabControl_DrawItem(object? sender, DrawItemEventArgs e)
    {
        if (tabControl == null) return;
        var g = e.Graphics;
        var tabPage = tabControl.TabPages[e.Index];
        var tabRect = tabControl.GetTabRect(e.Index);

        if (tabControl.SelectedIndex == e.Index)
        {
            g.FillRectangle(new SolidBrush(Color.FromArgb(0, 120, 215)), tabRect);
            TextRenderer.DrawText(g, tabPage.Text, new Font(tabControl.Font, FontStyle.Bold), tabRect, Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
        else
        {
            g.FillRectangle(new SolidBrush(Color.WhiteSmoke), tabRect);
            TextRenderer.DrawText(g, tabPage.Text, tabControl.Font, tabRect, Color.Black, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
    }

    private void BtnCapNhatGiaGoc_Click(object? sender, EventArgs e)
    {
        var db = new AIE.Data.DatabaseManager();
        var vlRepo = new AIE.Data.Repositories.VatLieuRepository(db.Context);
        var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);
        var dmMayRepo = new AIE.Data.Repositories.DinhMucCaMayRepository(db.Context);

        int updatedCount = 0;

        foreach (var vl in _vatLieuList)
        {
            var dbVL = vlRepo.GetByMa(vl.MaHieu);
            if (dbVL != null && dbVL.DonGia != vl.GiaGoc)
            {
                dbVL.DonGia = vl.GiaGoc;
                dbVL.NgayCapNhat = DateTime.Now;
                vlRepo.Upsert(dbVL);
                updatedCount++;
            }
        }

        foreach (var nc in _nhanCongList)
        {
            var dbNC = ncRepo.GetAll().FirstOrDefault(x => x.LoaiNhanCong == AIE.Core.Enums.LoaiNhanCong.XayDung && x.Nhom == nc.NhomNhanCong);
            if (dbNC != null && dbNC.GetDonGia(_vungApDung) != nc.GiaHienTruong)
            {
                dbNC.SetDonGia(_vungApDung, nc.GiaHienTruong);
                dbNC.NgayCapNhat = DateTime.Now;
                ncRepo.Upsert(dbNC);
                updatedCount++;
            }
        }

        foreach (var may in _mayThiCongList)
        {
            var dmMay = dmMayRepo.GetByMaMay(may.MaHieu);
            if (dmMay != null)
            {
                bool changed = false;
                if (dmMay.NguyenGia != may.NguyenGia) { dmMay.NguyenGia = may.NguyenGia; changed = true; }
                if (dmMay.KhauHao != may.TyLeKhauHao) { dmMay.KhauHao = may.TyLeKhauHao; changed = true; }
                if (dmMay.SuaChua != may.TyLeSuaChua) { dmMay.SuaChua = may.TyLeSuaChua; changed = true; }
                if (dmMay.ChiPhiKhac != (decimal)may.TyLeKhac) { dmMay.ChiPhiKhac = (decimal)may.TyLeKhac; changed = true; }
                if (changed)
                {
                    dmMayRepo.Upsert(dmMay);
                    updatedCount++;
                }
            }
        }

        if (updatedCount > 0)
        {
            MessageBox.Show($"Đã cập nhật {updatedCount} bản ghi vào cơ sở dữ liệu gốc thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        else
        {
            MessageBox.Show("Không có thay đổi nào cần cập nhật.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private TextBox AddFuelInput(Control parent, string labelText, ref int x, string defaultValue)
    {
        var lbl = new Label { Text = labelText, AutoSize = true, Location = new Point(x, 20), Font = new Font("Be Vietnam Pro", 9f) };
        int lblWidth = lbl.PreferredWidth;
        var txt = new TextBox { Text = defaultValue, Width = 80, Location = new Point(x + lblWidth + 5, 17), Font = new Font("Be Vietnam Pro", 9f), TextAlign = HorizontalAlignment.Right };

        if (_giaMayDebounce == null)
        {
            _giaMayDebounce = new System.Windows.Forms.Timer { Interval = 500 };
            _giaMayDebounce.Tick += (s, e) =>
            {
                _giaMayDebounce.Stop();
                RecalculateMachineCosts();
            };
        }

        txt.TextChanged += (s, e) => { _giaMayDebounce.Stop(); _giaMayDebounce.Start(); };
        parent.Controls.Add(lbl);
        parent.Controls.Add(txt);
        x += lblWidth + 5 + txt.Width + 20;
        return txt;
    }

    private DataGridView CreateGrid()
    {
        var dgv = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = false,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            BackgroundColor = Color.White,
            RowTemplate = { Height = 28 },
            Font = new Font("Be Vietnam Pro", 9.5f),
            EnableHeadersVisualStyles = false,
            ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single,
            RowHeadersVisible = true,
            RowHeadersWidth = 25,
            AllowUserToResizeRows = true,
            AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        };
        
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(235, 235, 235);
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
        dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Be Vietnam Pro", 9f, FontStyle.Bold);
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        dgv.ColumnHeadersHeight = 35;
        
        return dgv;
    }

    private void LoadFromDatabase(int boId)
    {
        var db = new DatabaseManager();
        var repo = new BoDonGiaRepository(db.Context);
        var vlRepo = new VatLieuRepository(db.Context);
        var ncRepo = new NhanCongRepository(db.Context);
        var mayRepo = new MayThiCongRepository(db.Context);
        var dmMayRepo = new DinhMucCaMayRepository(db.Context);

        _vatLieuList = new List<DgVatLieuModel>();
        _nhanCongList = new List<DgNhanCongModel>();
        _mayThiCongList = new List<DgMayThiCongModel>();

        // 1. Nạp thông tin chung & giá nhiên liệu trước
        var boInfo = repo.GetById(boId);
        if (boInfo != null)
        {
            txtTenBoDonGia.Text = !string.IsNullOrWhiteSpace(boInfo.TenBo) ? boInfo.TenBo : $"Bộ đơn giá #{boId}";
            if (boInfo.GiaXang > 0) txtGiaXang.Text = boInfo.GiaXang.ToString("0.####");
            if (boInfo.GiaDiezel > 0) txtGiaDiezel.Text = boInfo.GiaDiezel.ToString("0.####");
            if (boInfo.GiaDien > 0) txtGiaDien.Text = boInfo.GiaDien.ToString("0.####");
        }

        using var conn = db.Context.GetConnection();

        // 2. Nạp Vật liệu (loại trừ sạch vật liệu khác)
        var dsVL = repo.GetGiaVL(boId);
        foreach (var vl in dsVL)
        {
            if (vl.MaVL == "VL_KHAC" || vl.MaVL.StartsWith("VLK")) continue;
            var master = vlRepo.GetByMa(vl.MaVL);
            string name = master?.TenVL ?? conn.QueryFirstOrDefault<string>("SELECT TenHaoPhi FROM HaoPhi WHERE MaHieuHP = @Ma", new { Ma = vl.MaVL }) ?? vl.MaVL;
            string donVi = master?.DonVi ?? conn.QueryFirstOrDefault<string>("SELECT DonVi FROM HaoPhi WHERE MaHieuHP = @Ma", new { Ma = vl.MaVL }) ?? "";
            if (donVi == "%" || name.ToLower().Contains("vật liệu khác")) continue;

            decimal giaGoc = vl.GiaGoc;
            if (giaGoc == 0 && vl.GiaHienTruong > 0) giaGoc = vl.GiaHienTruong;

            _vatLieuList.Add(new DgVatLieuModel
            {
                MaHieu = vl.MaVL,
                Ten = name,
                DonVi = donVi,
                GiaGoc = giaGoc,
                ChiPhiBocXep = vl.ChiPhiBocXep,
                CuocVCOTo = vl.CuocVCOTo,
                CuocVCBo = vl.CuocVCBo
            });
        }

        // 3. Nạp Nhân công
        var dsNC = repo.GetGiaNC(boId);
        foreach (var nc in dsNC)
        {
            if (nc.MaNC.Contains("NC_KHAC")) continue;
            var master = ncRepo.GetByMa(nc.MaNC);
            string name = master?.TenNC ?? conn.QueryFirstOrDefault<string>("SELECT TenHaoPhi FROM HaoPhi WHERE MaHieuHP = @Ma", new { Ma = nc.MaNC }) ?? nc.MaNC;
            string donVi = master?.DonVi ?? conn.QueryFirstOrDefault<string>("SELECT DonVi FROM HaoPhi WHERE MaHieuHP = @Ma", new { Ma = nc.MaNC }) ?? "công";
            if (donVi == "%" || name.ToLower().Contains("nhân công khác")) continue;

            int nhom = master?.Nhom ?? 1;
            _nhanCongList.Add(new DgNhanCongModel
            {
                MaHieu = nc.MaNC,
                Ten = name,
                DonVi = donVi,
                NhomNhanCong = nhom,
                GiaHienTruong = nc.DonGia
            });
        }

        // 4. Nạp Máy thi công
        var dsMay = repo.GetGiaMay(boId);
        var allNC = ncRepo.GetAll();
        foreach (var m in dsMay)
        {
            if (m.MaMay == "M7016") continue;
            var master = mayRepo.GetByMa(m.MaMay);
            string name = master?.TenMay ?? conn.QueryFirstOrDefault<string>("SELECT TenHaoPhi FROM HaoPhi WHERE MaHieuHP = @Ma", new { Ma = m.MaMay }) ?? m.MaMay;
            string donVi = master?.DonVi ?? conn.QueryFirstOrDefault<string>("SELECT DonVi FROM HaoPhi WHERE MaHieuHP = @Ma", new { Ma = m.MaMay }) ?? "ca";
            if (donVi == "%" || name.ToLower().Contains("máy khác")) continue;

            var dmMay = dmMayRepo.GetByMaMay(m.MaMay);
            var mayM = new DgMayThiCongModel
            {
                MaHieu = m.MaMay,
                Ten = name,
                DonVi = donVi,
                DonGiaSaved = m.DonGia
            };

            if (dmMay != null)
            {
                mayM.SoCaNam = dmMay.SoCaNam > 0 ? dmMay.SoCaNam : 250;
                mayM.NguyenGia = dmMay.NguyenGia;
                mayM.TyLeKhauHao = dmMay.KhauHao;
                mayM.TyLeSuaChua = dmMay.SuaChua;
                mayM.TyLeKhac = dmMay.ChiPhiKhac;
                mayM.NhomNhanCong = dmMay.NhomNhanCong;
                mayM.SoLuongNhanCong = dmMay.SoLuongNhanCong;
                mayM.HeSoNhienLieuPhu = dmMay.HeSoNhienLieuPhu;
                mayM.NhanCongString = dmMay.NhanCongString;

                decimal g_th = dmMay.NguyenGia >= 30000000m ? dmMay.NguyenGia * 0.1m : 0m;
                mayM.KhauHao = ((dmMay.NguyenGia - g_th) * dmMay.KhauHao / 100m) / mayM.SoCaNam;
                mayM.SuaChua = (dmMay.NguyenGia * dmMay.SuaChua / 100m) / mayM.SoCaNam;
                mayM.ChiPhiKhac = (dmMay.NguyenGia * dmMay.ChiPhiKhac / 100m) / mayM.SoCaNam;

                mayM.DinhMucXang = dmMay.DinhMucXang;
                mayM.DinhMucDiezel = dmMay.DinhMucDiezel;
                mayM.DinhMucDien = dmMay.DinhMucDien;

                mayM.LuongTho = 0;
                var tpNC = dmMay.GetThanhPhanNhanCong();
                foreach (var tp in tpNC)
                {
                    var ncMay = allNC.FirstOrDefault(n => n.LoaiNhanCong == AIE.Core.Enums.LoaiNhanCong.VanHanhMay && n.Nhom == tp.Nhom);
                    if (ncMay != null) mayM.LuongTho += ncMay.GetDonGia(_vungApDung) * tp.SoLuong;
                }
            }

            _mayThiCongList.Add(mayM);
        }

        dgvVL.DataSource = new BindingSource { DataSource = _vatLieuList };
        dgvNC.DataSource = new BindingSource { DataSource = _nhanCongList };
        dgvMay.DataSource = new BindingSource { DataSource = _mayThiCongList };

        RecalculateMachineCosts();
    }

    private void PrepareData(List<VatTuGiaModel> extractedItems)
    {
        _vatLieuList = new List<DgVatLieuModel>();
        _nhanCongList = new List<DgNhanCongModel>();
        _mayThiCongList = new List<DgMayThiCongModel>();

        var db = new DatabaseManager();
        var dmMayRepo = new DinhMucCaMayRepository(db.Context);
        var mayRepo = new MayThiCongRepository(db.Context);
        var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);
        var bdgRepo = new AIE.Data.Repositories.BoDonGiaRepository(db.Context);

        if (_loadedBoId.HasValue && _loadedBoId.Value > 0)
        {
            try
            {
                var boInfo = bdgRepo.GetById(_loadedBoId.Value);
                if (boInfo != null)
                {
                    if (!string.IsNullOrWhiteSpace(boInfo.TenBo))
                        txtTenBoDonGia.Text = boInfo.TenBo;
                    if (boInfo.GiaXang > 0) txtGiaXang.Text = boInfo.GiaXang.ToString("0.####");
                    if (boInfo.GiaDiezel > 0) txtGiaDiezel.Text = boInfo.GiaDiezel.ToString("0.####");
                    if (boInfo.GiaDien > 0) txtGiaDien.Text = boInfo.GiaDien.ToString("0.####");
                }
            }
            catch { }
        }

        Dictionary<string, BoDonGiaRepository.GiaVatLieuBo> giaVLMap = null;
        Dictionary<string, decimal> giaNCMap = null;
        try
        {
            if (_loadedBoId.HasValue && _loadedBoId.Value > 0)
            {
                giaVLMap = bdgRepo.GetGiaVL(_loadedBoId.Value).ToDictionary(x => x.MaVL, x => x);
                giaNCMap = bdgRepo.GetGiaNC(_loadedBoId.Value).ToDictionary(x => x.MaNC, x => x.DonGia);
            }
        }
        catch { }

        foreach (var item in extractedItems)
        {
            if (item.LoaiHP == AIE.Core.Enums.LoaiHaoPhi.VL)
            {
                string tenVl = (item.TenVatTu ?? "").ToLower();
                string dv = (item.DonVi ?? "").Trim();
                if (dv == "%" || tenVl.Contains("vật liệu khác") || item.MaHieu == "VL_KHAC" || item.MaHieu.StartsWith("VLK"))
                    continue;

                decimal giaGoc = item.GiaChuan ?? 0;
                decimal bocXep = 0;
                decimal vcOTo = 0;
                decimal vcBo = 0;

                if (giaVLMap != null && giaVLMap.TryGetValue(item.MaHieu, out var savedVl))
                {
                    giaGoc = savedVl.GiaGoc;
                    bocXep = savedVl.ChiPhiBocXep;
                    vcOTo = savedVl.CuocVCOTo;
                    vcBo = savedVl.CuocVCBo;
                }

                _vatLieuList.Add(new DgVatLieuModel
                {
                    MaHieu = item.MaHieu,
                    Ten = item.TenVatTu,
                    DonVi = item.DonVi,
                    KhoiLuong = item.KhoiLuong,
                    GiaGoc = giaGoc,
                    ChiPhiBocXep = bocXep,
                    CuocVCOTo = vcOTo,
                    CuocVCBo = vcBo
                });
            }
            else if (item.LoaiHP == AIE.Core.Enums.LoaiHaoPhi.NC)
            {
                string tenNC = (item.TenVatTu ?? "").ToLower();
                string dv = (item.DonVi ?? "").Trim();
                if (dv == "%" || tenNC.Contains("nhân công khác"))
                    continue;

                var ncDb = ncRepo.GetAll().FirstOrDefault(x => x.MaNC == item.MaHieu);
                decimal giaNC = item.GiaChuan ?? 0;

                if (giaNCMap != null && giaNCMap.TryGetValue(item.MaHieu, out var savedNCPrice) && savedNCPrice > 0)
                {
                    giaNC = savedNCPrice;
                }
                
                int nhomNC = 1;
                if (ncDb != null)
                {
                    nhomNC = ncDb.Nhom;
                }
                else if (!string.IsNullOrEmpty(item.TenVatTu))
                {
                    string t = item.TenVatTu.ToLower();
                    if (t.Contains("nhóm 2") || t.Contains("nhóm ii")) nhomNC = 2;
                    else if (t.Contains("nhóm 3") || t.Contains("nhóm iii")) nhomNC = 3;
                    else if (t.Contains("nhóm 4") || t.Contains("nhóm iv")) nhomNC = 4;
                    else if (t.Contains("nhóm 5") || t.Contains("nhóm v")) nhomNC = 5;
                    else if (t.Contains("nhóm 6") || t.Contains("nhóm vi")) nhomNC = 6;
                    else if (t.Contains("nhóm 7") || t.Contains("nhóm vii")) nhomNC = 7;
                }

                if (giaNC == 0)
                {
                    var ncByNhom = ncRepo.GetAll().FirstOrDefault(x => x.LoaiNhanCong == AIE.Core.Enums.LoaiNhanCong.XayDung && x.Nhom == nhomNC);
                    if (ncByNhom != null)
                        giaNC = ncByNhom.GetDonGia(_vungApDung);
                }

                _nhanCongList.Add(new DgNhanCongModel
                {
                    MaHieu = item.MaHieu,
                    Ten = item.TenVatTu,
                    DonVi = item.DonVi,
                    GiaHienTruong = giaNC,
                    NhomNhanCong = nhomNC
                });
            }
            else if (item.LoaiHP == AIE.Core.Enums.LoaiHaoPhi.MAY)
            {
                string tenMay = (item.TenVatTu ?? "").ToLower();
                string dv = (item.DonVi ?? "").Trim();
                if (dv == "%" || item.MaHieu == "M7016" || tenMay.Contains("máy khác"))
                    continue;

                var dmMay = dmMayRepo.GetByMaMay(item.MaHieu);
                var mayM = new DgMayThiCongModel
                {
                    MaHieu = item.MaHieu,
                    Ten = item.TenVatTu,
                    DonVi = item.DonVi
                };
                
                if (dmMay != null)
                {
                    mayM.SoCaNam = dmMay.SoCaNam > 0 ? dmMay.SoCaNam : 250;
                    mayM.NguyenGia = dmMay.NguyenGia;
                    mayM.TyLeKhauHao = dmMay.KhauHao;
                    mayM.TyLeSuaChua = dmMay.SuaChua;
                    mayM.TyLeKhac = dmMay.ChiPhiKhac;
                    mayM.NhomNhanCong = dmMay.NhomNhanCong;
                    mayM.SoLuongNhanCong = dmMay.SoLuongNhanCong;
                    mayM.HeSoNhienLieuPhu = dmMay.HeSoNhienLieuPhu;
                    mayM.NhanCongString = dmMay.NhanCongString;
                    
                    decimal g_th = dmMay.NguyenGia >= 30000000m ? dmMay.NguyenGia * 0.1m : 0m;
                    mayM.KhauHao = ((dmMay.NguyenGia - g_th) * dmMay.KhauHao / 100m) / mayM.SoCaNam;
                    
                    mayM.SuaChua = (dmMay.NguyenGia * dmMay.SuaChua / 100m) / mayM.SoCaNam;
                    mayM.ChiPhiKhac = (dmMay.NguyenGia * dmMay.ChiPhiKhac / 100m) / mayM.SoCaNam;
                    
                    mayM.DinhMucXang = dmMay.DinhMucXang;
                    mayM.DinhMucDiezel = dmMay.DinhMucDiezel;
                    mayM.DinhMucDien = dmMay.DinhMucDien;
                    
                    mayM.LuongTho = 0;
                    var allNC = ncRepo.GetAll();
                    var tpNC = dmMay.GetThanhPhanNhanCong();
                    foreach (var tp in tpNC)
                    {
                        var ncMay = allNC.FirstOrDefault(n => n.LoaiNhanCong == AIE.Core.Enums.LoaiNhanCong.VanHanhMay && n.Nhom == tp.Nhom);
                        if (ncMay != null) mayM.LuongTho += ncMay.GetDonGia(_vungApDung) * tp.SoLuong;
                    }
                }
                _mayThiCongList.Add(mayM);
            }
        }
    }

    private void LoadDataToGrids()
    {
        if (_loadedBoId.HasValue && _loadedBoId.Value > 0)
        {
            try
            {
                var db = new DatabaseManager();
                var repo = new BoDonGiaRepository(db.Context);
                var boInfo = repo.GetById(_loadedBoId.Value);
                if (boInfo != null)
                {
                    if (!string.IsNullOrWhiteSpace(boInfo.TenBo))
                        txtTenBoDonGia.Text = boInfo.TenBo;
                    if (boInfo.GiaXang > 0) txtGiaXang.Text = boInfo.GiaXang.ToString("0.####");
                    if (boInfo.GiaDiezel > 0) txtGiaDiezel.Text = boInfo.GiaDiezel.ToString("0.####");
                    if (boInfo.GiaDien > 0) txtGiaDien.Text = boInfo.GiaDien.ToString("0.####");
                }
            }
            catch { }
        }

        if (string.IsNullOrWhiteSpace(txtGiaXang.Text) && string.IsNullOrWhiteSpace(txtGiaDiezel.Text) && string.IsNullOrWhiteSpace(txtGiaDien.Text))
        {
            try
            {
                var settingsDir = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "AIE_DuToan");
                var fuelFile = System.IO.Path.Combine(settingsDir, "fuel_prices.json");
                if (System.IO.File.Exists(fuelFile))
                {
                    var json = System.IO.File.ReadAllText(fuelFile);
                    dynamic doc = Newtonsoft.Json.JsonConvert.DeserializeObject(json);
                    decimal gx = (decimal)doc.GiaXang;
                    decimal gdz = (decimal)doc.GiaDiezel;
                    decimal gdi = (decimal)doc.GiaDien;
                    if (gx > 0) txtGiaXang.Text = gx.ToString("0.####");
                    if (gdz > 0) txtGiaDiezel.Text = gdz.ToString("0.####");
                    if (gdi > 0) txtGiaDien.Text = gdi.ToString("0.####");
                }
            }
            catch { }
        }

        if (string.IsNullOrWhiteSpace(txtGiaXang.Text)) txtGiaXang.Text = "22150";
        if (string.IsNullOrWhiteSpace(txtGiaDiezel.Text)) txtGiaDiezel.Text = "28090";
        if (string.IsNullOrWhiteSpace(txtGiaDien.Text)) txtGiaDien.Text = "2204";

        dgvVL.DataSource = new BindingSource { DataSource = _vatLieuList };
        dgvNC.DataSource = new BindingSource { DataSource = _nhanCongList };
        dgvMay.DataSource = new BindingSource { DataSource = _mayThiCongList };
        
        dgvVL?.AutoFit();
        dgvNC?.AutoFit();
        dgvMay?.AutoFit();
        RecalculateMachineCosts();
    }

    private void DgvNC_CurrentCellDirtyStateChanged(object? sender, EventArgs e)
    {
        if (dgvNC.IsCurrentCellDirty && dgvNC.CurrentCell is DataGridViewComboBoxCell)
        {
            dgvNC.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }
    }

    private void DgvNC_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
        
        if (dgvNC.Columns[e.ColumnIndex].Name == "NhomNhanCong")
        {
            var nhom = dgvNC["NhomNhanCong", e.RowIndex].Value;
            if (nhom is int n)
            {
                var db = new AIE.Data.DatabaseManager();
                var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);
                var allNC = ncRepo.GetAll();
                var ncDb = allNC.FirstOrDefault(x => x.LoaiNhanCong == AIE.Core.Enums.LoaiNhanCong.XayDung && x.Nhom == n);
                if (ncDb != null)
                {
                    _nhanCongList[e.RowIndex].GiaHienTruong = ncDb.GetDonGia(_vungApDung);
                    dgvNC.InvalidateRow(e.RowIndex);
                }
            }
        }
    }

    private void DgvNC_CellContentClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0) return;
        
        if (dgvNC.Columns[e.ColumnIndex].Name == "LuuGiaGoc")
        {
            var maNC = _nhanCongList[e.RowIndex].MaHieu;
            var currentPrice = _nhanCongList[e.RowIndex].GiaHienTruong;
            
            var db = new AIE.Data.DatabaseManager();
            var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);
            var ncDb = ncRepo.GetAll().FirstOrDefault(x => x.MaNC == maNC);
            if (ncDb != null)
            {
                ncDb.SetDonGia(_vungApDung, currentPrice);
                ncDb.NgayCapNhat = DateTime.Now;
                ncRepo.Upsert(ncDb);
                MessageBox.Show($"Đã cập nhật giá gốc cho nhân công {maNC} = {currentPrice:N0}", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Không tìm thấy mã nhân công trong cơ sở dữ liệu gốc.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private void DgvVL_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (_suppressRecalc || e.RowIndex < 0) return;
        dgvVL.InvalidateRow(e.RowIndex);
    }

    private void DgvVL_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.RowIndex >= _vatLieuList.Count) return;
        var vl = _vatLieuList[e.RowIndex];
        if (vl == null) return;

        string colName = dgvVL.Columns[e.ColumnIndex].Name;

        if (colName == "ChiPhiBocXep")
        {
            var savedCfg = AIE.ExcelAddIn.Services.BocXepStorage.GetConfig(vl.Ten);
            string? maDm = vl.MaDinhMucBocXep ?? savedCfg?.MaDinhMuc;
            int phamVi = vl.PhamViBocXep != 0 ? vl.PhamViBocXep : (savedCfg?.PhamVi ?? 0);
            decimal dmNC = vl.DmNCBocXep > 0 ? vl.DmNCBocXep : (savedCfg?.DmNC ?? 0);
            decimal dmMay = vl.DmMayBocXep > 0 ? vl.DmMayBocXep : (savedCfg?.DmMay ?? 0);

            var matchedDm = !string.IsNullOrEmpty(maDm)
                ? DinhMucBocXepDatabase.DanhSach.FirstOrDefault(x => x.MaHieu == maDm)
                : DinhMucBocXepDatabase.NhanDienDinhMuc(vl.Ten, vl.DonVi);

            if (matchedDm == null && vl.ChiPhiBocXep == 0)
            {
                var res = MessageBox.Show(
                    $"Vật liệu \"{vl.Ten}\" không thuộc danh mục có định mức bốc xếp quy định trong Chương XII (Định mức Thông tư số 38/2026/TT-BXD).\n\nBạn có muốn tự chọn một định mức bốc xếp để tính không?",
                    "Thông báo định mức bốc xếp",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);
                if (res != DialogResult.Yes) return;
            }

            // 1. Tìm đơn giá Nhân công nhóm I trong _nhanCongList hoặc truy vấn CSDL
            var ncNhom1 = _nhanCongList.FirstOrDefault(n => n.NhomNhanCong == 1);
            decimal giaNC1 = 0;
            if (ncNhom1 != null && ncNhom1.GiaHienTruong > 0)
            {
                giaNC1 = ncNhom1.GiaHienTruong;
            }
            else
            {
                var db = new DatabaseManager();
                var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);
                var ncDb = ncRepo.GetAll().FirstOrDefault(x => x.LoaiNhanCong == AIE.Core.Enums.LoaiNhanCong.XayDung && x.Nhom == 1);
                if (ncDb != null)
                {
                    giaNC1 = ncDb.GetDonGia(_vungApDung);
                    
                    // Tự động bổ sung vào tab Nhân công để người dùng quản lý
                    var ncMoi = new DgNhanCongModel
                    {
                        MaHieu = ncDb.MaNC,
                        Ten = ncDb.TenNC,
                        DonVi = "công",
                        GiaHienTruong = giaNC1,
                        NhomNhanCong = 1
                    };
                    _nhanCongList.Add(ncMoi);
                    dgvNC.DataSource = new BindingSource { DataSource = _nhanCongList };
                    dgvNC.Refresh();
                }
            }

            // Dọn dẹp dòng M103.0101 cũ nếu có do phiên bản trước chèn nhầm
            var oldWrongCrane = _mayThiCongList.FirstOrDefault(m => m.MaHieu == "M103.0101");
            if (oldWrongCrane != null)
            {
                _mayThiCongList.Remove(oldWrongCrane);
                dgvMay.DataSource = new BindingSource { DataSource = _mayThiCongList };
                dgvMay.Refresh();
            }

            // Đảm bảo tính toán lại giá máy thi công mới nhất
            RecalculateMachineCosts();

            // 2. Đơn giá máy thi công (Cần cẩu bánh hơi 6T - M102.0201)
            decimal giaMay = LayDonGiaMayThiCong("M102.0201", "Cần cẩu bánh hơi - sức nâng: 6 T");

            // 3. Mở modal tính bốc xếp
            dgvVL.EndEdit();
            using var frm = new TinhChiPhiBocXepForm(
                vl.Ten, 
                vl.DonVi, 
                giaNC1, 
                giaMay, 
                vl.ChiPhiBocXep,
                maDm,
                phamVi,
                dmNC,
                dmMay);

            if (frm.ShowDialog() == DialogResult.OK)
            {
                vl.ChiPhiBocXep = frm.KetQuaChiPhiBocXep;
                vl.MaDinhMucBocXep = frm.SelectedDinhMuc?.MaHieu;
                vl.PhamViBocXep = frm.SelectedPhamVi;
                vl.DmNCBocXep = frm.DmNC;
                vl.DmMayBocXep = frm.DmMay;
                vl.MaMayBocXep = frm.MaMay;

                AIE.ExcelAddIn.Services.BocXepStorage.SaveConfig(
                    vl.Ten, 
                    vl.MaDinhMucBocXep, 
                    vl.PhamViBocXep, 
                    vl.DmNCBocXep, 
                    vl.DmMayBocXep, 
                    vl.MaMayBocXep, 
                    vl.ChiPhiBocXep);

                dgvVL.EndEdit();
                dgvVL[e.ColumnIndex, e.RowIndex].Value = frm.KetQuaChiPhiBocXep;
                dgvVL.InvalidateRow(e.RowIndex);
                dgvVL.Refresh();
            }
        }
        else if (colName == "CuocVCOTo")
        {
            var savedCfg = AIE.ExcelAddIn.Services.VanChuyenStorage.GetConfigOTo(vl.Ten);
            string? maDm = vl.MaDinhMucVCOTo ?? savedCfg?.MaDinhMuc;
            var matchedDm = !string.IsNullOrEmpty(maDm) 
                ? DinhMucVanChuyenDatabase.DanhSachOTo.FirstOrDefault(x => x.MaHieu == maDm) 
                : DinhMucVanChuyenDatabase.NhanDienOTo(vl.Ten);

            if (matchedDm == null && vl.CuocVCOTo == 0)
            {
                var res = MessageBox.Show(
                    $"Vật liệu \"{vl.Ten}\" không thuộc danh mục có định mức vận chuyển ô tô quy định trong Chương XII (Định mức Thông tư 12/2021/TT-BXD & TT 38/2026/TT-BXD).\n\nBạn có muốn tự chọn một định mức vận chuyển ô tô để tính không?",
                    "Thông báo định mức vận chuyển ô tô",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);
                if (res != DialogResult.Yes) return;
            }

            string maMay = matchedDm?.MaMay ?? vl.MaMayVCOTo ?? savedCfg?.MaMay ?? "M106.0203";
            string tenMay = matchedDm?.TenMay ?? "Ô tô tự đổ - trọng tải: 7 T";
            decimal giaMay = LayDonGiaMayThiCong(maMay, tenMay);

            dgvVL.EndEdit();
            using var frm = new TinhCuocVCOToForm(
                vl.Ten,
                vl.DonVi,
                giaMay,
                vl.CuocVCOTo,
                maDm,
                maMay,
                (m, t) => LayDonGiaMayThiCong(m, t));

            if (frm.ShowDialog(this) == DialogResult.OK)
            {
                vl.CuocVCOTo = frm.KetQuaCuocOTo;
                vl.MaDinhMucVCOTo = frm.SelectedDinhMuc?.MaHieu;
                vl.MaMayVCOTo = frm.MaMay;

                dgvVL.EndEdit();
                dgvVL[e.ColumnIndex, e.RowIndex].Value = frm.KetQuaCuocOTo;
                dgvVL.InvalidateRow(e.RowIndex);
                dgvVL.Refresh();
            }
        }
        else if (colName == "CuocVCBo")
        {
            var savedCfg = AIE.ExcelAddIn.Services.VanChuyenStorage.GetConfigBo(vl.Ten);
            string? maDm = vl.MaDinhMucVCBo ?? savedCfg?.MaDinhMuc;
            var matchedDm = !string.IsNullOrEmpty(maDm) 
                ? DinhMucVanChuyenDatabase.DanhSachBo.FirstOrDefault(x => x.MaHieu == maDm) 
                : DinhMucVanChuyenDatabase.NhanDienBo(vl.Ten);

            if (matchedDm == null && vl.CuocVCBo == 0)
            {
                var res = MessageBox.Show(
                    $"Vật liệu \"{vl.Ten}\" không thuộc danh mục có định mức vận chuyển bộ quy định trong Chương XII (Định mức Thông tư 12/2021/TT-BXD & TT 38/2026/TT-BXD).\n\nBạn có muốn tự chọn một định mức vận chuyển bộ để tính không?",
                    "Thông báo định mức vận chuyển bộ",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);
                if (res != DialogResult.Yes) return;
            }

            var ncNhom1 = _nhanCongList.FirstOrDefault(n => 
                n.NhomNhanCong == 1 || n.Ten.ToLower().Contains("nhóm 1") || n.Ten.ToLower().Contains("nhóm i") || n.MaHieu.EndsWith(".01"));
            decimal giaNC = ncNhom1 != null && ncNhom1.GiaHienTruong > 0 ? ncNhom1.GiaHienTruong : 254498m;

            dgvVL.EndEdit();
            using var frm = new TinhCuocVCBoForm(
                vl.Ten,
                vl.DonVi,
                giaNC,
                vl.CuocVCBo,
                maDm,
                ncNhom1?.MaHieu);

            if (frm.ShowDialog(this) == DialogResult.OK)
            {
                vl.CuocVCBo = frm.KetQuaCuocBo;
                vl.MaDinhMucVCBo = frm.SelectedDinhMuc?.MaHieu;

                dgvVL.EndEdit();
                dgvVL[e.ColumnIndex, e.RowIndex].Value = frm.KetQuaCuocBo;
                dgvVL.InvalidateRow(e.RowIndex);
                dgvVL.Refresh();
            }
        }
    }

    private decimal LayDonGiaMayThiCong(string maMay, string tenMayDefault)
    {
        RecalculateMachineCosts();
        var may = _mayThiCongList.FirstOrDefault(m => m.MaHieu == maMay);
        if (may != null && may.GiaHienTruong > 0) return may.GiaHienTruong;

        var db = new DatabaseManager();
        var dmMayRepo = new DinhMucCaMayRepository(db.Context);
        var dm = dmMayRepo.GetByMaMay(maMay);
        if (dm != null)
        {
            if (may == null)
            {
                may = new DgMayThiCongModel
                {
                    MaHieu = dm.MaMay,
                    Ten = tenMayDefault,
                    DonVi = "ca",
                    SoCaNam = dm.SoCaNam > 0 ? dm.SoCaNam : 240,
                    NguyenGia = dm.NguyenGia,
                    TyLeKhauHao = dm.KhauHao,
                    TyLeSuaChua = dm.SuaChua,
                    TyLeKhac = dm.ChiPhiKhac,
                    HeSoNhienLieuPhu = dm.HeSoNhienLieuPhu,
                    NhanCongString = dm.NhanCongString,
                    DinhMucXang = dm.DinhMucXang,
                    DinhMucDiezel = dm.DinhMucDiezel,
                    DinhMucDien = dm.DinhMucDien
                };

                decimal g_th = dm.NguyenGia >= 30000000m ? dm.NguyenGia * 0.1m : 0m;
                may.KhauHao = ((dm.NguyenGia - g_th) * dm.KhauHao / 100m) / may.SoCaNam;
                may.SuaChua = (dm.NguyenGia * dm.SuaChua / 100m) / may.SoCaNam;
                may.ChiPhiKhac = (dm.NguyenGia * dm.ChiPhiKhac / 100m) / may.SoCaNam;

                _mayThiCongList.Add(may);
            }
            RecalculateMachineCosts();
            dgvMay.DataSource = new BindingSource { DataSource = _mayThiCongList };
            dgvMay.Refresh();
            return may.GiaHienTruong;
        }
        return may?.GiaHienTruong ?? 0;
    }

    private void RecalculateMachineCosts()
    {
        if (_suppressRecalc) return;
        if (_mayThiCongList == null || _vatLieuList == null || _nhanCongList == null) return;
        
        decimal.TryParse(txtGiaXang.Text, out decimal gx);
        decimal.TryParse(txtGiaDiezel.Text, out decimal gdz);
        decimal.TryParse(txtGiaDien.Text, out decimal gdi);
        
        _suppressRecalc = true;
        
        if (_cachedAllNC == null || _cachedDmMay == null)
        {
            var db = new DatabaseManager();
            var ncRepo = new NhanCongRepository(db.Context);
            var dmMayRepo = new DinhMucCaMayRepository(db.Context);
            _cachedAllNC = ncRepo.GetAll().ToList();
            _cachedDmMay = dmMayRepo.GetAll().ToDictionary(x => x.MaMay, x => x);
        }
        var allNC = _cachedAllNC;
        var dmMayCache = _cachedDmMay;
        
        foreach (var m in _mayThiCongList)
        {
            m.NhienLieu = (m.DinhMucXang * gx + m.DinhMucDiezel * gdz + m.DinhMucDien * gdi) * m.HeSoNhienLieuPhu;
            
            dmMayCache.TryGetValue(m.MaHieu ?? "", out var dmMay);
            if (dmMay != null)
            {
                m.LuongTho = 0;
                var tpNC = dmMay.GetThanhPhanNhanCong();
                foreach (var tp in tpNC)
                {
                    var ncMay = allNC.FirstOrDefault(n => n.LoaiNhanCong == AIE.Core.Enums.LoaiNhanCong.VanHanhMay && n.Nhom == tp.Nhom);
                    if (ncMay != null) m.LuongTho += ncMay.GetDonGia(_vungApDung) * tp.SoLuong;
                }
            }
        }
        dgvMay.Refresh();
        _suppressRecalc = false;

        CapNhatLaiChiPhiBocXepTheoGiaMay();
    }

    private void CapNhatLaiChiPhiBocXepTheoGiaMay()
    {
        if (_suppressRecalc) return;
        if (_vatLieuList == null || _mayThiCongList == null || _nhanCongList == null) return;
        bool hasChanges = false;
        foreach (var vl in _vatLieuList)
        {
            var savedCfg = AIE.ExcelAddIn.Services.BocXepStorage.GetConfig(vl.Ten);
            string? maDm = vl.MaDinhMucBocXep ?? savedCfg?.MaDinhMuc;
            decimal dmNC = vl.DmNCBocXep > 0 ? vl.DmNCBocXep : (savedCfg?.DmNC ?? 0);
            decimal dmMay = vl.DmMayBocXep > 0 ? vl.DmMayBocXep : (savedCfg?.DmMay ?? 0);

            if (!string.IsNullOrEmpty(maDm))
            {
                var dm = DinhMucBocXepDatabase.DanhSach.FirstOrDefault(x => x.MaHieu == maDm);
                if (dm == null) continue;

                decimal giaMay = 0;
                if (dm.CoMay)
                {
                    var may = _mayThiCongList.FirstOrDefault(m => 
                        m.MaHieu == dm.MaMay || 
                        m.MaHieu == "M102.0201" || 
                        (m.Ten != null && m.Ten.ToLower().Contains("cần cẩu") && (m.Ten.Contains("6 T") || m.Ten.Contains("6 t") || m.Ten.Contains("6t"))));
                    if (may != null) giaMay = may.GiaHienTruong;
                }

                var ncNhom1 = _nhanCongList.FirstOrDefault(n => 
                    n.NhomNhanCong == 1 || n.Ten.ToLower().Contains("nhóm 1") || n.Ten.ToLower().Contains("nhóm i") || n.MaHieu.EndsWith(".01"));
                decimal giaNC = ncNhom1 != null && ncNhom1.GiaHienTruong > 0 ? ncNhom1.GiaHienTruong : 254498m;

                decimal ttNC = dmNC * giaNC;
                decimal ttMay = dmMay * giaMay;
                decimal heSoQuyDoi = DinhMucBocXepDatabase.TinhHeSoQuyDoi(vl.DonVi, dm.DonViDinhMuc);
                decimal chiPhiMoi = Math.Round((ttNC + ttMay) * heSoQuyDoi, 0, MidpointRounding.AwayFromZero);

                if (chiPhiMoi > 0 && vl.ChiPhiBocXep != chiPhiMoi)
                {
                    vl.ChiPhiBocXep = chiPhiMoi;
                    vl.MaDinhMucBocXep = maDm;
                    vl.DmNCBocXep = dmNC;
                    vl.DmMayBocXep = dmMay;
                    hasChanges = true;
                }
                else if (vl.ChiPhiBocXep == 0 && savedCfg != null && savedCfg.ChiPhiBocXep > 0)
                {
                    vl.ChiPhiBocXep = savedCfg.ChiPhiBocXep;
                    vl.MaDinhMucBocXep = maDm;
                    hasChanges = true;
                }
            }
            else if (vl.ChiPhiBocXep == 0 && savedCfg != null && savedCfg.ChiPhiBocXep > 0)
            {
                vl.ChiPhiBocXep = savedCfg.ChiPhiBocXep;
                vl.MaDinhMucBocXep = savedCfg.MaDinhMuc;
                hasChanges = true;
            }

            // Cập nhật lại CuocVCOTo nếu có cấu hình ô tô
            var savedCfgOTo = AIE.ExcelAddIn.Services.VanChuyenStorage.GetConfigOTo(vl.Ten);
            string? maDmOTo = vl.MaDinhMucVCOTo ?? savedCfgOTo?.MaDinhMuc;
            string? maMayOTo = vl.MaMayVCOTo ?? savedCfgOTo?.MaMay;
            if (!string.IsNullOrEmpty(maDmOTo) && savedCfgOTo != null && savedCfgOTo.CungDuongs != null && savedCfgOTo.CungDuongs.Count > 0)
            {
                var dmOTo = DinhMucVanChuyenDatabase.DanhSachOTo.FirstOrDefault(x => x.MaHieu == maDmOTo);
                if (dmOTo != null)
                {
                    string mm = maMayOTo ?? dmOTo.MaMay;
                    var mayXe = _mayThiCongList.FirstOrDefault(m => m.MaHieu == mm);
                    decimal giaMayXe = (mayXe != null && mayXe.GiaHienTruong > 0) ? mayXe.GiaHienTruong : (savedCfgOTo.DonGiaCaMay > 0 ? savedCfgOTo.DonGiaCaMay : 0);
                    if (giaMayXe > 0)
                    {
                        var (caXe, _, _, _, _) = DinhMucVanChuyenDatabase.TinhHaoPhiCaXeOTo(dmOTo, savedCfgOTo.CungDuongs);
                        decimal gia1Dm = (caXe * giaMayXe) / 10m;
                        decimal heSoQuyDoi = DinhMucVanChuyenDatabase.TinhHeSoQuyDoiOTo(vl.DonVi, dmOTo.DonViDinhMuc);
                        decimal cuocMoi = Math.Round(gia1Dm * heSoQuyDoi * 10m, 0, MidpointRounding.AwayFromZero);
                        if (cuocMoi > 0 && vl.CuocVCOTo != cuocMoi)
                        {
                            vl.CuocVCOTo = cuocMoi;
                            vl.MaDinhMucVCOTo = maDmOTo;
                            vl.MaMayVCOTo = mm;
                            hasChanges = true;
                        }
                    }
                    else if (vl.CuocVCOTo == 0 && savedCfgOTo.CuocVCOTo > 0)
                    {
                        vl.CuocVCOTo = savedCfgOTo.CuocVCOTo;
                        vl.MaDinhMucVCOTo = maDmOTo;
                        vl.MaMayVCOTo = mm;
                        hasChanges = true;
                    }
                }
                else if (vl.CuocVCOTo == 0 && savedCfgOTo.CuocVCOTo > 0)
                {
                    vl.CuocVCOTo = savedCfgOTo.CuocVCOTo;
                    hasChanges = true;
                }
            }
            else if (vl.CuocVCOTo == 0 && savedCfgOTo != null && savedCfgOTo.CuocVCOTo > 0)
            {
                vl.CuocVCOTo = savedCfgOTo.CuocVCOTo;
                vl.MaDinhMucVCOTo = savedCfgOTo.MaDinhMuc;
                vl.MaMayVCOTo = savedCfgOTo.MaMay;
                hasChanges = true;
            }

            // Cập nhật lại CuocVCBo nếu có cấu hình bộ
            var savedCfgBo = AIE.ExcelAddIn.Services.VanChuyenStorage.GetConfigBo(vl.Ten);
            string? maDmBo = vl.MaDinhMucVCBo ?? savedCfgBo?.MaDinhMuc;
            if (!string.IsNullOrEmpty(maDmBo) && savedCfgBo != null)
            {
                var dmBo = DinhMucVanChuyenDatabase.DanhSachBo.FirstOrDefault(x => x.MaHieu == maDmBo);
                if (dmBo != null)
                {
                    var ncNhom1 = _nhanCongList.FirstOrDefault(n => 
                        n.NhomNhanCong == 1 || 
                        (n.Ten.ToLower().Contains("nhóm 1") || n.Ten.ToLower().Contains("nhóm i") || n.MaHieu.EndsWith(".01")));
                    decimal giaNC = ncNhom1 != null && ncNhom1.GiaHienTruong > 0 ? ncNhom1.GiaHienTruong : 254498m;

                    decimal cong = DinhMucVanChuyenDatabase.TinhHaoPhiNhanCongBo(dmBo, savedCfgBo.CuLyMet, savedCfgBo.HeSoDiaHinh, savedCfgBo.SoTang);
                    decimal heSoQuyDoi = DinhMucVanChuyenDatabase.TinhHeSoQuyDoiBo(vl.DonVi, dmBo.DonViDinhMuc);
                    decimal cuocMoi = Math.Round(cong * giaNC * heSoQuyDoi, 0, MidpointRounding.AwayFromZero);
                    if (cuocMoi > 0 && vl.CuocVCBo != cuocMoi)
                    {
                        vl.CuocVCBo = cuocMoi;
                        vl.MaDinhMucVCBo = maDmBo;
                        hasChanges = true;
                    }
                    else if (vl.CuocVCBo == 0 && savedCfgBo.CuocVCBo > 0)
                    {
                        vl.CuocVCBo = savedCfgBo.CuocVCBo;
                        vl.MaDinhMucVCBo = maDmBo;
                        hasChanges = true;
                    }
                }
                else if (vl.CuocVCBo == 0 && savedCfgBo.CuocVCBo > 0)
                {
                    vl.CuocVCBo = savedCfgBo.CuocVCBo;
                    hasChanges = true;
                }
            }
            else if (vl.CuocVCBo == 0 && savedCfgBo != null && savedCfgBo.CuocVCBo > 0)
            {
                vl.CuocVCBo = savedCfgBo.CuocVCBo;
                vl.MaDinhMucVCBo = savedCfgBo.MaDinhMuc;
                hasChanges = true;
            }
        }
        if (hasChanges)
        {
            dgvVL.Refresh();
        }
    }

    private void DgvMay_CellValueChanged(object sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
        
        var colName = dgvMay.Columns[e.ColumnIndex].Name;
        if (colName == "KhauHao" || colName == "SuaChua" || colName == "ChiPhiKhac" || colName == "NguyenGia")
        {
            var m = _mayThiCongList[e.RowIndex];
            var db = new DatabaseManager();
            var dmMayRepo = new DinhMucCaMayRepository(db.Context);
            var dmMay = dmMayRepo.GetByMaMay(m.MaHieu);
            
            if (dmMay != null)
            {
                dmMay.NguyenGia = m.NguyenGia;
                dmMay.KhauHao = m.TyLeKhauHao;
                dmMay.SuaChua = m.TyLeSuaChua;
                dmMay.ChiPhiKhac = (decimal)m.TyLeKhac;
                dmMay.NguyenGia = m.NguyenGia;
                dmMayRepo.Upsert(dmMay);
                RecalculateMachineCosts();
            }
            decimal g_th = m.NguyenGia >= 30000000m ? m.NguyenGia * 0.1m : 0m;
            m.KhauHao = ((m.NguyenGia - g_th) * m.TyLeKhauHao / 100m) / m.SoCaNam;
            m.SuaChua = (m.NguyenGia * m.TyLeSuaChua / 100m) / m.SoCaNam;
            m.ChiPhiKhac = (m.NguyenGia * m.TyLeKhac / 100m) / m.SoCaNam;
            
            dgvMay.InvalidateRow(e.RowIndex);
        }
    }

    private void BtnApGiaTuBoCu_Click(object? sender, EventArgs e)
    {
        var db = new AIE.Data.DatabaseManager();
        var repo = new AIE.Data.Repositories.BoDonGiaRepository(db.Context);
        
        using (var form = new ChonBoDonGiaForm(repo))
        {
            form.Text = "Chọn Bộ Đơn Giá để Áp Giá";
            if (form.ShowDialog() == DialogResult.OK && form.SelectedBoDonGiaId > 0)
            {
                int boId = form.SelectedBoDonGiaId;
                int updatedVl = 0, updatedNc = 0, updatedMay = 0;

                // 1. Load Vật Liệu
                var giaVL = repo.GetGiaVL(boId).ToDictionary(x => x.MaVL, x => x);
                foreach (var vl in _vatLieuList)
                {
                    if (giaVL.TryGetValue(vl.MaHieu, out var gia))
                    {
                        vl.GiaGoc = gia.GiaGoc;
                        vl.ChiPhiBocXep = gia.ChiPhiBocXep;
                        vl.CuocVCOTo = gia.CuocVCOTo;
                        vl.CuocVCBo = gia.CuocVCBo;
                        if (vl.ChiPhiBocXep == 0 && vl.CuocVCOTo == 0 && vl.CuocVCBo == 0 && gia.CuocVC > 0)
                        {
                            vl.CuocVCOTo = gia.CuocVC;
                        }
                        updatedVl++;
                    }
                }

                // 2. Load Nhân Công
                var giaNC = repo.GetGiaNC(boId).ToDictionary(x => x.MaNC, x => x);
                foreach (var nc in _nhanCongList)
                {
                    if (giaNC.TryGetValue(nc.MaHieu, out var gia))
                    {
                        if (nc.GiaHienTruong != gia.DonGia)
                        {
                            nc.GiaHienTruong = gia.DonGia;
                            updatedNc++;
                        }
                    }
                }

                // 3. Load Máy Thi Công
                var boInfo = repo.GetById(boId);
                if (boInfo != null)
                {
                    txtGiaXang.Text = boInfo.GiaXang > 0 ? boInfo.GiaXang.ToString("0.####") : "";
                    txtGiaDiezel.Text = boInfo.GiaDiezel > 0 ? boInfo.GiaDiezel.ToString("0.####") : "";
                    txtGiaDien.Text = boInfo.GiaDien > 0 ? boInfo.GiaDien.ToString("0.####") : "";
                }

                var giaMay = repo.GetGiaMay(boId).ToDictionary(x => x.MaMay, x => x);
                foreach (var m in _mayThiCongList)
                {
                    if (giaMay.TryGetValue(m.MaHieu, out var gm))
                    {
                        m.DonGiaSaved = gm.DonGia;
                    }
                }
                updatedMay = _mayThiCongList.Count;

                dgvVL.Invalidate();
                dgvNC.Invalidate();
                dgvMay.Invalidate();
                
                RecalculateMachineCosts();

                MessageBox.Show($"Đã áp giá thành công!\n- Cập nhật {updatedVl} Vật liệu.\n- Cập nhật {updatedNc} Nhân công.\n- Cập nhật giá nhiên liệu cho Máy thi công.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    private void BtnLuu_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtTenBoDonGia.Text))
        {
            MessageBox.Show("Vui lòng nhập tên Bộ đơn giá", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            dgvVL.EndEdit();
            dgvNC.EndEdit();
            dgvMay.EndEdit();

            var db = new DatabaseManager();
            var repo = new BoDonGiaRepository(db.Context);
            
            decimal.TryParse(txtGiaXang.Text, out decimal gx);
            decimal.TryParse(txtGiaDiezel.Text, out decimal gdz);
            decimal.TryParse(txtGiaDien.Text, out decimal gdi);

            int id;
            if (_loadedBoId.HasValue && _loadedBoId.Value > 0)
            {
                id = _loadedBoId.Value;
                repo.UpdateInfo(id, txtTenBoDonGia.Text, gx, gdz, gdi);
            }
            else
            {
                // 1. Create new BoDonGia
                id = repo.Create(txtTenBoDonGia.Text, gx, gdz, gdi, "Thẩm định từ Excel");
            }

            // 2. Save VL
            foreach (var vl in _vatLieuList)
                repo.SaveGiaVL(id, vl.MaHieu, vl.GiaGoc, vl.CuocVC, vl.GiaHienTruong, vl.ChiPhiBocXep, vl.CuocVCOTo, vl.CuocVCBo);

            // 3. Save NC
            foreach (var nc in _nhanCongList)
                repo.SaveGiaNC(id, nc.MaHieu, nc.GiaHienTruong);

            // 4. Save May
            foreach (var m in _mayThiCongList)
                repo.SaveGiaMay(id, m.MaHieu, m.GiaHienTruong);

            try
            {
                var settingsDir = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "AIE_DuToan");
                if (!System.IO.Directory.Exists(settingsDir)) System.IO.Directory.CreateDirectory(settingsDir);
                var fuelFile = System.IO.Path.Combine(settingsDir, "fuel_prices.json");
                var fuelData = new { GiaXang = gx, GiaDiezel = gdz, GiaDien = gdi };
                System.IO.File.WriteAllText(fuelFile, Newtonsoft.Json.JsonConvert.SerializeObject(fuelData));
            }
            catch { }

            this.SavedBoDonGiaId = id;
            this.DialogResult = DialogResult.OK;
            MessageBox.Show($"Đã lưu thành công bộ đơn giá: \"{txtTenBoDonGia.Text}\"!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show("Lỗi lưu đơn giá: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // Nút cập nhật gốc đã bị gỡ bỏ khỏi form này theo yêu cầu trước đó, 
    // nhưng nếu vẫn còn logic cũ thì ta có thể giữ lại hàm này hoặc xóa đi.
    // Tôi sẽ xóa hàm BtnCapNhatGoc_Click vì btnCapNhatGoc đã bị xóa.
    
    public void SetTenBoDonGia(string ten, decimal xang, decimal diezel, decimal dien)
    {
        txtTenBoDonGia.Text = ten;
        txtGiaXang.Text = xang > 0 ? xang.ToString("0.####") : "";
        txtGiaDiezel.Text = diezel > 0 ? diezel.ToString("0.####") : "";
        txtGiaDien.Text = dien > 0 ? dien.ToString("0.####") : "";
        RecalculateMachineCosts();
    }
    private void DgvMay_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
    {
        // Định mức: cols 3, 4, 5  |  Chi phí: cols 10, 11, 12
        AIE.ExcelAddIn.Helpers.GridHelper.PaintMergedHeader(sender, e, dgvMay, 3, 5, "Định mức", 10, 12, "Chi phí");
    }

    private void DgvMay_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0) return;
        var maMay = dgvMay["MaHieu", e.RowIndex].Value?.ToString();
        var tenMay = dgvMay["Ten", e.RowIndex].Value?.ToString();
        if (!string.IsNullOrEmpty(maMay))
        {
            var db = new AIE.Data.DatabaseManager();
            var mayDmRepo = new AIE.Data.Repositories.DinhMucCaMayRepository(db.Context);
            using var frm = new NhapDinhMucCaMayForm(maMay, tenMay ?? "", mayDmRepo);
            if (frm.ShowDialog() == DialogResult.OK)
            {
                var dmMay = mayDmRepo.GetByMaMay(maMay);
                var mayM = _mayThiCongList.FirstOrDefault(x => x.MaHieu == maMay);
                if (dmMay != null && mayM != null)
                {
                    mayM.SoCaNam = dmMay.SoCaNam > 0 ? dmMay.SoCaNam : 250;
                    mayM.NguyenGia = dmMay.NguyenGia;
                    mayM.TyLeKhauHao = dmMay.KhauHao;
                    mayM.TyLeSuaChua = dmMay.SuaChua;
                    mayM.TyLeKhac = dmMay.ChiPhiKhac;
                    mayM.NhomNhanCong = dmMay.NhomNhanCong;
                    mayM.SoLuongNhanCong = dmMay.SoLuongNhanCong;
                    mayM.HeSoNhienLieuPhu = dmMay.HeSoNhienLieuPhu;
                    mayM.NhanCongString = dmMay.NhanCongString;
                    mayM.DinhMucXang = dmMay.DinhMucXang;
                    mayM.DinhMucDiezel = dmMay.DinhMucDiezel;
                    mayM.DinhMucDien = dmMay.DinhMucDien;
                    
                    decimal g_th = dmMay.NguyenGia >= 30000000m ? dmMay.NguyenGia * 0.1m : 0m;
                    mayM.KhauHao = ((dmMay.NguyenGia - g_th) * dmMay.KhauHao / 100m) / mayM.SoCaNam;
                    mayM.SuaChua = (dmMay.NguyenGia * dmMay.SuaChua / 100m) / mayM.SoCaNam;
                    mayM.ChiPhiKhac = (dmMay.NguyenGia * dmMay.ChiPhiKhac / 100m) / mayM.SoCaNam;
                }
                RecalculateMachineCosts();
            }
        }
    }
}

public class DgVatLieuModel
{
    public string MaHieu { get; set; } = string.Empty;
    public string Ten { get; set; } = string.Empty;
    public string DonVi { get; set; } = string.Empty;
    public decimal KhoiLuong { get; set; }
    public decimal GiaGoc { get; set; }
    public decimal ChiPhiBocXep { get; set; }
    public decimal CuocVCOTo { get; set; }
    public decimal CuocVCBo { get; set; }

    // Lưu cấu hình bốc xếp để tự động cập nhật lại khi đơn giá máy/nhiên liệu thay đổi
    public string? MaDinhMucBocXep { get; set; }
    public int PhamViBocXep { get; set; } = 0;
    public decimal DmNCBocXep { get; set; }
    public decimal DmMayBocXep { get; set; }
    public string? MaMayBocXep { get; set; }

    // Lưu cấu hình vận chuyển ô tô và vận chuyển bộ
    public string? MaDinhMucVCOTo { get; set; }
    public string? MaMayVCOTo { get; set; }
    public string? MaDinhMucVCBo { get; set; }

    // Giữ thuộc tính CuocVC để tương thích ngược
    public decimal CuocVC
    {
        get => ChiPhiBocXep + CuocVCOTo + CuocVCBo;
        set => ChiPhiBocXep = value;
    }

    public decimal GiaHienTruong => GiaGoc + ChiPhiBocXep + CuocVCOTo + CuocVCBo;
    public decimal ThanhTien => KhoiLuong * GiaHienTruong;
}

public class DgNhanCongModel
{
    public string MaHieu { get; set; } = string.Empty;
    public string Ten { get; set; } = string.Empty;
    public string DonVi { get; set; } = string.Empty;
    public int NhomNhanCong { get; set; }
    public decimal GiaHienTruong { get; set; }
}

public class DgMayThiCongModel
{
    public string MaHieu { get; set; } = string.Empty;
    public string Ten { get; set; } = string.Empty;
    public string DonVi { get; set; } = string.Empty;
    public decimal NguyenGia { get; set; }
    public decimal KhauHao { get; set; }
    public decimal SuaChua { get; set; }
    public decimal ChiPhiKhac { get; set; }
    public decimal LuongTho { get; set; }
    public decimal NhienLieu { get; set; }
    public decimal DonGiaSaved { get; set; }
    public decimal GiaHienTruong
    {
        get
        {
            decimal sum = KhauHao + SuaChua + ChiPhiKhac + LuongTho + NhienLieu;
            return sum > 0 ? sum : DonGiaSaved;
        }
    }
    
    // Internal TT37 factors
    public decimal SoCaNam { get; set; } = 250;
    public decimal TyLeKhauHao { get; set; }
    public decimal TyLeSuaChua { get; set; }
    public decimal TyLeKhac { get; set; }
    public decimal DinhMucXang { get; set; }
    public decimal DinhMucDiezel { get; set; }
    public decimal DinhMucDien { get; set; }
    public int NhomNhanCong { get; set; }
    public decimal SoLuongNhanCong { get; set; }
    public decimal HeSoNhienLieuPhu { get; set; } = 1.0m;
    public string NhanCongString { get; set; }
    
    public string DinhMucNhienLieuDisplay
    {
        get
        {
            var parts = new System.Collections.Generic.List<string>();
            if (DinhMucXang > 0) parts.Add($"{DinhMucXang:#.##} lít xăng");
            if (DinhMucDiezel > 0) parts.Add($"{DinhMucDiezel:#.##} lít diezel");
            if (DinhMucDien > 0) parts.Add($"{DinhMucDien:#.##} kWh");
            return string.Join(" + ", parts);
        }
    }
    
    public string NhanCongVanHanhDisplay
    {
        get
        {
            if (!string.IsNullOrWhiteSpace(NhanCongString))
                return NhanCongString;

            if (SoLuongNhanCong <= 0) return "";
            string tenTho = NhomNhanCong switch
            {
                1 => "nhân công vận hành",
                2 => "lái xe",
                3 => "thủy thủ, thợ máy, thợ điện",
                4 => "máy trưởng, thuyền trưởng",
                5 => "máy trưởng tàu biển",
                6 => "thuyền trưởng, thuyền phó",
                _ => $"nhân công nhóm {NhomNhanCong}"
            };
            return $"{SoLuongNhanCong:#.##} {tenTho}";
        }
    }
}
