using AIE.Core.Models;
using AIE.Core.Services.Shared;
using AIE.ExcelAddIn.Helpers;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace AIE.ExcelAddIn.Forms;

/// <summary>
/// Hộp thoại thiết lập thông tin hồ sơ và cấu hình xuất Báo cáo Thẩm định Dự toán
/// theo Thông tư 36/2026/TT-BXD, Thông tư 38/2026/TT-BXD và Nghị định 10/2021/NĐ-CP.
/// Hỗ trợ hiển thị Modeless, chuyển đổi Alt+Tab độc lập với Excel, tự động nhận diện
/// loại công trình và cập nhật các tỷ lệ định mức chi phí chuẩn xác.
/// </summary>
public class XuatBaoCaoThamDinhForm : Form
{
    private TextBox txtDuAn;
    private TextBox txtHangMuc;
    private TextBox txtDiaDiem;
    private TextBox txtChuDauTu;
    private TextBox txtTuVan;
    private TextBox txtThamDinh;
    private TextBox txtSoVanBan;
    private DateTimePicker dtpNgayLap;

    private ComboBox cbLoaiCongTrinh;
    private ComboBox cbCapCongTrinh;
    private ComboBox cbBuocThietKe;

    private NumericUpDown nudCPC;
    private NumericUpDown nudNhaTam;
    private NumericUpDown nudKXD;
    private NumericUpDown nudTNCTTT;
    private NumericUpDown nudVAT;
    private NumericUpDown nudQLDA;
    private NumericUpDown nudTuVan;
    private NumericUpDown nudChiPhiKhac;
    private NumericUpDown nudDuPhong;

    private CheckBox chkTongHop;
    private CheckBox chkChiPhiXD;
    private CheckBox chkDoiChieu;
    private CheckBox chkGiaVatTu;

    public ThongTinBaoCaoThamDinh ResultInfo { get; private set; }
    public ThongTinBaoCaoThamDinh ThongTinBaoCao => ResultInfo;
    public Action<ThongTinBaoCaoThamDinh> OnExportRequested { get; set; }

    public XuatBaoCaoThamDinhForm(string defaultProjectName = "")
    {
        InitializeComponent();
        FormStateHelper.Attach(this);

        if (!string.IsNullOrWhiteSpace(defaultProjectName))
        {
            txtDuAn.Text = defaultProjectName;
            AutoDetectProjectInfo(defaultProjectName);
        }
        else
        {
            CapNhatTyLeDinhMuc();
        }
    }

    private void InitializeComponent()
    {
        this.Text = "Xuất Hồ Sơ Báo Cáo Thẩm Định Dự Toán (Thông tư 36/2026/TT-BXD & TT 38/2026/TT-BXD)";
        this.Size = new Size(880, 780);
        this.MinimumSize = new Size(820, 680);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.MaximizeBox = true;
        this.MinimizeBox = true;
        this.ShowInTaskbar = true;
        this.Font = new Font("Segoe UI", 9.5f);
        this.BackColor = Color.FromArgb(248, 249, 250);

        // -------------------------------------------------------------
        // BOTTOM ACTION PANEL (Cố định ở đáy Form, luôn hiển thị nút)
        // -------------------------------------------------------------
        var pnlBottom = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 62,
            Padding = new Padding(16, 10, 16, 10),
            BackColor = Color.FromArgb(240, 244, 248)
        };

        var lblGuide = new Label
        {
            Text = "📋 Hồ sơ thẩm định chuẩn hóa theo Thông tư 36/2026/TT-BXD & Nghị định 10/2021/NĐ-CP",
            AutoSize = true,
            Font = new Font("Segoe UI", 9f, FontStyle.Italic),
            ForeColor = Color.FromArgb(70, 80, 95),
            Dock = DockStyle.Left,
            TextAlign = ContentAlignment.MiddleLeft
        };

        var flpButtons = new FlowLayoutPanel
        {
            Dock = DockStyle.Right,
            FlowDirection = FlowDirection.LeftToRight,
            AutoSize = true,
            WrapContents = false
        };

        var btnXuat = new Button
        {
            Text = "📊  Xuất Báo Cáo Excel",
            Width = 220,
            Height = 40,
            BackColor = Color.FromArgb(0, 120, 215),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10f, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Margin = new Padding(0, 0, 12, 0)
        };
        btnXuat.FlatAppearance.BorderSize = 0;
        btnXuat.Click += BtnXuat_Click;

        var btnDong = new Button
        {
            Text = "Đóng",
            Width = 100,
            Height = 40,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.White,
            Font = new Font("Segoe UI", 10f),
            Cursor = Cursors.Hand,
            Margin = new Padding(0)
        };
        btnDong.FlatAppearance.BorderColor = Color.FromArgb(200, 205, 215);
        btnDong.Click += (s, e) => this.Close();

        flpButtons.Controls.Add(btnXuat);
        flpButtons.Controls.Add(btnDong);

        pnlBottom.Controls.Add(lblGuide);
        pnlBottom.Controls.Add(flpButtons);

        // -------------------------------------------------------------
        // MAIN SCROLLABLE CONTENT PANEL
        // -------------------------------------------------------------
        var pnlMain = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            Padding = new Padding(16, 12, 16, 12)
        };

        var tlpMain = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 1,
            RowCount = 3,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(0, 0, 0, 10)
        };
        tlpMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

        // Group 1: Thông tin pháp lý dự án
        var grpInfo = new GroupBox
        {
            Text = "1. Thông tin Hành chính & Pháp lý Dự án",
            Dock = DockStyle.Top,
            Height = 220,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 70, 140),
            Padding = new Padding(14, 12, 14, 12),
            Margin = new Padding(0, 0, 0, 14)
        };

        var tlpInfo = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 4,
            RowCount = 5,
            Font = new Font("Segoe UI", 9.2f, FontStyle.Regular),
            ForeColor = Color.Black
        };
        tlpInfo.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160f));
        tlpInfo.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
        tlpInfo.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160f));
        tlpInfo.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));

        txtDuAn = new TextBox { Dock = DockStyle.Fill, Text = "Công trình xây dựng" };
        txtHangMuc = new TextBox { Dock = DockStyle.Fill, Text = "Toàn bộ công trình" };
        txtDiaDiem = new TextBox { Dock = DockStyle.Fill, Text = "" };
        txtChuDauTu = new TextBox { Dock = DockStyle.Fill, Text = "Ban Quản lý dự án" };
        txtTuVan = new TextBox { Dock = DockStyle.Fill, Text = "Công ty Cổ phần Tư vấn Xây dựng" };
        txtThamDinh = new TextBox { Dock = DockStyle.Fill, Text = "Cơ quan Thẩm định / Thẩm tra" };
        txtSoVanBan = new TextBox { Dock = DockStyle.Fill, Text = $"       /BC-TTD_{DateTime.Now:yyyy}" };
        dtpNgayLap = new DateTimePicker { Dock = DockStyle.Fill, Format = DateTimePickerFormat.Short, Value = DateTime.Today };

        AddTableRowSpan(tlpInfo, 0, "Tên công trình/DA:", txtDuAn, colSpan: 3);
        AddTableRowPair(tlpInfo, 1, "Hạng mục:", txtHangMuc, "Địa điểm:", txtDiaDiem);
        AddTableRowPair(tlpInfo, 2, "Chủ đầu tư:", txtChuDauTu, "Đơn vị lập dự toán:", txtTuVan);
        AddTableRowPair(tlpInfo, 3, "Đơn vị thẩm định:", txtThamDinh, "Số văn bản:", txtSoVanBan);
        AddTableRowPair(tlpInfo, 4, "Ngày lập báo cáo:", dtpNgayLap, null, null);

        grpInfo.Controls.Add(tlpInfo);
        tlpMain.Controls.Add(grpInfo, 0, 0);

        // Group 2: Phân loại công trình & Tỷ lệ chi phí theo TT 36/2026 & TT 38/2026
        var grpTyLe = new GroupBox
        {
            Text = "2. Phân loại Công trình & Định mức Tỷ lệ % Chi phí (Thông tư 36/2026 & TT 38/2026)",
            Dock = DockStyle.Top,
            Height = 265,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 70, 140),
            Padding = new Padding(14, 12, 14, 12),
            Margin = new Padding(0, 0, 0, 14)
        };

        var tlpTyLe = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 4,
            RowCount = 6,
            Font = new Font("Segoe UI", 9.2f, FontStyle.Regular),
            ForeColor = Color.Black
        };
        // Cột nhãn 200px đảm bảo 100% không bao giờ bị cắt chữ
        tlpTyLe.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200f));
        tlpTyLe.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
        tlpTyLe.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200f));
        tlpTyLe.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));

        cbLoaiCongTrinh = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
        cbLoaiCongTrinh.Items.AddRange(new object[]
        {
            "Công trình Nông nghiệp và Môi trường",
            "Công trình Giao thông",
            "Công trình Dân dụng",
            "Công trình Hạ tầng kỹ thuật",
            "Công trình Công nghiệp"
        });
        cbLoaiCongTrinh.SelectedIndex = 0;
        cbLoaiCongTrinh.SelectedIndexChanged += (s, e) => CapNhatTyLeDinhMuc();

        cbCapCongTrinh = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
        cbCapCongTrinh.Items.AddRange(new object[] { "Cấp I", "Cấp II", "Cấp III", "Cấp IV" });
        cbCapCongTrinh.SelectedIndex = 2;
        cbCapCongTrinh.SelectedIndexChanged += (s, e) => CapNhatTyLeDinhMuc();

        cbBuocThietKe = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
        cbBuocThietKe.Items.AddRange(new object[] { "Thiết kế 1 bước", "Thiết kế 2 bước", "Thiết kế 3 bước" });
        cbBuocThietKe.SelectedIndex = 1;
        cbBuocThietKe.SelectedIndexChanged += (s, e) => CapNhatTyLeDinhMuc();

        nudCPC = CreatePercentInput(6.1m);
        nudNhaTam = CreatePercentInput(2.2m);
        nudKXD = CreatePercentInput(2.0m);
        nudTNCTTT = CreatePercentInput(5.5m);
        nudVAT = CreatePercentInput(10.0m);
        nudQLDA = CreatePercentInput(3.284m);
        nudTuVan = CreatePercentInput(6.8m);
        nudChiPhiKhac = CreatePercentInput(2.5m);
        nudDuPhong = CreatePercentInput(5.0m);

        AddTableRowPair(tlpTyLe, 0, "Loại công trình:", cbLoaiCongTrinh, "Cấp công trình:", cbCapCongTrinh);
        AddTableRowPair(tlpTyLe, 1, "Bước thiết kế:", cbBuocThietKe, "Thuế suất GTGT (%):", nudVAT);
        AddTableRowPair(tlpTyLe, 2, "Chi phí chung (CPC) (%):", nudCPC, "Chi phí nhà tạm (Gnt) (%):", nudNhaTam);
        AddTableRowPair(tlpTyLe, 3, "CP không XĐ từ TK (%):", nudKXD, "Thu nhập chịu thuế (TL) (%):", nudTNCTTT);
        AddTableRowPair(tlpTyLe, 4, "Chi phí Quản lý dự án (%):", nudQLDA, "Chi phí Tư vấn ĐTXD (%):", nudTuVan);
        AddTableRowPair(tlpTyLe, 5, "Chi phí khác (%):", nudChiPhiKhac, "Chi phí dự phòng (%):", nudDuPhong);

        grpTyLe.Controls.Add(tlpTyLe);
        tlpMain.Controls.Add(grpTyLe, 0, 1);

        // Group 3: Tùy chọn Sheet báo cáo xuất ra Excel
        var grpSheets = new GroupBox
        {
            Text = "3. Tùy chọn Bảng biểu Hồ sơ Báo cáo Xuất ra Excel",
            Dock = DockStyle.Top,
            Height = 135,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 70, 140),
            Padding = new Padding(16, 12, 16, 12),
            Margin = new Padding(0, 0, 0, 10)
        };

        var tlpSheets = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 2,
            Font = new Font("Segoe UI", 9.2f, FontStyle.Regular),
            ForeColor = Color.Black
        };
        tlpSheets.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
        tlpSheets.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));

        chkTongHop = new CheckBox
        {
            Text = "1. Báo cáo thẩm định tổng hợp (Sheet BC_ThamDinh_TongHop)",
            Checked = true,
            AutoSize = true,
            Dock = DockStyle.Fill
        };

        chkChiPhiXD = new CheckBox
        {
            Text = "2. Bảng tổng hợp chi phí xây dựng (Sheet TH_ChiPhiXD_ThamDinh)",
            Checked = true,
            AutoSize = true,
            Dock = DockStyle.Fill
        };

        chkDoiChieu = new CheckBox
        {
            Text = "3. Bảng đối chiếu chi tiết từng công tác, sai lệch (Sheet DoiChieu_ChiTiet)",
            Checked = true,
            AutoSize = true,
            Dock = DockStyle.Fill
        };

        chkGiaVatTu = new CheckBox
        {
            Text = "4. Bảng đơn giá vật tư, nhân công, máy thẩm tra (Sheet TH_GiaVatTu_ThamDinh)",
            Checked = true,
            AutoSize = true,
            Dock = DockStyle.Fill
        };

        tlpSheets.Controls.Add(chkTongHop, 0, 0);
        tlpSheets.Controls.Add(chkChiPhiXD, 1, 0);
        tlpSheets.Controls.Add(chkDoiChieu, 0, 1);
        tlpSheets.Controls.Add(chkGiaVatTu, 1, 1);

        grpSheets.Controls.Add(tlpSheets);
        tlpMain.Controls.Add(grpSheets, 0, 2);

        pnlMain.Controls.Add(tlpMain);

        // THỨ TỰ THÊM VÀO FORM:
        this.Controls.Add(pnlMain);
        this.Controls.Add(pnlBottom);
        pnlBottom.SendToBack();
        pnlMain.BringToFront();

        this.AcceptButton = btnXuat;
        this.CancelButton = btnDong;
    }

    private void AutoDetectProjectInfo(string name)
    {
        string n = name.ToLower();
        if (n.Contains("muong") || n.Contains("mương") ||
            n.Contains("kenh") || n.Contains("kênh") ||
            n.Contains("ho ") || n.Contains("hồ ") || n.Contains("hồ_") ||
            n.Contains("dap") || n.Contains("đập") ||
            n.Contains("de ") || n.Contains("đê ") ||
            n.Contains("thuy loi") || n.Contains("thủy lợi") ||
            n.Contains("tram bom") || n.Contains("trạm bơm") ||
            n.Contains("tuoi") || n.Contains("tưới") || n.Contains("tieu") || n.Contains("tiêu"))
        {
            cbLoaiCongTrinh.SelectedItem = "Công trình Nông nghiệp và Môi trường";
            txtHangMuc.Text = "Kênh mương, công trình trên kênh";
        }
        else if (n.Contains("duong") || n.Contains("đường") ||
                 n.Contains("cau") || n.Contains("cầu") ||
                 n.Contains("giao thong") || n.Contains("giao thông") ||
                 n.Contains("via he") || n.Contains("vỉa hè") ||
                 n.Contains("nut giao") || n.Contains("nút giao"))
        {
            cbLoaiCongTrinh.SelectedItem = "Công trình Giao thông";
            txtHangMuc.Text = "Nền, mặt đường và công trình phụ trợ";
        }
        else if (n.Contains("ha tang") || n.Contains("hạ tầng") ||
                 n.Contains("chieu sang") || n.Contains("chiếu sáng") ||
                 n.Contains("cong vien") || n.Contains("công viên") ||
                 n.Contains("thoat nuoc") || n.Contains("thoát nước") ||
                 n.Contains("cap nuoc") || n.Contains("cấp nước") ||
                 n.Contains("cay xanh") || n.Contains("cây xanh"))
        {
            cbLoaiCongTrinh.SelectedItem = "Công trình Hạ tầng kỹ thuật";
            txtHangMuc.Text = "Hệ thống hạ tầng kỹ thuật";
        }
        else if (n.Contains("nha xuong") || n.Contains("nhà xưởng") ||
                 n.Contains("nha may") || n.Contains("nhà máy") ||
                 n.Contains("bien ap") || n.Contains("biến áp") ||
                 n.Contains("cong nghiep") || n.Contains("công nghiệp"))
        {
            cbLoaiCongTrinh.SelectedItem = "Công trình Công nghiệp";
            txtHangMuc.Text = "Hạng mục công nghệ và kết cấu";
        }
        else
        {
            cbLoaiCongTrinh.SelectedItem = "Công trình Dân dụng";
        }

        CapNhatTyLeDinhMuc();
    }

    private void CapNhatTyLeDinhMuc()
    {
        string loai = cbLoaiCongTrinh.SelectedItem?.ToString() ?? "";
        string buoc = cbBuocThietKe.SelectedItem?.ToString() ?? "Thiết kế 2 bước";

        // Nguồn định mức dùng chung: QLDA (TT 38/2026 Bảng 1.1), CPC (Bảng 3.3 TT 36/2026), TNCTTT (Bảng 3.6 TT 36/2026)
        string loaiChuan = DinhMucTT38Database.ChuanHoaLoaiCongTrinh(loai);
        decimal qldaMacDinh = DinhMucTT38Database.TiLeQLDA.TryGetValue(loaiChuan, out var qldaArr) ? qldaArr[0] : 3.2m;
        decimal cpcMacDinh = InterpolationHelper.LayTiLeCPCTMDT(loai);
        decimal tnctttMacDinh = InterpolationHelper.LayTiLeTNCTTT(loai);

        if (loai.Contains("Nông nghiệp"))
        {
            nudNhaTam.Value = 2.2m;     // Bảng 3.7 TT 36/2026: Kênh mương / Thủy lợi là công trình theo tuyến = 2.2%
            nudKXD.Value = 2.0m;        // Bảng 3.5 TT 36/2026
            nudTuVan.Value = buoc.Contains("1 bước") ? 5.5m : (buoc.Contains("3 bước") ? 7.5m : 6.8m);
        }
        else if (loai.Contains("Giao thông"))
        {
            nudNhaTam.Value = 2.2m;     // Bảng 3.7 TT 36/2026: Giao thông là công trình theo tuyến = 2.2%
            nudKXD.Value = 2.0m;
            nudTuVan.Value = buoc.Contains("1 bước") ? 5.2m : (buoc.Contains("3 bước") ? 7.2m : 6.5m);
        }
        else if (loai.Contains("Công nghiệp"))
        {
            nudNhaTam.Value = 1.2m;     // Bảng 3.7 TT 36/2026
            nudKXD.Value = 2.0m;
            nudTuVan.Value = buoc.Contains("1 bước") ? 5.2m : (buoc.Contains("3 bước") ? 7.2m : 6.5m);
        }
        else if (loai.Contains("Hạ tầng"))
        {
            nudNhaTam.Value = 1.2m;     // Bảng 3.7 TT 36/2026
            nudKXD.Value = 2.0m;
            nudTuVan.Value = buoc.Contains("1 bước") ? 5.0m : (buoc.Contains("3 bước") ? 7.0m : 6.2m);
        }
        else // Dân dụng
        {
            nudNhaTam.Value = 1.2m;     // Bảng 3.7 TT 36/2026
            nudKXD.Value = 2.0m;
            nudTuVan.Value = buoc.Contains("1 bước") ? 5.8m : (buoc.Contains("3 bước") ? 8.0m : 7.2m);
        }

        nudCPC.Value = cpcMacDinh;          // Bảng 3.3 TT 36/2026
        nudTNCTTT.Value = tnctttMacDinh;    // Bảng 3.6 TT 36/2026
        nudQLDA.Value = qldaMacDinh;        // TT 38/2026 Bảng 1.1 (mốc <= 10 tỷ)
        nudChiPhiKhac.Value = 2.5m;
        nudDuPhong.Value = 5.0m;
    }

    private void BtnXuat_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtDuAn.Text))
        {
            MessageBox.Show("Vui lòng nhập Tên dự án / Công trình.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            txtDuAn.Focus();
            return;
        }

        ResultInfo = new ThongTinBaoCaoThamDinh
        {
            TenDuAn = txtDuAn.Text.Trim(),
            TenHangMuc = txtHangMuc.Text.Trim(),
            DiaDiem = txtDiaDiem.Text.Trim(),
            ChuDauTu = txtChuDauTu.Text.Trim(),
            DonViTuVan = txtTuVan.Text.Trim(),
            DonViThamDinh = txtThamDinh.Text.Trim(),
            SoVanBan = txtSoVanBan.Text.Trim(),
            NgayLap = dtpNgayLap.Value,
            LoaiCongTrinh = cbLoaiCongTrinh.SelectedItem?.ToString() ?? "Công trình Dân dụng",
            CapCongTrinh = cbCapCongTrinh.SelectedItem?.ToString() ?? "Cấp III",
            BuocThietKe = cbBuocThietKe.SelectedItem?.ToString() ?? "Thiết kế 2 bước",

            TiLeCPC = nudCPC.Value,
            TiLeNhaTam = nudNhaTam.Value,
            TiLeKhongXacDinh = nudKXD.Value,
            TiLeTNCTTT = nudTNCTTT.Value,
            TiLeVAT = nudVAT.Value,

            TiLeQLDA = nudQLDA.Value,
            TiLeTuVan = nudTuVan.Value,
            TiLeChiPhiKhac = nudChiPhiKhac.Value,
            TiLeDuPhong = nudDuPhong.Value,

            XuatTongHop = chkTongHop.Checked,
            XuatChiPhiXD = chkChiPhiXD.Checked,
            XuatDoiChieuChiTiet = chkDoiChieu.Checked,
            XuatGiaVatTu = chkGiaVatTu.Checked
        };

        if (!ResultInfo.XuatTongHop && !ResultInfo.XuatChiPhiXD && !ResultInfo.XuatDoiChieuChiTiet && !ResultInfo.XuatGiaVatTu)
        {
            MessageBox.Show("Vui lòng chọn ít nhất 1 bảng biểu để xuất báo cáo.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        this.DialogResult = DialogResult.OK;
        if (OnExportRequested != null)
        {
            OnExportRequested(ResultInfo);
            this.Close();
        }
        else
        {
            this.Close();
        }
    }

    private NumericUpDown CreatePercentInput(decimal defaultValue)
    {
        return new NumericUpDown
        {
            Dock = DockStyle.Fill,
            DecimalPlaces = 3,
            Increment = 0.1m,
            Minimum = 0.0m,
            Maximum = 100.0m,
            Value = defaultValue,
            TextAlign = HorizontalAlignment.Right
        };
    }

    private void AddTableRowPair(TableLayoutPanel tlp, int row, string label1, Control ctrl1, string label2, Control ctrl2)
    {
        tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 36f));

        if (!string.IsNullOrEmpty(label1))
        {
            var lbl1 = new Label
            {
                Text = label1,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(0, 0, 8, 0)
            };
            tlp.Controls.Add(lbl1, 0, row);
        }

        if (ctrl1 != null)
        {
            tlp.Controls.Add(ctrl1, 1, row);
        }

        if (!string.IsNullOrEmpty(label2))
        {
            var lbl2 = new Label
            {
                Text = label2,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(0, 0, 8, 0)
            };
            tlp.Controls.Add(lbl2, 2, row);
        }

        if (ctrl2 != null)
        {
            tlp.Controls.Add(ctrl2, 3, row);
        }
    }

    private void AddTableRowSpan(TableLayoutPanel tlp, int row, string label, Control ctrl, int colSpan = 3)
    {
        tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 36f));

        var lbl = new Label
        {
            Text = label,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleRight,
            Padding = new Padding(0, 0, 8, 0)
        };
        tlp.Controls.Add(lbl, 0, row);

        if (ctrl != null)
        {
            tlp.Controls.Add(ctrl, 1, row);
            tlp.SetColumnSpan(ctrl, colSpan);
        }
    }
}
