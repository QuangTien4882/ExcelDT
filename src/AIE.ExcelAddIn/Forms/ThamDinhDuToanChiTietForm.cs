using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using AIE.Core.Models;
using AIE.ExcelAddIn.Helpers;
using AIE.ExcelAddIn.Services;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Button = System.Windows.Forms.Button;
using ComboBox = System.Windows.Forms.ComboBox;
using Font = System.Drawing.Font;
using Label = System.Windows.Forms.Label;
using Panel = System.Windows.Forms.Panel;
using Point = System.Drawing.Point;
using TextBox = System.Windows.Forms.TextBox;

namespace AIE.ExcelAddIn.Forms;

/// <summary>
/// Biểu mẫu Giai đoạn 2: Bảng Dự Toán Chi Tiết Sau Thẩm Định (Theo Từng Hạng Mục)
/// Áp đơn giá đã thẩm định ở Giai đoạn 1 vào bảng tiên lượng khối lượng thực tế của dự toán
/// </summary>
public class ThamDinhDuToanChiTietForm : Form
{
    private readonly string _tenCongTrinh;
    private readonly int? _boDonGiaId;
    private readonly Dictionary<string, DonGiaThamDinhItem> _donGiaMap;
    private readonly List<KetQuaCongTacThamDinh> _results;
    private readonly DuToanExcelReader _reader;

    private List<HangMuc> _hangMucList = new();
    private string _currentSheetName = string.Empty;
    private bool _isLoading = true;

    // Controls
    private Label lblTieuDe;
    private Label lblThongTin;
    private ComboBox cboSheets;
    private ComboBox cboHangMucFilter;
    private TextBox txtSearch;
    private Button btnDocLai;

    private DataGridView dgvChiTiet;

    // Summary Cards
    private Label lblTongVL;
    private Label lblTongNC;
    private Label lblTongMay;
    private Label lblTongT;
    private Label lblChenhLech;

    private Button btnXuatExcel;
    private Button btnXacNhan;
    private Button btnDong;

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

    public ThamDinhDuToanChiTietForm(
        string tenCongTrinh,
        int? boDonGiaId,
        Dictionary<string, DonGiaThamDinhItem> donGiaMap,
        List<KetQuaCongTacThamDinh> results,
        string initialSheetName = "")
    {
        _tenCongTrinh = !string.IsNullOrEmpty(tenCongTrinh) ? tenCongTrinh : "Công trình xây dựng";
        _boDonGiaId = boDonGiaId;
        _donGiaMap = donGiaMap ?? new Dictionary<string, DonGiaThamDinhItem>(StringComparer.OrdinalIgnoreCase);
        _results = results ?? new List<KetQuaCongTacThamDinh>();
        _currentSheetName = initialSheetName;
        _reader = new DuToanExcelReader();

        InitializeComponent();
        LoadAvailableSheets();
        DocDuLieuTienLuong();

        FormStateHelper.Attach(this);
    }

    private void InitializeComponent()
    {
        this.Text = "Bảng Dự Toán Chi Tiết Sau Thẩm Định (Theo Từng Hạng Mục)";
        this.Size = new Size(1400, 850);
        this.MinimumSize = new Size(1150, 680);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.WindowState = FormWindowState.Maximized;
        this.ShowInTaskbar = true;
        this.MinimizeBox = true;
        this.MaximizeBox = true;
        this.Font = UIHelper.GetFont(10f);
        this.BackColor = Color.FromArgb(246, 248, 252);

        // 1. TOP HEADER PANEL
        var pnlHeader = new Panel
        {
            Dock = DockStyle.Top,
            Height = 82,
            BackColor = Color.FromArgb(0, 51, 102),
            Padding = new Padding(20, 10, 20, 10)
        };

        lblTieuDe = new Label
        {
            Text = "📋 BẢNG DỰ TOÁN CHI TIẾT SAU THẨM ĐỊNH (THEO TỪNG HẠNG MỤC)",
            Font = UIHelper.GetFont(12.5f, FontStyle.Bold),
            ForeColor = Color.White,
            AutoSize = true,
            Location = new Point(18, 12)
        };

        lblThongTin = new Label
        {
            Text = $"Dự án: {_tenCongTrinh}   |   Đang đọc dữ liệu bảng tiên lượng khối lượng...",
            Font = UIHelper.GetFont(9.5f),
            ForeColor = Color.FromArgb(210, 230, 255),
            AutoSize = true,
            Location = new Point(20, 44)
        };

        pnlHeader.Controls.Add(lblTieuDe);
        pnlHeader.Controls.Add(lblThongTin);

        // 2. TOOLBAR PANEL
        var pnlFilter = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 50,
            BackColor = Color.White,
            WrapContents = false,
            AutoScroll = true,
            Padding = new Padding(12, 8, 12, 8)
        };

        var lblSheet = new Label
        {
            Text = "📄 Sheet Tiên lượng:",
            AutoSize = true,
            Margin = new Padding(3, 7, 3, 0),
            Font = UIHelper.GetFont(9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(50, 50, 50)
        };

        cboSheets = new ComboBox
        {
            Width = 190,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Margin = new Padding(0, 4, 15, 0),
            Font = UIHelper.GetFont(9.5f)
        };
        cboSheets.SelectedIndexChanged += (s, e) =>
        {
            if (!_isLoading && cboSheets.SelectedItem != null)
            {
                _currentSheetName = cboSheets.SelectedItem.ToString();
                DocDuLieuTienLuong();
            }
        };

        var lblHM = new Label
        {
            Text = "Hạng mục:",
            AutoSize = true,
            Margin = new Padding(3, 7, 3, 0),
            Font = UIHelper.GetFont(9.5f),
            ForeColor = Color.FromArgb(50, 50, 50)
        };

        cboHangMucFilter = new ComboBox
        {
            Width = 220,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Margin = new Padding(0, 4, 15, 0),
            Font = UIHelper.GetFont(9.5f)
        };
        cboHangMucFilter.SelectedIndexChanged += (s, e) => NapDuLieuLenGrid();

        var lblSearch = new Label
        {
            Text = "🔍 Tìm kiếm:",
            AutoSize = true,
            Margin = new Padding(3, 7, 3, 0),
            Font = UIHelper.GetFont(9.5f),
            ForeColor = Color.FromArgb(50, 50, 50)
        };

        txtSearch = new TextBox
        {
            Width = 160,
            Margin = new Padding(0, 4, 15, 0),
            Font = UIHelper.GetFont(9.5f)
        };
        txtSearch.TextChanged += (s, e) => NapDuLieuLenGrid();

        btnDocLai = new Button
        {
            Text = "⚡ Đọc lại Sheet",
            AutoSize = true,
            Height = 32,
            Padding = new Padding(12, 0, 12, 0),
            Margin = new Padding(0, 3, 10, 0),
            BackColor = Color.FromArgb(235, 245, 255),
            ForeColor = Color.FromArgb(0, 102, 204),
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnDocLai.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 215);
        btnDocLai.Click += (s, e) => DocDuLieuTienLuong();

        pnlFilter.Controls.Add(lblSheet);
        pnlFilter.Controls.Add(cboSheets);
        pnlFilter.Controls.Add(lblHM);
        pnlFilter.Controls.Add(cboHangMucFilter);
        pnlFilter.Controls.Add(lblSearch);
        pnlFilter.Controls.Add(txtSearch);
        pnlFilter.Controls.Add(btnDocLai);

        // 3. BOTTOM PANEL: TỔNG HỢP VÀ NÚT ĐIỀU HƯỚNG
        var pnlBottom = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 130,
            BackColor = Color.White,
            Padding = new Padding(15, 8, 15, 8)
        };

        int cardY = 16;
        int cardH = 58;
        int cardW = 245;

        var cardVL = TaoCardThongKe("CHI PHÍ VẬT LIỆU (VL_TĐ)", out lblTongVL, new Point(15, cardY), cardW, cardH, Color.FromArgb(0, 102, 204));
        var cardNC = TaoCardThongKe("CHI PHÍ NHÂN CÔNG (NC_TĐ)", out lblTongNC, new Point(270, cardY), cardW, cardH, Color.FromArgb(0, 102, 204));
        var cardMay = TaoCardThongKe("CHI PHÍ MÁY TC (M_TĐ)", out lblTongMay, new Point(525, cardY), cardW, cardH, Color.FromArgb(0, 102, 204));
        var cardTong = TaoCardThongKe("TỔNG TRỰC TIẾP THẨM ĐỊNH (T_TĐ)", out lblTongT, new Point(780, cardY), 310, cardH, Color.FromArgb(0, 128, 64), isLarge: true);

        pnlBottom.Controls.Add(cardVL);
        pnlBottom.Controls.Add(cardNC);
        pnlBottom.Controls.Add(cardMay);
        pnlBottom.Controls.Add(cardTong);

        lblChenhLech = new Label
        {
            Text = "So với Dự toán gốc: 0 đ",
            AutoSize = true,
            Location = new Point(1105, 33),
            Font = UIHelper.GetFont(10f, FontStyle.Bold),
            ForeColor = Color.FromArgb(180, 50, 50)
        };
        pnlBottom.Controls.Add(lblChenhLech);

        btnXuatExcel = new Button
        {
            Text = "📄 Xuất Excel Dự Toán Chi Tiết",
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            Location = new Point(pnlBottom.Width - 780, 80),
            Width = 230,
            Height = 44,
            BackColor = Color.FromArgb(235, 245, 255),
            ForeColor = Color.FromArgb(0, 102, 204),
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9.5f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnXuatExcel.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 215);
        btnXuatExcel.Click += BtnXuatExcel_Click;

        btnXacNhan = new Button
        {
            Text = "➔  XÁC NHẬN & CHUYỂN SANG TỔNG HỢP KINH PHÍ\n(BẢNG 3.8 & THDT)",
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            Location = new Point(pnlBottom.Width - 540, 80),
            Width = 430,
            Height = 44,
            BackColor = Color.FromArgb(16, 124, 65),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = UIHelper.GetFont(9.5f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnXacNhan.FlatAppearance.BorderSize = 0;
        btnXacNhan.Click += BtnXacNhan_Click;

        btnDong = new Button
        {
            Text = "Đóng",
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            Location = new Point(pnlBottom.Width - 100, 80),
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
            btnXuatExcel.Location = new Point(pnlBottom.Width - 760, 80);
            btnXacNhan.Location = new Point(pnlBottom.Width - 520, 80);
            btnDong.Location = new Point(pnlBottom.Width - 100, 80);
        };

        pnlBottom.Controls.Add(btnXuatExcel);
        pnlBottom.Controls.Add(btnXacNhan);
        pnlBottom.Controls.Add(btnDong);

        // 4. MAIN DATAGRIDVIEW
        dgvChiTiet = new DataGridView
        {
            Dock = DockStyle.Fill,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false,
            AutoGenerateColumns = false
        };

        UIHelper.ApplyStyle(dgvChiTiet);
        dgvChiTiet.DataError += (s, e) => { e.ThrowException = false; };
        KhoiTaoCotGrid();

        this.Controls.Add(dgvChiTiet);
        this.Controls.Add(pnlBottom);
        this.Controls.Add(pnlFilter);
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
        dgvChiTiet.Columns.Clear();
        dgvChiTiet.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        dgvChiTiet.ColumnHeadersHeight = 74;
        dgvChiTiet.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
        dgvChiTiet.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        dgvChiTiet.ColumnHeadersDefaultCellStyle.Font = UIHelper.GetFont(9.5f, FontStyle.Bold);

        dgvChiTiet.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
        dgvChiTiet.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;

        // 1. STT (Frozen)
        dgvChiTiet.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colSTT",
            HeaderText = "STT",
            Width = 48,
            ReadOnly = true,
            Frozen = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
        });

        // 2. Mã số / Mã hiệu (Frozen)
        dgvChiTiet.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colMa",
            HeaderText = "Mã số",
            Width = 95,
            ReadOnly = true,
            Frozen = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter, Font = UIHelper.GetFont(9.5f, FontStyle.Bold) }
        });

        // 3. Tên công tác xây dựng (Frozen)
        dgvChiTiet.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colTen",
            HeaderText = "Tên công tác xây dựng",
            Width = 360,
            MinimumWidth = 280,
            ReadOnly = true,
            Frozen = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft, WrapMode = DataGridViewTriState.True }
        });

        // 4. ĐVT
        dgvChiTiet.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colDonVi",
            HeaderText = "Đơn vị",
            Width = 65,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
        });

        // 5. Khối lượng
        dgvChiTiet.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colKhoiLuong",
            HeaderText = "Khối lượng",
            Width = 105,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "#,##0.000", Font = UIHelper.GetFont(9.5f, FontStyle.Bold) }
        });

        // 6. Đơn giá Dự toán
        dgvChiTiet.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colDonGiaDT",
            HeaderText = "Đơn giá\nDự toán (đ)",
            Width = 110,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "#,##0" }
        });

        // 7. Đơn giá Thẩm định
        dgvChiTiet.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colDonGiaTD",
            HeaderText = "Đơn giá\nThẩm định (đ)",
            Width = 115,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "#,##0", Font = UIHelper.GetFont(9.5f, FontStyle.Bold), ForeColor = Color.FromArgb(0, 102, 204) }
        });

        // 8. Thành tiền Dự toán
        dgvChiTiet.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colThanhTienDT",
            HeaderText = "Thành tiền\nDự toán (đ)",
            Width = 135,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "#,##0" }
        });

        // 9. Thành tiền Thẩm định
        dgvChiTiet.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colThanhTienTD",
            HeaderText = "Thành tiền\nThẩm định (đ)",
            Width = 140,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "#,##0", Font = UIHelper.GetFont(9.5f, FontStyle.Bold) }
        });

        // 10. Chênh lệch
        dgvChiTiet.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colChenhLech",
            HeaderText = "Chênh lệch\n(đ)",
            Width = 115,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "#,##0" }
        });

        // 11. Trạng thái áp dụng
        dgvChiTiet.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colTrangThai",
            HeaderText = "Trạng thái\nđơn giá",
            Width = 135,
            ReadOnly = true,
            DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter, Font = UIHelper.GetFont(9f) }
        });
    }

    private void LoadAvailableSheets()
    {
        try
        {
            var app = (Application)ExcelDnaUtil.Application;
            var wb = app?.ActiveWorkbook;
            if (wb == null) return;

            cboSheets.Items.Clear();
            foreach (Worksheet sheet in wb.Worksheets)
            {
                cboSheets.Items.Add(sheet.Name);
            }

            if (!string.IsNullOrEmpty(_currentSheetName) && cboSheets.Items.Contains(_currentSheetName))
            {
                cboSheets.SelectedItem = _currentSheetName;
            }
        }
        catch { }
        finally
        {
            _isLoading = false;
        }
    }

    private void DocDuLieuTienLuong()
    {
        try
        {
            var result = _reader.ReadBangDuToanChiTiet(null, _currentSheetName);
            _currentSheetName = result.SheetName;
            _hangMucList = result.DanhSachHangMuc;

            if (cboSheets.SelectedItem == null || cboSheets.SelectedItem.ToString() != _currentSheetName)
            {
                _isLoading = true;
                cboSheets.SelectedItem = _currentSheetName;
                _isLoading = false;
            }

            // Ánh xạ đơn giá thẩm định vào từng công tác
            ApDonGiaThamDinhVaoHangMuc();

            // Cập nhật bộ lọc hạng mục
            cboHangMucFilter.Items.Clear();
            cboHangMucFilter.Items.Add("Tất cả hạng mục");
            foreach (var hm in _hangMucList)
            {
                cboHangMucFilter.Items.Add(hm.TenHangMuc);
            }
            cboHangMucFilter.SelectedIndex = 0;

            NapDuLieuLenGrid();
            CapNhatTongChiPhi();
        }
        catch (Exception ex)
        {
            MessageBox.Show("Lỗi khi đọc bảng tiên lượng dự toán: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ApDonGiaThamDinhVaoHangMuc()
    {
        foreach (var hm in _hangMucList)
        {
            if (hm.DanhSachCongTac == null) continue;
            foreach (var ct in hm.DanhSachCongTac)
            {
                string cleanMa = ct.MaHieu?.Trim().ToUpper() ?? "";

                if (_donGiaMap.TryGetValue(cleanMa, out var dgItem))
                {
                    // Đã có kết quả thẩm định ở Giai đoạn 1
                    if (dgItem.LuaChon == "Theo Dự toán")
                    {
                        // Giữ nguyên theo Dự toán
                        if (ct.DonGiaVL == 0 && ct.DonGiaNC == 0 && ct.DonGiaMay == 0 && dgItem.DonGiaTong_DT > 0)
                        {
                            ct.DonGiaVL = dgItem.DonGiaVL_DT;
                            ct.DonGiaNC = dgItem.DonGiaNC_DT;
                            ct.DonGiaMay = dgItem.DonGiaMay_DT;
                        }
                    }
                    else
                    {
                        // Áp theo Thẩm định
                        ct.DonGiaVL = dgItem.DonGiaVL_TD;
                        ct.DonGiaNC = dgItem.DonGiaNC_TD;
                        ct.DonGiaMay = dgItem.DonGiaMay_TD;
                    }
                }
                else
                {
                    // Công tác không có sai khác hoặc không nằm trong danh mục sai khác
                    var kq = _results.FirstOrDefault(r => r?.DuToan != null && string.Equals(r.DuToan.MaHieu, ct.MaHieu, StringComparison.OrdinalIgnoreCase));
                    if (kq != null && kq.DuToan?.DanhSachHaoPhi != null)
                    {
                        decimal vl = kq.DuToan.DanhSachHaoPhi.Where(h => h.Loai == AIE.Core.Enums.LoaiHaoPhi.VL).Sum(h => h.DinhMuc * h.DonGia);
                        decimal nc = kq.DuToan.DanhSachHaoPhi.Where(h => h.Loai == AIE.Core.Enums.LoaiHaoPhi.NC).Sum(h => h.DinhMuc * h.DonGia);
                        decimal may = kq.DuToan.DanhSachHaoPhi.Where(h => h.Loai == AIE.Core.Enums.LoaiHaoPhi.MAY).Sum(h => h.DinhMuc * h.DonGia);

                        if (vl > 0 || nc > 0 || may > 0)
                        {
                            ct.DonGiaVL = vl;
                            ct.DonGiaNC = nc;
                            ct.DonGiaMay = may;
                        }
                    }
                }
            }
        }
    }

    private void NapDuLieuLenGrid()
    {
        dgvChiTiet.Rows.Clear();

        string keyword = txtSearch?.Text?.Trim()?.ToLower() ?? "";
        string selectedHM = cboHangMucFilter?.SelectedItem?.ToString() ?? "Tất cả hạng mục";

        int totalCT = 0;

        foreach (var hm in _hangMucList)
        {
            if (selectedHM != "Tất cả hạng mục" && hm.TenHangMuc != selectedHM)
                continue;

            var filteredItems = hm.DanhSachCongTac.Where(ct =>
            {
                if (string.IsNullOrEmpty(keyword)) return true;
                return ct.MaHieu.ToLower().Contains(keyword) || ct.TenCongTac.ToLower().Contains(keyword);
            }).ToList();

            if (filteredItems.Count == 0 && !string.IsNullOrEmpty(keyword))
                continue;

            // 1. DÒNG TIÊU ĐỀ HẠNG MỤC (HEADER ROW)
            decimal hmTtDT = hm.DanhSachCongTac.Sum(x =>
            {
                if (_donGiaMap.TryGetValue(x.MaHieu.Trim().ToUpper(), out var dg))
                    return Math.Round(x.KhoiLuong * dg.DonGiaTong_DT, 0);
                return x.ThanhTien;
            });
            decimal hmTtTD = hm.DanhSachCongTac.Sum(x => x.ThanhTien);
            decimal hmChenhLech = hmTtTD - hmTtDT;

            int hmIdx = dgvChiTiet.Rows.Add(
                "",
                "",
                "📂 " + hm.TenHangMuc.ToUpper(),
                "",
                "",
                "",
                "",
                hmTtDT,
                hmTtTD,
                hmChenhLech,
                $"{hm.DanhSachCongTac.Count} công tác"
            );

            var hmRow = dgvChiTiet.Rows[hmIdx];
            hmRow.DefaultCellStyle.BackColor = Color.FromArgb(228, 238, 252);
            hmRow.DefaultCellStyle.Font = UIHelper.GetFont(9.5f, FontStyle.Bold);
            hmRow.DefaultCellStyle.ForeColor = Color.FromArgb(0, 51, 102);
            hmRow.MinimumHeight = 34;

            // 2. CÁC DÒNG CÔNG TÁC THUỘC HẠNG MỤC
            for (int i = 0; i < filteredItems.Count; i++)
            {
                var ct = filteredItems[i];
                totalCT++;

                decimal dgDT = ct.DonGiaTongHop;
                decimal dgTD = ct.DonGiaTongHop;
                string status = "Giữ theo Dự toán";

                if (_donGiaMap.TryGetValue(ct.MaHieu.Trim().ToUpper(), out var dgItem))
                {
                    dgDT = dgItem.DonGiaTong_DT;
                    dgTD = dgItem.DonGiaTong_TD;
                    status = dgItem.LuaChon == "Theo Dự toán" ? "Theo Dự toán" : "Đã áp ĐG TĐ";
                }

                decimal ttDT = Math.Round(ct.KhoiLuong * dgDT, 0);
                decimal ttTD = Math.Round(ct.KhoiLuong * dgTD, 0);
                decimal diff = ttTD - ttDT;

                int ctIdx = dgvChiTiet.Rows.Add(
                    ct.STT > 0 ? (object)ct.STT : (i + 1),
                    ct.MaHieu,
                    ct.TenCongTac,
                    ct.DonVi,
                    ct.KhoiLuong,
                    dgDT,
                    dgTD,
                    ttDT,
                    ttTD,
                    diff,
                    status
                );

                var ctRow = dgvChiTiet.Rows[ctIdx];
                ctRow.MinimumHeight = 30;

                if (diff != 0)
                {
                    ctRow.Cells["colChenhLech"].Style.ForeColor = diff < 0 ? Color.FromArgb(0, 128, 64) : Color.FromArgb(190, 40, 40);
                    ctRow.Cells["colChenhLech"].Style.Font = UIHelper.GetFont(9.5f, FontStyle.Bold);
                    ctRow.Cells["colTrangThai"].Style.ForeColor = Color.FromArgb(0, 102, 204);
                    ctRow.Cells["colTrangThai"].Style.Font = UIHelper.GetFont(9f, FontStyle.Bold);
                }
            }
        }

        dgvChiTiet.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders);

        int countHM = _hangMucList.Count;
        lblThongTin.Text = $"Dự án: {_tenCongTrinh}   |   Sheet: '{_currentSheetName}'   |   Số hạng mục: {countHM}   |   Tổng số công tác: {totalCT}";
    }

    private void CapNhatTongChiPhi()
    {
        decimal vlTD = 0m, ncTD = 0m, mayTD = 0m;
        decimal vlDT = 0m, ncDT = 0m, mayDT = 0m;

        foreach (var hm in _hangMucList)
        {
            foreach (var ct in hm.DanhSachCongTac)
            {
                decimal dgVlTD = ct.DonGiaVL;
                decimal dgNcTD = ct.DonGiaNC;
                decimal dgMayTD = ct.DonGiaMay;

                decimal dgVlDT = ct.DonGiaVL;
                decimal dgNcDT = ct.DonGiaNC;
                decimal dgMayDT = ct.DonGiaMay;

                if (_donGiaMap.TryGetValue(ct.MaHieu.Trim().ToUpper(), out var dgItem))
                {
                    dgVlDT = dgItem.DonGiaVL_DT;
                    dgNcDT = dgItem.DonGiaNC_DT;
                    dgMayDT = dgItem.DonGiaMay_DT;

                    if (dgItem.LuaChon != "Theo Dự toán")
                    {
                        dgVlTD = dgItem.DonGiaVL_TD;
                        dgNcTD = dgItem.DonGiaNC_TD;
                        dgMayTD = dgItem.DonGiaMay_TD;
                    }
                }

                vlTD += Math.Round(ct.KhoiLuong * dgVlTD, 0);
                ncTD += Math.Round(ct.KhoiLuong * dgNcTD, 0);
                mayTD += Math.Round(ct.KhoiLuong * dgMayTD, 0);

                vlDT += Math.Round(ct.KhoiLuong * dgVlDT, 0);
                ncDT += Math.Round(ct.KhoiLuong * dgNcDT, 0);
                mayDT += Math.Round(ct.KhoiLuong * dgMayDT, 0);
            }
        }

        decimal tongTTD = vlTD + ncTD + mayTD;
        decimal tongTDT = vlDT + ncDT + mayDT;
        decimal diff = tongTTD - tongTDT;

        lblTongVL.Text = UIHelper.FormatTien(vlTD) + " đ";
        lblTongNC.Text = UIHelper.FormatTien(ncTD) + " đ";
        lblTongMay.Text = UIHelper.FormatTien(mayTD) + " đ";
        lblTongT.Text = UIHelper.FormatTien(tongTTD) + " đ";

        string diffStr = (diff >= 0 ? "+" : "") + UIHelper.FormatTien(diff) + " đ";
        lblChenhLech.Text = $"So với Dự toán gốc: {diffStr}";
        lblChenhLech.ForeColor = diff < 0 ? Color.FromArgb(0, 128, 64) : (diff > 0 ? Color.FromArgb(190, 40, 40) : Color.FromArgb(70, 70, 70));
    }

    private void BtnXacNhan_Click(object sender, EventArgs e)
    {
        try
        {
            if (_hangMucList.Count == 0 || _hangMucList.Sum(h => h.DanhSachCongTac.Count) == 0)
            {
                MessageBox.Show("Không có dữ liệu công tác nào để xuất tổng hợp kinh phí!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Khởi tạo DuToan thẩm định hoàn chỉnh
            string loaiCT = _hangMucList.FirstOrDefault()?.LoaiCongTrinh ?? "Công trình Dân dụng";

            var duToanThamDinh = new DuToan
            {
                TenCongTrinh = _tenCongTrinh,
                DiaDiem = "",
                ChuDauTu = "Ban Quản lý dự án",
                LoaiCongTrinh = loaiCT,
                CapCongTrinh = "Cấp III",
                SoBuocThietKe = 2,
                BoDonGiaId = _boDonGiaId,
                DanhSachHangMuc = _hangMucList
            };

            var formTongHop = new TongHopKinhPhiForm(duToanThamDinh, macDinhTongMucDauTu: false, isFromThamDinh: true);
            formTongHop.Show();
        }
        catch (Exception ex)
        {
            MessageBox.Show("Lỗi khi chuyển sang Bảng Tổng Hợp Kinh Phí: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void BtnXuatExcel_Click(object sender, EventArgs e)
    {
        try
        {
            var app = (Application)ExcelDnaUtil.Application;
            var wb = app?.ActiveWorkbook;
            if (wb == null)
            {
                MessageBox.Show("Không tìm thấy file Excel nào đang mở.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Worksheet ws = wb.Worksheets.Add();
            ws.Name = "DuToan_ThamDinh_" + DateTime.Now.ToString("HHmmss");

            // Viết tiêu đề
            ws.Range["A1", "K1"].Merge();
            ws.Range["A1"].Value2 = "BẢNG DỰ TOÁN CHI TIẾT SAU THẨM ĐỊNH (THEO TỪNG HẠNG MỤC)";
            ws.Range["A1"].Font.Bold = true;
            ws.Range["A1"].Font.Size = 14;
            ws.Range["A1"].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            ws.Range["A2", "K2"].Merge();
            ws.Range["A2"].Value2 = $"Công trình: {_tenCongTrinh}";
            ws.Range["A2"].Font.Bold = true;
            ws.Range["A2"].Font.Size = 11;
            ws.Range["A2"].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // Header table
            string[] headers = { "STT", "Mã hiệu", "Tên công tác xây dựng", "ĐVT", "Khối lượng", "Đơn giá DT", "Đơn giá TĐ", "Thành tiền DT", "Thành tiền TĐ", "Chênh lệch", "Ghi chú" };
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cells[4, i + 1].Value2 = headers[i];
                ws.Cells[4, i + 1].Font.Bold = true;
                ws.Cells[4, i + 1].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(230, 240, 255));
                ws.Cells[4, i + 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            }

            int r = 5;
            foreach (var hm in _hangMucList)
            {
                // Dòng Hạng mục
                ws.Cells[r, 1].Value2 = "";
                ws.Cells[r, 2].Value2 = "";
                ws.Cells[r, 3].Value2 = hm.TenHangMuc.ToUpper();
                ws.Range[ws.Cells[r, 1], ws.Cells[r, 11]].Font.Bold = true;
                ws.Range[ws.Cells[r, 1], ws.Cells[r, 11]].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(220, 235, 252));
                r++;

                foreach (var ct in hm.DanhSachCongTac)
                {
                    decimal dgDT = ct.DonGiaTongHop;
                    decimal dgTD = ct.DonGiaTongHop;
                    string status = "Giữ theo Dự toán";

                    if (_donGiaMap.TryGetValue(ct.MaHieu.Trim().ToUpper(), out var dgItem))
                    {
                        dgDT = dgItem.DonGiaTong_DT;
                        dgTD = dgItem.DonGiaTong_TD;
                        status = dgItem.LuaChon == "Theo Dự toán" ? "Theo Dự toán" : "Đã áp ĐG TĐ";
                    }

                    decimal ttDT = Math.Round(ct.KhoiLuong * dgDT, 0);
                    decimal ttTD = Math.Round(ct.KhoiLuong * dgTD, 0);

                    ws.Cells[r, 1].Value2 = ct.STT;
                    ws.Cells[r, 2].Value2 = ct.MaHieu;
                    ws.Cells[r, 3].Value2 = ct.TenCongTac;
                    ws.Cells[r, 4].Value2 = ct.DonVi;
                    ws.Cells[r, 5].Value2 = ct.KhoiLuong;
                    ws.Cells[r, 6].Value2 = dgDT;
                    ws.Cells[r, 7].Value2 = dgTD;
                    ws.Cells[r, 8].Value2 = ttDT;
                    ws.Cells[r, 9].Value2 = ttTD;
                    ws.Cells[r, 10].Value2 = ttTD - ttDT;
                    ws.Cells[r, 11].Value2 = status;

                    r++;
                }
            }

            // Kẻ khung & Định dạng số
            Range dataRange = ws.Range[ws.Cells[4, 1], ws.Cells[r - 1, 11]];
            dataRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            ws.Columns["C"].ColumnWidth = 45;
            ws.Columns["E"].NumberFormat = "#,##0.000";
            ws.Columns["F"].NumberFormat = "#,##0";
            ws.Columns["G"].NumberFormat = "#,##0";
            ws.Columns["H"].NumberFormat = "#,##0";
            ws.Columns["I"].NumberFormat = "#,##0";
            ws.Columns["J"].NumberFormat = "#,##0";

            MessageBox.Show($"Đã xuất thành công Bảng Dự toán chi tiết thẩm định sang sheet '{ws.Name}'!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show("Lỗi khi xuất Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
