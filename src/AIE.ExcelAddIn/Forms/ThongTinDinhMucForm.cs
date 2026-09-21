using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using AIE.Core.Models;
using AIE.Core.Services;
using AIE.ExcelAddIn.Helpers;
using Microsoft.Win32;

namespace AIE.ExcelAddIn.Forms
{
    /// <summary>
    /// Popup "Kiểm tra định mức" (Phần D): hiển thị chi tiết bảng định mức TT 38/2026/TT-BXD
    /// đang áp dụng cho một khoản mục chi phí trên Bảng Tổng hợp kinh phí (Tab 2 / Tab TMĐT).
    /// - Lưu kích thước cửa sổ giữa các lần mở (FormStateHelper)
    /// - File PDF Phụ lục VIII được lưu sẵn trong phần mềm, mở ngay đúng trang
    /// - Có công cụ tính nội suy tay theo đúng công thức của Thông tư
    /// </summary>
    public class ThongTinDinhMucForm : Form
    {
        private readonly ThongTinKiemTraDinhMuc _tt;
        private readonly ChiPhiKinhPhiItem _item;
        private readonly BangTongHopKinhPhiModel _model;
        private readonly Action _refreshGrid;

        private const string PdfFileName = "38.2026.TT-BXD-PL8.pdf";
        private static readonly string PdfStorageDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AIE_DuToan", "pdf");
        private static readonly string PdfStoredPath = Path.Combine(PdfStorageDir, PdfFileName);

        private Label _lblTenInfo;
        private Label _lblGhiChu;
        private Label _lblTyLeNoiSuy;
        private Label _lblTyLeHienTai;
        private DataGridView _dgvBangTra;
        private Button _btnApDung;
        private Button _btnMoPdf;
        private Button _btnDong;
        private Panel _pnlBottom;

        private TextBox _txtQ1, _txtN1, _txtQ2, _txtN2, _txtQ;
        private Label _lblNoiSuyResult;
        private decimal _tyLeTinhTay;
        private bool _isInit = true;

        private string _pdfPath;

        public ThongTinDinhMucForm(ThongTinKiemTraDinhMuc tt, ChiPhiKinhPhiItem item, BangTongHopKinhPhiModel model, Action refreshGrid)
        {
            _tt = tt ?? throw new ArgumentNullException(nameof(tt));
            _item = item ?? throw new ArgumentNullException(nameof(item));
            _model = model ?? throw new ArgumentNullException(nameof(model));
            _refreshGrid = refreshGrid;

            _isInit = true;
            try
            {
                InitializeComponent2();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khởi tạo giao diện Kiểm tra định mức: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            FormStateHelper.Attach(this);

            try
            {
                NapDuLieu();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi nạp dữ liệu định mức: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            _isInit = false;
            try { TinhNoiSuyTay(); } catch { }
            ReLayoutBottom();
        }

        private void InitializeComponent2()
        {
            this.Text = "Kiểm tra định mức chi phí (Thông tư 38/2026/TT-BXD)";
            this.Size = new Size(960, 960);
            this.MinimumSize = new Size(860, 780);
            this.StartPosition = FormStartPosition.CenterParent;
            this.Font = UIHelper.GetFont(10f);
            this.BackColor = Color.FromArgb(245, 246, 250);
            this.ShowIcon = false;
            this.WindowState = FormWindowState.Normal;

            // ===== HEADER =====
            var pnlHeader = new Panel { Dock = DockStyle.Top, Height = 88, BackColor = Color.White };
            pnlHeader.Paint += (s, e) =>
            {
                ControlPaint.DrawBorder(e.Graphics, pnlHeader.ClientRectangle, Color.FromArgb(220, 221, 225), ButtonBorderStyle.Solid);
            };

            var lblTieuDe = new Label
            {
                Dock = DockStyle.Top,
                Height = 28,
                Text = "KIỂM TRA ĐỊNH MỨC CHI PHÍ",
                Font = UIHelper.GetFont(12f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(16, 2, 0, 0)
            };

            _lblTenInfo = new Label
            {
                Dock = DockStyle.Fill,
                Font = UIHelper.GetFont(11f, FontStyle.Bold),
                ForeColor = Color.FromArgb(47, 54, 64),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(16, 0, 0, 8)
            };

            pnlHeader.Controls.Add(_lblTenInfo);
            pnlHeader.Controls.Add(lblTieuDe);

            // ===== BOTTOM: CÁC NÚT THAO TÁC =====
            var pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 64,
                Padding = new Padding(14, 10, 14, 10),
                BackColor = Color.White
            };
            pnlBottom.Paint += (s, e) =>
            {
                using var pen = new Pen(Color.FromArgb(220, 221, 225));
                e.Graphics.DrawLine(pen, 0, 0, pnlBottom.Width, 0);
            };

            var ttToolTip = new ToolTip();
            _btnMoPdf = new Button
            {
                Text = "Mở PDF",
                Size = new Size(160, 42),
                BackColor = Color.FromArgb(0, 102, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            _btnMoPdf.Click += BtnMoPdf_Click;
            ttToolTip.SetToolTip(_btnMoPdf, "Mở trực tiếp đúng trang của Phụ lục VIII - TT 38/2026\r\n(Giữ phím Ctrl + bấm để chọn lại file PDF)");

            _btnApDung = new Button
            {
                Text = "Áp dụng tỷ lệ nội suy",
                Size = new Size(250, 42),
                BackColor = Color.FromArgb(33, 115, 70),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            _btnApDung.Click += BtnApDung_Click;

            var btnDong = new Button
            {
                Text = "Đóng",
                Size = new Size(110, 42),
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f),
                Cursor = Cursors.Hand
            };
            btnDong.Click += (s, e) => this.Close();
            _btnDong = btnDong;

            _pnlBottom = pnlBottom;
            pnlBottom.Controls.Add(btnDong);
            pnlBottom.Controls.Add(_btnApDung);
            pnlBottom.Controls.Add(_btnMoPdf);
            pnlBottom.Resize += (s, e) => ReLayoutBottom();
            ReLayoutBottom();

            // ===== CENTER =====
            var pnlCenter = new Panel { Dock = DockStyle.Fill, Padding = new Padding(14, 10, 14, 10), AutoScroll = true };

            _lblGhiChu = new Label
            {
                Dock = DockStyle.Top,
                AutoSize = false,
                Height = 30,
                Font = UIHelper.GetFont(9.5f, FontStyle.Italic),
                ForeColor = Color.FromArgb(74, 85, 104),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(2, 4, 0, 0),
                Visible = false
            };

            // ---- KẾT QUẢ NỘI SUY (định mức đang áp dụng) ----
            var grpKetQua = new GroupBox
            {
                Dock = DockStyle.Top,
                Height = 96,
                Text = "  KẾT QUẢ NỘI SUY THEO QUY MÔ HIỆN TẠI  ",
                Font = UIHelper.GetFont(10f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                Padding = new Padding(12, 6, 12, 6)
            };

            _lblTyLeNoiSuy = new Label
            {
                AutoSize = true,
                Font = UIHelper.GetFont(13f, FontStyle.Bold),
                ForeColor = Color.FromArgb(33, 115, 70),
                Location = new Point(16, 24)
            };
            _lblTyLeHienTai = new Label
            {
                AutoSize = true,
                Font = UIHelper.GetFont(11f),
                ForeColor = Color.FromArgb(120, 80, 0),
                Location = new Point(16, 56)
            };
            grpKetQua.Controls.Add(_lblTyLeNoiSuy);
            grpKetQua.Controls.Add(_lblTyLeHienTai);

            // ---- CÔNG CỤ TÍNH NỘI SUY TAY ----
            var grpNoiSuy = new GroupBox
            {
                Dock = DockStyle.Top,
                Height = 214,
                Text = "  CÔNG CỤ TÍNH NỘI SUY TAY (THEO ĐÚNG CÔNG THỨC THÔNG TƯ)  ",
                Font = UIHelper.GetFont(10f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                Padding = new Padding(8, 2, 8, 4)
            };

            var tblNoiSuy = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 5,
                Padding = new Padding(0)
            };
            tblNoiSuy.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
            tblNoiSuy.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14));
            tblNoiSuy.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            tblNoiSuy.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14));
            tblNoiSuy.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            tblNoiSuy.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            tblNoiSuy.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            tblNoiSuy.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            tblNoiSuy.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            _txtQ1 = MakeNumText(tblNoiSuy, "Mốc dưới  Q1 (tỷ đồng):", 0, 0);
            _txtN1 = MakeNumText(tblNoiSuy, "Tỷ lệ tại Q1  (%):", 0, 2);
            _txtQ2 = MakeNumText(tblNoiSuy, "Mốc trên  Q2 (tỷ đồng):", 1, 0);
            _txtN2 = MakeNumText(tblNoiSuy, "Tỷ lệ tại Q2  (%):", 1, 2);
            _txtQ = MakeNumText(tblNoiSuy, "Quy mô cần kiểm tra  Q (tỷ đồng):", 2, 0);
            tblNoiSuy.SetColumnSpan(_txtQ, 3);

            _txtQ1.TextChanged += (s, e) => TinhNoiSuyTay();
            _txtN1.TextChanged += (s, e) => TinhNoiSuyTay();
            _txtQ2.TextChanged += (s, e) => TinhNoiSuyTay();
            _txtN2.TextChanged += (s, e) => TinhNoiSuyTay();
            _txtQ.TextChanged += (s, e) => TinhNoiSuyTay();

            var btnNapMoc = new Button
            {
                Text = "Nạp mốc từ bảng định mức",
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                MinimumSize = new Size(40, 34),
                Margin = new Padding(0, 4, 10, 4),
                Padding = new Padding(18, 0, 18, 0),
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(9.5f),
                Cursor = Cursors.Hand
            };
            btnNapMoc.Click += (s, e) => { NapMocTuBang(); TinhNoiSuyTay(); };

            var btnApDungTay = new Button
            {
                Text = "Áp dụng kết quả này",
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                MinimumSize = new Size(40, 34),
                Margin = new Padding(10, 4, 0, 4),
                Padding = new Padding(18, 0, 18, 0),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(33, 115, 70),
                ForeColor = Color.White,
                Font = UIHelper.GetFont(9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnApDungTay.Click += (s, e) => ApDungTyLe(_tyLeTinhTay);

            var flowBtns = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0),
                Padding = new Padding(0),
                Anchor = AnchorStyles.None
            };
            flowBtns.Controls.Add(btnNapMoc);
            flowBtns.Controls.Add(btnApDungTay);

            var tblNoiSuyBtns = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                Height = 40,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            tblNoiSuyBtns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            tblNoiSuyBtns.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            tblNoiSuyBtns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            tblNoiSuyBtns.Controls.Add(flowBtns, 1, 0);

            tblNoiSuy.Controls.Add(tblNoiSuyBtns, 0, 3);
            tblNoiSuy.SetColumnSpan(tblNoiSuyBtns, 4);

            _lblNoiSuyResult = new Label
            {
                Dock = DockStyle.Fill,
                Font = UIHelper.GetFont(9.5f),
                ForeColor = Color.FromArgb(31, 41, 55),
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(8, 4, 8, 4),
                BackColor = Color.FromArgb(248, 250, 252)
            };
            tblNoiSuy.Controls.Add(_lblNoiSuyResult, 0, 4);
            tblNoiSuy.SetColumnSpan(_lblNoiSuyResult, 4);

            grpNoiSuy.Controls.Add(tblNoiSuy);

            // ---- BẢNG TRA ----
            var lblBangTra = new Label
            {
                Dock = DockStyle.Top,
                Height = 26,
                Text = "BẢNG TRA TỶ LỆ THEO MỐC QUY MÔ",
                Font = UIHelper.GetFont(10f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(2, 0, 0, 0)
            };

            _dgvBangTra = new DataGridView
            {
                Dock = DockStyle.Top,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                CellBorderStyle = DataGridViewCellBorderStyle.Single,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ScrollBars = ScrollBars.Vertical,
                RowTemplate = { Height = 28 },
                Font = UIHelper.GetFont(10f)
            };
            UIHelper.ApplyStyle(_dgvBangTra);
            _dgvBangTra.DataError += (s, e) => e.ThrowException = false;
            _dgvBangTra.Columns.Add("ColMoc", "Mốc quy mô (tỷ đồng)");
            _dgvBangTra.Columns.Add("ColTiLe", "Tỷ lệ định mức (%)");
            _dgvBangTra.Columns[0].FillWeight = 40;
            _dgvBangTra.Columns[1].FillWeight = 60;
            _dgvBangTra.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            _dgvBangTra.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            _dgvBangTra.Columns[1].DefaultCellStyle.Font = UIHelper.GetFont(10f, FontStyle.Bold);

            // ---- THÔNG TIN BẢNG ----
            var tblInfo = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 128,
                ColumnCount = 2,
                RowCount = 4,
                BackColor = Color.White,
                Padding = new Padding(14, 8, 14, 4)
            };
            tblInfo.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 265));
            tblInfo.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            tblInfo.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            tblInfo.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            tblInfo.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            tblInfo.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            tblInfo.Controls.Add(MakeLabel("Bảng định mức áp dụng:"), 0, 0);
            tblInfo.Controls.Add(MakeValueLabel(""), 1, 0);
            tblInfo.Controls.Add(MakeLabel("Cơ sở tính:"), 0, 1);
            tblInfo.Controls.Add(MakeValueLabel(""), 1, 1);
            tblInfo.Controls.Add(MakeLabel("Loại/Cấp công trình:"), 0, 2);
            tblInfo.Controls.Add(MakeValueLabel(""), 1, 2);
            tblInfo.Controls.Add(MakeLabel("Quy mô nội suy:"), 0, 3);
            tblInfo.Controls.Add(MakeValueLabel(""), 1, 3);

            // Thêm theo thứ tự dưới -> trên
            pnlCenter.Controls.Add(_lblGhiChu);
            pnlCenter.Controls.Add(grpKetQua);
            pnlCenter.Controls.Add(grpNoiSuy);
            pnlCenter.Controls.Add(_dgvBangTra);
            pnlCenter.Controls.Add(lblBangTra);
            pnlCenter.Controls.Add(tblInfo);

            this.Controls.Add(pnlCenter);
            this.Controls.Add(pnlBottom);
            this.Controls.Add(pnlHeader);

            this.KeyPreview = true;
            this.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape) this.Close();
            };
        }

        private static TextBox MakeNumText(TableLayoutPanel parent, string label, int row, int col)
        {
            var lbl = new Label
            {
                Text = label,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = UIHelper.GetFont(9.5f)
            };
            parent.Controls.Add(lbl, col, row);
            var txt = new TextBox
            {
                Dock = DockStyle.Fill,
                Font = UIHelper.GetFont(10f),
                TextAlign = HorizontalAlignment.Right,
                Margin = new Padding(4, 3, 4, 3)
            };
            parent.Controls.Add(txt, col + 1, row);
            return txt;
        }

        private static Label MakeLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                Font = UIHelper.GetFont(10f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        private static Label MakeValueLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                Font = UIHelper.GetFont(10.5f),
                ForeColor = Color.FromArgb(47, 54, 64),
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };
        }

        private static string Fmt(decimal value, string format = "0.####")
            => value.ToString(format, UIHelper.ViCulture);

        private void NapDuLieu()
        {
            _lblTenInfo.Text = $"{_item.TenChiPhi}   [{_item.MaChiPhi}]";

            var pnlCenter = this.Controls.OfType<Panel>().FirstOrDefault(p => p.Dock == DockStyle.Fill);
            var tblInfo = pnlCenter?.Controls.OfType<TableLayoutPanel>().FirstOrDefault();
            if (tblInfo != null)
            {
                ((Label)tblInfo.GetControlFromPosition(1, 0)).Text = _tt.TenBang;
                ((Label)tblInfo.GetControlFromPosition(1, 1)).Text = _tt.CoSoTinh;
                ((Label)tblInfo.GetControlFromPosition(1, 2)).Text = $"{_tt.LoaiCongTrinh}  -  Cấp {_tt.CapCongTrinh}";
                ((Label)tblInfo.GetControlFromPosition(1, 3)).Text = $"{Fmt(_tt.QuyMoTy, "0.###")} tỷ đồng";
            }

            if (_tt.SoTrangPdf > 0)
            {
                _btnMoPdf.Text = $"Mở PDF - trang {_tt.SoTrangPdf}";
                _btnMoPdf.Enabled = true;
            }
            else
            {
                _btnMoPdf.Text = "Mở PDF (BTC)";
                _btnMoPdf.Enabled = false;
            }
            FitButton(_btnMoPdf);

            _dgvBangTra.Rows.Clear();
            int n = Math.Min(_tt.MocQuyMo?.Length ?? 0, _tt.TiLeDinhMuc?.Length ?? 0);
            int highlight = -1;
            for (int i = 0; i < n; i++)
            {
                int idx = _dgvBangTra.Rows.Add(
                    Fmt(_tt.MocQuyMo[i], "0.###"),
                    Fmt(_tt.TiLeDinhMuc[i]));
                if (_tt.QuyMoTy >= _tt.MocQuyMo[i]) highlight = idx;
            }
            if (n == 0)
            {
                _dgvBangTra.Rows.Add("-", "-");
            }

            foreach (DataGridViewRow r in _dgvBangTra.Rows)
            {
                if (r.Index == highlight)
                {
                    r.DefaultCellStyle.BackColor = Color.FromArgb(220, 235, 252);
                    r.DefaultCellStyle.Font = UIHelper.GetFont(10f, FontStyle.Bold);
                }
            }

            // Chiều cao bảng tra = vừa đủ hiển thị hết các mốc, không cần cuộn dọc
            int gridH = _dgvBangTra.ColumnHeadersHeight
                        + (_dgvBangTra.Rows.Count * _dgvBangTra.RowTemplate.Height)
                        + 2;
            _dgvBangTra.Height = Math.Min(480, gridH);
            _dgvBangTra.ScrollBars = (_dgvBangTra.Rows.Count * _dgvBangTra.RowTemplate.Height) > 480
                ? ScrollBars.Vertical
                : ScrollBars.None;

            _lblTyLeNoiSuy.Text = $"Tỷ lệ nội suy: {Fmt(_tt.TyLeNoiSuy)} %    x    Hệ số k = {Fmt(_tt.HeSo, "0.##")}";
            _lblTyLeHienTai.Text = $"Tỷ lệ đang áp dụng trên dòng: {Fmt(_tt.TyLeHienTai)} %";

            if (!string.IsNullOrEmpty(_tt.GhiChu))
            {
                _lblGhiChu.Text = "   " + _tt.GhiChu;
                _lblGhiChu.Visible = true;
            }

            if (_item.MaChiPhi == "K_TD_DA")
            {
                _btnApDung.Text = "Tỷ lệ tự động theo TMĐT";
                _btnApDung.Enabled = false;
                if (string.IsNullOrEmpty(_tt.GhiChu))
                {
                    _lblGhiChu.Text = "   K_TD_DA được tự động nội suy lại theo Tổng mức đầu tư mỗi khi dự toán thay đổi; không cần áp dụng thủ công.";
                    _lblGhiChu.Visible = true;
                }
            }
            else
            {
                _btnApDung.Enabled = true;
            }

            // Công cụ nội suy tay: nạp sẵn mốc từ bảng định mức
            NapMocTuBang();
            ReLayoutBottom();
        }

        private static void FitButton(Button btn)
        {
            btn.AutoSize = false;
            var size = TextRenderer.MeasureText(btn.Text, btn.Font);
            btn.Width = Math.Max(150, size.Width + 36);
        }

        /// <summary>Căn giữa 3 nút thao tác ở đáy modal, có khoảng cách đều nhau</summary>
        private void ReLayoutBottom()
        {
            try
            {
                if (_btnMoPdf == null || _btnApDung == null || _btnDong == null || _pnlBottom == null) return;
                const int gap = 14;
                int total = _btnMoPdf.Width + gap + _btnApDung.Width + gap + _btnDong.Width;
                int x = Math.Max(14, (_pnlBottom.Width - total) / 2);
                int y = (Math.Max(_pnlBottom.Height, 42) - 42) / 2;
                _btnMoPdf.Location = new Point(x, y);
                _btnApDung.Location = new Point(x + _btnMoPdf.Width + gap, y);
                _btnDong.Location = new Point(_btnApDung.Left + _btnApDung.Width + gap, y);
            }
            catch { }
        }

        private void NapMocTuBang()
        {
            _txtQ.Text = Fmt(_tt.QuyMoTy, "0.###");

            var mocs = _tt.MocQuyMo;
            var tils = _tt.TiLeDinhMuc;
            if (mocs == null || tils == null || mocs.Length != tils.Length || mocs.Length == 0) return;

            decimal q = _tt.QuyMoTy;
            _txtQ1.Text = Fmt(mocs[0], "0.###");
            _txtN1.Text = Fmt(tils[0]);

            if (q <= mocs[0])
            {
                _txtQ1.Text = Fmt(mocs[0], "0.###");
                _txtN1.Text = Fmt(tils[0]);
                _txtQ2.Text = Fmt(mocs[0], "0.###");
                _txtN2.Text = Fmt(tils[0]);
                return;
            }
            if (q >= mocs[mocs.Length - 1])
            {
                _txtQ1.Text = Fmt(mocs[mocs.Length - 1], "0.###");
                _txtN1.Text = Fmt(tils[tils.Length - 1]);
                _txtQ2.Text = Fmt(mocs[mocs.Length - 1], "0.###");
                _txtN2.Text = Fmt(tils[tils.Length - 1]);
                return;
            }

            for (int i = 0; i < mocs.Length - 1; i++)
            {
                if (q >= mocs[i] && q <= mocs[i + 1])
                {
                    _txtQ1.Text = Fmt(mocs[i], "0.###");
                    _txtN1.Text = Fmt(tils[i]);
                    _txtQ2.Text = Fmt(mocs[i + 1], "0.###");
                    _txtN2.Text = Fmt(tils[i + 1]);
                    return;
                }
            }
        }

        private void TinhNoiSuyTay()
        {
            if (_lblNoiSuyResult == null || _isInit) return;

            try
            {
            decimal q1 = UIHelper.ParseFlexibleDecimal(_txtQ1.Text);
            decimal n1 = UIHelper.ParseFlexibleDecimal(_txtN1.Text);
            decimal q2 = UIHelper.ParseFlexibleDecimal(_txtQ2.Text);
            decimal n2 = UIHelper.ParseFlexibleDecimal(_txtN2.Text);
            decimal q = UIHelper.ParseFlexibleDecimal(_txtQ.Text);

            if (q2 < q1)
            {
                (q1, q2) = (q2, q1);
                (n1, n2) = (n2, n1);
            }

            if (q2 == q1)
            {
                _lblNoiSuyResult.Text = "Q2 phải khác Q1 để thực hiện nội suy. Hãy nhập mốc dưới (Q1) và mốc trên (Q2) khác nhau.";
                return;
            }

            decimal nt;
            string note;
            if (q <= q1)
            {
                nt = n1;
                note = "Quy mô Q ≤ mốc dưới Q1 → lấy tỷ lệ tại mốc dưới (không nội suy).";
            }
            else if (q >= q2)
            {
                nt = n2;
                note = "Quy mô Q ≥ mốc trên Q2 → lấy tỷ lệ tại mốc trên (không nội suy).";
            }
            else
            {
                nt = n1 - (n1 - n2) * (q - q1) / (q2 - q1);
                note = "Nội suy tuyến tính giữa hai mốc.";
            }

            _tyLeTinhTay = Math.Round(nt, 4);

            _lblNoiSuyResult.Text =
                $"Công thức TT 38/2026:  N = N1 - ((N1 - N2) / (Q2 - Q1)) × (Q - Q1)\r\n" +
                $"  = {Fmt(n1)} - (({Fmt(n1)} - {Fmt(n2)}) / ({Fmt(q2)} - {Fmt(q1)})) × ({Fmt(q)} - {Fmt(q1)})\r\n" +
                $"  = {Fmt(_tyLeTinhTay)} %    ({note})";
            }
            catch { }
        }

        private void ApDungTyLe(decimal tyLe)
        {
            if (_item.MaChiPhi == "K_TD_DA")
            {
                MessageBox.Show("K_TD_DA được tính tự động theo Tổng mức đầu tư, không áp dụng thủ công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (tyLe <= 0)
            {
                MessageBox.Show("Chưa có kết quả tỷ lệ để áp dụng (kết quả phải > 0).", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _item.TyLePhanTram = tyLe;
            if (_tt.HeSo > 0) _item.HeSoDieuChinh = _tt.HeSo;

            _model.TinhToanLai();
            DinhMucTT38Engine.CapNhatTiLeThamDinhDuAn(_model, null, _model.YeuCauThueThamTra);

            _refreshGrid?.Invoke();

            _lblTyLeHienTai.Text = $"Tỷ lệ đang áp dụng trên dòng: {Fmt(_item.TyLePhanTram)} %";
            MessageBox.Show(
                $"Đã áp dụng tỷ lệ {Fmt(_item.TyLePhanTram)} % với hệ số k = {Fmt(_item.HeSoDieuChinh, "0.##")} cho khoản mục [{_item.TenChiPhi}].",
                "Áp dụng định mức", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnApDung_Click(object sender, EventArgs e)
        {
            // 1.3: tra cứu lại theo model hiện tại để không áp dụng tỷ lệ nội suy đã cũ
            try
            {
                var ttMoi = DinhMucTT38Engine.TraCuuThongTinChiPhiItem(_model, _item.MaChiPhi);
                if (ttMoi != null && ttMoi.TyLeNoiSuy > 0)
                {
                    ApDungTyLe(ttMoi.TyLeNoiSuy);
                    return;
                }
            }
            catch { }
            ApDungTyLe(_tt.TyLeNoiSuy);
        }

        // ==================== MỞ PDF / LƯU TRỮ PDF TRONG PHẦN MỀM ====================

        private void BtnMoPdf_Click(object sender, EventArgs e)
        {
            if (_tt.SoTrangPdf <= 0)
            {
                MessageBox.Show("Bảng định mức này nằm trong văn bản Bộ Tài chính, không thuộc Phụ lục VIII của TT 38/2026.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            bool chonLai = (ModifierKeys & Keys.Control) == Keys.Control;

            if (_pdfPath == null || !File.Exists(_pdfPath))
            {
                _pdfPath = File.Exists(PdfStoredPath) ? PdfStoredPath : TimFilePdfDefault();
            }

            if (chonLai || string.IsNullOrEmpty(_pdfPath))
            {
                using var ofd = new OpenFileDialog();
                ofd.Filter = "PDF | *.pdf";
                ofd.Title = chonLai
                    ? "Chọn lại file 38.2026.TT-BXD-PL8.pdf (Phụ lục VIII - TT 38/2026)"
                    : "Chọn file 38.2026.TT-BXD-PL8.pdf (Phụ lục VIII - TT 38/2026)";
                if (ofd.ShowDialog() != DialogResult.OK) return;
                _pdfPath = LuuPdf(ofd.FileName);
                ThongBaoDaLuuPdf();
            }
            else if (_pdfPath != PdfStoredPath && File.Exists(_pdfPath))
            {
                _pdfPath = LuuPdf(_pdfPath);
                ThongBaoDaLuuPdf();
            }

            if (string.IsNullOrEmpty(_pdfPath) || !File.Exists(_pdfPath)) return;

            try
            {
                MoPdfDungTrang(_pdfPath, _tt.SoTrangPdf);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể mở PDF: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void ThongBaoDaLuuPdf()
        {
            try
            {
                if (Directory.Exists(PdfStorageDir))
                {
                    MessageBox.Show(
                        "Đã lưu file PDF vào phần mềm để sử dụng cho các lần kiểm tra sau.\n\n" +
                        "Vị trí lưu: " + PdfStoredPath + "\n\n" +
                        "(Muốn đổi file khác: giữ phím Ctrl + bấm nút Mở PDF)",
                        "Kiểm tra định mức", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch { }
        }

        private static string LuuPdf(string source)
        {
            try
            {
                if (string.IsNullOrEmpty(source) || !File.Exists(source)) return null;
                if (!Directory.Exists(PdfStorageDir)) Directory.CreateDirectory(PdfStorageDir);
                File.Copy(source, PdfStoredPath, true);
                return PdfStoredPath;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể lưu file PDF vào phần mềm: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return source;
            }
        }

        private static string TimFilePdfDefault()
        {
            string[] dirs =
            {
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"),
                Environment.GetFolderPath(Environment.SpecialFolder.MyComputer)
            };
            string[] dirsAi =
            {
                PdfStorageDir,
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AIE_DuToan"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AIE")
            };

            string[] names = { PdfFileName, "38.2026.TT-BXD.PL8.pdf", "TT38-2026-PL8.pdf", "PhuLuc8-TT38-2026.pdf" };

            foreach (string dir in dirsAi.Where(d => !string.IsNullOrEmpty(d) && Directory.Exists(d)))
            {
                foreach (string name in names)
                {
                    string full = Path.Combine(dir, name);
                    if (File.Exists(full)) return full;
                }
            }
            foreach (string dir in dirs.Where(d => !string.IsNullOrEmpty(d) && Directory.Exists(d)))
            {
                foreach (string name in names)
                {
                    string full = Path.Combine(dir, name);
                    if (File.Exists(full)) return full;
                }
            }
            return null;
        }

        /// <summary>
        /// Mở PDF và nhảy thẳng đến trang chứa bảng định mức.
        /// Ưu tiên Adobe Reader/Acrobat (hỗ trợ /A "page=N"), fallback sang trình duyệt/reader mặc định qua file://...#page=N
        /// </summary>
        private static void MoPdfDungTrang(string pdfPath, int page)
        {
            var adobe = TimAdobeReader();
            if (adobe != null && File.Exists(adobe))
            {
                Process.Start(new ProcessStartInfo(adobe)
                {
                    UseShellExecute = false,
                    Arguments = $"/A \"page={page}\" \"{pdfPath}\""
                });
                return;
            }

            string uri = "file:///" + pdfPath.Replace('\\', '/') + "#page=" + page;
            Process.Start(new ProcessStartInfo(uri) { UseShellExecute = true });
        }

        private static string TimAdobeReader()
        {
            string[] keys =
            {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Acrobat.exe",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\Acrobat.exe"
            };
            foreach (var k in keys)
            {
                try
                {
                    using var key = Registry.LocalMachine.OpenSubKey(k) ?? Registry.CurrentUser.OpenSubKey(k);
                    if (key?.GetValue(null) is string s && !string.IsNullOrEmpty(s) && File.Exists(s))
                        return s;
                }
                catch { }
            }
            return null;
        }
    }
}