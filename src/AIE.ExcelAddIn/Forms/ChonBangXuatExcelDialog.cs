using System;
using System.Drawing;
using System.Windows.Forms;
using AIE.ExcelAddIn.Services;
using AIE.ExcelAddIn.Helpers;

namespace AIE.ExcelAddIn.Forms
{
    public class ChonBangXuatExcelDialog : Form
    {
        private CheckBox chkTMDT;
        private CheckBox chkTHDT;
        private CheckBox chkTHCPXD;

        private CheckBox chkDuToan;
        private CheckBox chkPhanTich;
        private CheckBox chkTHVL;
        private CheckBox chkTHNC;
        private CheckBox chkTHMay;
        private CheckBox chkCuocVC;
        private CheckBox chkHeSo;

        private Button btnChonTatCa;
        private Button btnBoChonTatCa;
        private Button btnMacDinh;

        private Button btnXuat;
        private Button btnDong;

        private readonly bool _isCheDoTMDT;

        public LuaChonXuatExcel LuaChon { get; private set; }

        public ChonBangXuatExcelDialog(bool isCheDoTMDT = true)
        {
            _isCheDoTMDT = isCheDoTMDT;
            LuaChon = new LuaChonXuatExcel();

            InitializeComponent();
            FormStateHelper.Attach(this);
            ApDungMacDinh();
        }

        private void InitializeComponent()
        {
            this.Text = "Tùy chọn Bảng biểu Xuất sang Excel (AIE Dự Toán)";
            
            // Tính toán kích thước tự động co giãn rộng rãi theo độ phân giải màn hình
            var area = Screen.PrimaryScreen.WorkingArea;
            int targetW = Math.Min(1180, Math.Max(1080, area.Width - 40));
            int targetH = Math.Min(690, Math.Max(580, area.Height - 60));
            this.ClientSize = new Size(targetW, targetH);
            this.MinimumSize = new Size(980, 540);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = true;
            this.ShowInTaskbar = true;
            this.BackColor = Color.FromArgb(248, 250, 253);
            this.Font = UIHelper.GetFont(9.5f);
            this.AutoScaleMode = AutoScaleMode.Font;

            // =========================================================================
            // 1. HEADER BANNER
            // =========================================================================
            var pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 62,
                BackColor = Color.FromArgb(16, 124, 65), // Excel Green
                Padding = new Padding(20, 10, 20, 10)
            };

            var lblHeaderTitle = new Label
            {
                Text = "📥 TÙY CHỌN BẢNG BIỂU XUẤT SANG EXCEL",
                Dock = DockStyle.Top,
                Height = 24,
                Font = UIHelper.GetFont(12f, FontStyle.Bold),
                ForeColor = Color.White
            };

            var lblHeaderSub = new Label
            {
                Text = "Tích chọn các bảng biểu dự toán & tổng hợp kinh phí bạn muốn tạo/cập nhật vào Workbook:",
                Dock = DockStyle.Bottom,
                Height = 20,
                Font = UIHelper.GetFont(9.5f, FontStyle.Regular),
                ForeColor = Color.FromArgb(230, 244, 234)
            };

            pnlHeader.Controls.Add(lblHeaderTitle);
            pnlHeader.Controls.Add(lblHeaderSub);
            this.Controls.Add(pnlHeader);

            // =========================================================================
            // 2. TOOLBAR NÚT CHỌN NHANH
            // =========================================================================
            var pnlToolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 44,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(16, 7, 16, 7),
                BackColor = Color.FromArgb(240, 244, 248)
            };

            btnChonTatCa = new Button
            {
                Text = "📁 Trọn bộ hồ sơ",
                AutoSize = true,
                Height = 30,
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(9.5f),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 8, 0)
            };
            btnChonTatCa.FlatAppearance.BorderColor = Color.FromArgb(209, 213, 219);
            btnChonTatCa.Click += (s, e) => SetAllCheckboxes(true);

            btnBoChonTatCa = new Button
            {
                Text = "☐ Bỏ chọn tất cả",
                AutoSize = true,
                Height = 30,
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(9.5f),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 8, 0)
            };
            btnBoChonTatCa.FlatAppearance.BorderColor = Color.FromArgb(209, 213, 219);
            btnBoChonTatCa.Click += (s, e) => SetAllCheckboxes(false);

            btnMacDinh = new Button
            {
                Text = "↺ Mặc định đề xuất",
                AutoSize = true,
                Height = 30,
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 8, 0)
            };
            btnMacDinh.FlatAppearance.BorderColor = Color.FromArgb(0, 102, 204);
            btnMacDinh.Click += (s, e) => ApDungMacDinh();

            pnlToolbar.Controls.Add(btnChonTatCa);
            pnlToolbar.Controls.Add(btnBoChonTatCa);
            pnlToolbar.Controls.Add(btnMacDinh);
            this.Controls.Add(pnlToolbar);

            // =========================================================================
            // 3. BOTTOM PANEL: NÚT XUẤT & HỦY (CỐ ĐỊNH CHÂN MODAL, TABLELAYOUT KHÔNG BAO GIỜ WRAP)
            // =========================================================================
            var pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                BackColor = Color.White,
                Padding = new Padding(12, 8, 16, 8)
            };
            pnlBottom.Paint += (s, e) =>
            {
                using var pen = new Pen(Color.FromArgb(226, 232, 240));
                e.Graphics.DrawLine(pen, 0, 0, pnlBottom.Width, 0);
            };

            var tblBottom = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 1,
                ColumnCount = 2,
                BackColor = Color.Transparent,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            tblBottom.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            tblBottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); // Cột trái: Tip
            tblBottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // Cột phải: Buttons

            var lblBottomTip = new Label
            {
                Text = "💡 Mẹo: Bảng kỹ thuật đã được tạo khi 'Áp giá'. Tại đây chỉ cần tích chọn nếu muốn xuất lại toàn bộ.",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = UIHelper.GetFont(9f, FontStyle.Italic),
                ForeColor = Color.FromArgb(100, 116, 139)
            };

            var pnlBottomButtons = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                AutoSize = true,
                WrapContents = false, // KHÔNG BAO GIỜ WRAP DÒNG
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(0)
            };

            btnDong = new Button
            {
                Text = "Đóng",
                Size = new Size(100, 40),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f),
                Cursor = Cursors.Hand,
                UseCompatibleTextRendering = true,
                Margin = new Padding(0, 0, 10, 0)
            };
            btnDong.FlatAppearance.BorderColor = Color.FromArgb(209, 213, 219);
            btnDong.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            btnXuat = new Button
            {
                Text = "📥 Bắt đầu xuất Excel",
                Size = new Size(220, 40),
                BackColor = Color.FromArgb(16, 124, 65), // Excel Green
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                UseCompatibleTextRendering = true,
                Margin = new Padding(0)
            };
            btnXuat.FlatAppearance.BorderSize = 0;
            btnXuat.Click += BtnXuat_Click;

            pnlBottomButtons.Controls.Add(btnDong);
            pnlBottomButtons.Controls.Add(btnXuat);

            tblBottom.Controls.Add(lblBottomTip, 0, 0);
            tblBottom.Controls.Add(pnlBottomButtons, 1, 0);
            pnlBottom.Controls.Add(tblBottom);
            this.Controls.Add(pnlBottom);

            // =========================================================================
            // 4. MAIN CONTENT PANEL: BỐ CỤC 2 CỘT TỰ CO GIÃN (RESPONSIVE TABLELAYOUT)
            // =========================================================================
            var pnlContent = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(16, 8, 16, 8),
                BackColor = Color.FromArgb(248, 250, 253)
            };

            var tblContent = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 1,
                ColumnCount = 2,
                Margin = new Padding(0),
                Padding = new Padding(0),
                BackColor = Color.Transparent
            };
            tblContent.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            tblContent.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f)); // Cột trái: 50%
            tblContent.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f)); // Cột phải: 50%

            // CỘT 1: BẢNG TỔNG HỢP KINH PHÍ (TT 36/2026/TT-BXD)
            var grpTongHop = new GroupBox
            {
                Text = "  1. BẢNG TỔNG HỢP KINH PHÍ (TT 36/2026)  ",
                Dock = DockStyle.Fill,
                Font = UIHelper.GetFont(10f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                Padding = new Padding(12, 8, 12, 8),
                Margin = new Padding(0, 0, 6, 0)
            };

            var pnlGrp1 = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.Transparent
            };

            var tblGrp1 = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(4),
                BackColor = Color.Transparent
            };
            tblGrp1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            chkTMDT = CreateCheckbox("Bảng 1.2: Tổng mức đầu tư xây dựng", "TongMucDauTu", Color.FromArgb(153, 51, 0));
            chkTHDT = CreateCheckbox("Bảng 2.1: Tổng hợp dự toán công trình", "TH_DuToan", Color.FromArgb(0, 102, 204));
            chkTHCPXD = CreateCheckbox("Bảng 3.8: Bảng tổng hợp chi phí xây dựng", "TH_ChiPhiXD", Color.FromArgb(0, 102, 0));

            var pnlNoteTMDT = new Panel
            {
                Dock = DockStyle.Top,
                Height = 175,
                BackColor = Color.FromArgb(240, 248, 255),
                Padding = new Padding(12, 10, 12, 10),
                Margin = new Padding(2, 14, 2, 4)
            };
            pnlNoteTMDT.Paint += (s, e) =>
            {
                using var pen = new Pen(Color.FromArgb(186, 230, 253), 1);
                e.Graphics.DrawRectangle(pen, 0, 0, pnlNoteTMDT.Width - 1, pnlNoteTMDT.Height - 1);
            };

            var lblNoteTitle = new Label
            {
                Text = "📌 Quy chuẩn Thông tư 36/2026/TT-BXD:",
                Dock = DockStyle.Top,
                Height = 24,
                Font = UIHelper.GetFont(9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(3, 105, 161),
                UseCompatibleTextRendering = true
            };
            var lblNoteContent = new Label
            {
                Text = "• Bảng 1.2: Dành cho Báo cáo KT-KT hoặc Tổng mức đầu tư.\n" +
                       "• Bảng 2.1: Dành cho giai đoạn Lập dự toán xây dựng công trình.\n" +
                       "• Bảng 3.8: Bảng tổng hợp chi phí xây dựng làm cơ sở tính toán.",
                Dock = DockStyle.Fill,
                Font = UIHelper.GetFont(9f, FontStyle.Regular),
                ForeColor = Color.FromArgb(12, 74, 110),
                TextAlign = ContentAlignment.TopLeft,
                UseCompatibleTextRendering = true
            };
            pnlNoteTMDT.Controls.Add(lblNoteContent);
            pnlNoteTMDT.Controls.Add(lblNoteTitle);

            Action adjustNoteHeight = () =>
            {
                try
                {
                    int w = pnlNoteTMDT.ClientSize.Width - pnlNoteTMDT.Padding.Horizontal;
                    if (w <= 50) w = 400;
                    using var g = pnlNoteTMDT.CreateGraphics();
                    var szTitle = TextRenderer.MeasureText(g, lblNoteTitle.Text, lblNoteTitle.Font, new Size(w, int.MaxValue), TextFormatFlags.WordBreak);
                    var szContent = TextRenderer.MeasureText(g, lblNoteContent.Text, lblNoteContent.Font, new Size(w, int.MaxValue), TextFormatFlags.WordBreak);
                    int h = szTitle.Height + szContent.Height + pnlNoteTMDT.Padding.Vertical + 16;
                    pnlNoteTMDT.Height = Math.Max(170, h);
                }
                catch { }
            };
            pnlNoteTMDT.SizeChanged += (s, e) => adjustNoteHeight();

            tblGrp1.Controls.Add(chkTMDT, 0, 0);
            tblGrp1.Controls.Add(chkTHDT, 0, 1);
            tblGrp1.Controls.Add(chkTHCPXD, 0, 2);
            tblGrp1.Controls.Add(pnlNoteTMDT, 0, 3);

            pnlGrp1.Controls.Add(tblGrp1);
            grpTongHop.Controls.Add(pnlGrp1);

            // CỘT 2: BẢNG BIỂU KỸ THUẬT & DỰ TOÁN CHI TIẾT
            var grpKyThuat = new GroupBox
            {
                Text = "  2. HỆ THỐNG BẢNG BIỂU KỸ THUẬT DỰ TOÁN CHI TIẾT  ",
                Dock = DockStyle.Fill,
                Font = UIHelper.GetFont(10f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                Padding = new Padding(12, 8, 12, 8),
                Margin = new Padding(6, 0, 0, 0)
            };

            var pnlGrp2 = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.Transparent
            };

            var tblGrp2 = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                RowCount = 7,
                Padding = new Padding(4),
                BackColor = Color.Transparent
            };
            tblGrp2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            chkDuToan = CreateCheckbox("Dự toán chi phí xây dựng công trình", "DuToan", Color.Black);
            chkPhanTich = CreateCheckbox("Bảng phân tích đơn giá chi tiết", "PhanTich_DonGia", Color.Black);
            chkTHVL = CreateCheckbox("Bảng tổng hợp chênh lệch giá vật liệu", "TH_VatLieu", Color.Black);
            chkTHNC = CreateCheckbox("Bảng tổng hợp và chênh lệch nhân công", "TH_NhanCong", Color.Black);
            chkTHMay = CreateCheckbox("Bảng tổng hợp và chênh lệch máy thi công", "TH_CaMay", Color.Black);
            chkCuocVC = CreateCheckbox("Bảng chiết tính cước vận chuyển", "ChietTinh_CuocVC", Color.Black);
            chkHeSo = CreateCheckbox("Bảng xác định hệ số điều chỉnh", "HeSo_DieuChinh", Color.Black);

            tblGrp2.Controls.Add(chkDuToan, 0, 0);
            tblGrp2.Controls.Add(chkPhanTich, 0, 1);
            tblGrp2.Controls.Add(chkTHVL, 0, 2);
            tblGrp2.Controls.Add(chkTHNC, 0, 3);
            tblGrp2.Controls.Add(chkTHMay, 0, 4);
            tblGrp2.Controls.Add(chkCuocVC, 0, 5);
            tblGrp2.Controls.Add(chkHeSo, 0, 6);

            // Ràng buộc phụ thuộc công thức: TH_ChiPhiXD tham chiếu trực tiếp đến HeSo_DieuChinh
            chkTHCPXD.CheckedChanged += (s, e) =>
            {
                if (chkTHCPXD.Checked && !chkHeSo.Checked)
                {
                    chkHeSo.Checked = true;
                }
            };

            pnlGrp2.Controls.Add(tblGrp2);
            grpKyThuat.Controls.Add(pnlGrp2);

            tblContent.Controls.Add(grpTongHop, 0, 0);
            tblContent.Controls.Add(grpKyThuat, 1, 0);
            pnlContent.Controls.Add(tblContent);
            this.Controls.Add(pnlContent);

            // Thứ tự hiển thị Z-Order
            pnlHeader.BringToFront();
            pnlToolbar.BringToFront();
            pnlBottom.SendToBack();
            pnlContent.BringToFront();

            this.Shown += (s, e) =>
            {
                adjustNoteHeight();
                foreach (Control c in tblGrp1.Controls)
                    if (c is CheckBox cb) AdjustCheckboxHeight(cb);
                foreach (Control c in tblGrp2.Controls)
                    if (c is CheckBox cb) AdjustCheckboxHeight(cb);
            };
        }

        private CheckBox CreateCheckbox(string title, string sheetName, Color tagColor)
        {
            var chk = new CheckBox
            {
                Text = $"{title}  [{sheetName}]",
                AutoSize = false,
                Dock = DockStyle.Top,
                Height = 30,
                Font = UIHelper.GetFont(9.5f, FontStyle.Regular),
                ForeColor = tagColor,
                Margin = new Padding(2, 4, 2, 4),
                Padding = new Padding(2, 2, 2, 2),
                Cursor = Cursors.Hand,
                CheckAlign = ContentAlignment.TopLeft,
                TextAlign = ContentAlignment.TopLeft,
                UseCompatibleTextRendering = true
            };

            chk.SizeChanged += (s, e) => AdjustCheckboxHeight(chk);
            return chk;
        }

        private void AdjustCheckboxHeight(CheckBox chk)
        {
            try
            {
                int textWidth = chk.ClientSize.Width - 28;
                if (textWidth <= 50) textWidth = 400;
                using var g = chk.CreateGraphics();
                var size = TextRenderer.MeasureText(g, chk.Text, chk.Font, new Size(textWidth, int.MaxValue), TextFormatFlags.WordBreak);
                chk.Height = Math.Max(28, size.Height + 8);
            }
            catch { }
        }

        private void SetAllCheckboxes(bool isChecked)
        {
            chkTMDT.Checked = isChecked;
            chkTHDT.Checked = isChecked;
            chkTHCPXD.Checked = isChecked;
            chkDuToan.Checked = isChecked;
            chkPhanTich.Checked = isChecked;
            chkTHVL.Checked = isChecked;
            chkTHNC.Checked = isChecked;
            chkTHMay.Checked = isChecked;
            chkCuocVC.Checked = isChecked;
            chkHeSo.Checked = isChecked;
        }

        private void ApDungMacDinh()
        {
            // Mặc định chọn bảng tổng hợp theo chế độ đang mở
            if (_isCheDoTMDT)
            {
                chkTMDT.Checked = true;
                chkTHDT.Checked = false;
            }
            else
            {
                chkTMDT.Checked = false;
                chkTHDT.Checked = true;
            }

            // Bảng tổng hợp chi phí xây dựng luôn mặc định chọn
            chkTHCPXD.Checked = true;

            // Bảng xác định hệ số điều chỉnh luôn mặc định chọn (hệ số ảnh hưởng trực tiếp tới kết quả TH chi phí xây dựng)
            chkHeSo.Checked = true;

            // Kiểm tra xem trong Workbook hiện tại đã có các bảng kỹ thuật chưa (đã xuất từ bước 'Áp giá vào dự toán')
            bool daCoBangKyThuat = false;
            try
            {
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                var wb = app?.ActiveWorkbook;
                if (wb != null)
                {
                    foreach (Microsoft.Office.Interop.Excel.Worksheet ws in wb.Sheets)
                    {
                        if (ws.Name.StartsWith("PhanTich", StringComparison.OrdinalIgnoreCase) ||
                            ws.Name.StartsWith("TH_VatLieu", StringComparison.OrdinalIgnoreCase))
                        {
                            daCoBangKyThuat = true;
                            break;
                        }
                    }
                }
            }
            catch { }

            // Nếu đã có các bảng kỹ thuật (được tạo khi Áp giá), mặc định không tích lại để tránh ghi đè không cần thiết.
            // Nếu chưa có, vẫn tích chọn để xuất trọn bộ hồ sơ.
            bool chonBangKyThuat = !daCoBangKyThuat;
            chkDuToan.Checked = chonBangKyThuat;
            chkPhanTich.Checked = chonBangKyThuat;
            chkTHVL.Checked = chonBangKyThuat;
            chkTHNC.Checked = chonBangKyThuat;
            chkTHMay.Checked = chonBangKyThuat;
            chkCuocVC.Checked = chonBangKyThuat;
            chkHeSo.Checked = true;
        }


        private void BtnXuat_Click(object sender, EventArgs e)
        {
            LuaChon.XuatTongMucDauTu = chkTMDT.Checked;
            LuaChon.XuatTongHopDuToan = chkTHDT.Checked;
            LuaChon.XuatChiPhiXayDung = chkTHCPXD.Checked;
            LuaChon.XuatDuToanChiTiet = chkDuToan.Checked;
            LuaChon.XuatPhanTichDonGia = chkPhanTich.Checked;
            LuaChon.XuatTongHopVatLieu = chkTHVL.Checked;
            LuaChon.XuatTongHopNhanCong = chkTHNC.Checked;
            LuaChon.XuatTongHopCaMay = chkTHMay.Checked;
            LuaChon.XuatChietTinhCuocVC = chkCuocVC.Checked;
            LuaChon.XuatHeSoDieuChinh = chkHeSo.Checked;

            if (!LuaChon.CoItNhatMotBangDuocChon())
            {
                MessageBox.Show("Vui lòng tích chọn ít nhất một bảng biểu để xuất sang Excel.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
