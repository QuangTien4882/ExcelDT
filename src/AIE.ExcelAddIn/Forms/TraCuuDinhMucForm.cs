using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using AIE.Data;
using AIE.Data.Repositories;
using AIE.Core.Models;
using AIE.ExcelAddIn.Helpers;
using AIE.ExcelAddIn.Services;

namespace AIE.ExcelAddIn.Forms
{
    public class TraCuuDinhMucForm : Form
    {
        private TextBox txtSearch;
        private Button btnSearch;
        private Button btnInsert;
        private CheckBox chkAutoClose;
        private Button btnClose;
        private Label lblStatus;
        private Button btnPhuLuc;
        private ContextMenuStrip menuPhuLuc;
        private DataGridView dgvCongTac;
        private DataGridView dgvHaoPhi;
        private DatabaseManager _dbManager;
        private CongTacRepository _repo;
        private System.Windows.Forms.Timer _searchTimer;
        private BindingSource _bsCongTac;
        private List<CongTacXayDung> _allCongTac;
        
        private string _settingsPath;
        private string _initialSearchKeyword = "";
        private int _targetRow = -1;
        private int _firstInsertedRow = -1;

        public TraCuuDinhMucForm()
        {
            _dbManager = new DatabaseManager();
            _repo = new CongTacRepository(_dbManager.Context);
            _bsCongTac = new BindingSource();
            
            string appData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AIE_DuToan");
            if (!Directory.Exists(appData)) Directory.CreateDirectory(appData);
            _settingsPath = Path.Combine(appData, "grid_settings.txt");

            InitializeComponents();
            FormStateHelper.Attach(this);
        }

        public TraCuuDinhMucForm(string initialSearchKeyword, int targetRow = -1) : this()
        {
            _initialSearchKeyword = initialSearchKeyword ?? "";
            _targetRow = targetRow;
        }

        private void InitializeComponents()
        {
            this.Text = "Tra cứu định mức (Thông tư 38)";
            
            var workingArea = Screen.PrimaryScreen.WorkingArea;
            this.Size = new Size((int)(workingArea.Width * 0.9), (int)(workingArea.Height * 0.9));
            this.StartPosition = FormStartPosition.CenterScreen;
            
            this.Font = new Font("Be Vietnam Pro", 9.5F);
            this.BackColor = Color.FromArgb(245, 246, 250);
            this.ShowIcon = false;
            this.ShowInTaskbar = true;
            this.KeyPreview = true;
            this.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.Escape)
                {
                    this.Close();
                }
            };

            // --- HEADER PANEL ---
            var pnlTop = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.White };
            pnlTop.Paint += (s, e) => {
                ControlPaint.DrawBorder(e.Graphics, pnlTop.ClientRectangle, Color.FromArgb(220, 221, 225), ButtonBorderStyle.Solid);
            };
            
            var lblSearch = new Label { 
                Text = "Tìm kiếm:", 
                AutoSize = true, 
                Location = new Point(20, 20),
                Font = new Font("Be Vietnam Pro SemiBold", 9.5F),
                ForeColor = Color.FromArgb(47, 54, 64)
            };

            txtSearch = new TextBox { 
                Location = new Point(95, 17),
                Width = 230,
                Font = new Font("Be Vietnam Pro", 9.5F),
                BorderStyle = BorderStyle.FixedSingle
            };

            _searchTimer = new System.Windows.Forms.Timer { Interval = 300 };
            _searchTimer.Tick += (s, e) => {
                _searchTimer.Stop();
                FilterData();
            };

            txtSearch.TextChanged += (s, e) => {
                _searchTimer.Stop();
                _searchTimer.Start();
            };
            txtSearch.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true;
                    FilterData();
                    if (dgvCongTac.Rows.Count > 0)
                    {
                        dgvCongTac.Focus();
                    }
                }
                else if (e.KeyCode == Keys.Down)
                {
                    if (dgvCongTac.Rows.Count > 0)
                    {
                        dgvCongTac.Focus();
                    }
                }
            };
            
            btnSearch = new Button { 
                Text = "🔍", 
                Location = new Point(335, 15),
                Width = 36,
                Height = 30,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 168, 255),
                ForeColor = Color.White,
                Cursor = Cursors.Hand,
                Font = new Font("Segoe UI Emoji", 10F)
            };
            btnSearch.FlatAppearance.BorderSize = 0;
            ToolTip tt = new ToolTip();
            tt.SetToolTip(btnSearch, "Tìm kiếm định mức");
            btnSearch.Click += (s, e) => FilterData();

            btnInsert = new Button {
                Text = "📥 Chèn vào dự toán",
                Location = new Point(380, 14),
                Width = 150,
                Height = 32,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(46, 204, 113),
                ForeColor = Color.White,
                Font = new Font("Be Vietnam Pro SemiBold", 9.5F),
                Cursor = Cursors.Hand
            };
            btnInsert.FlatAppearance.BorderSize = 0;
            tt.SetToolTip(btnInsert, "Chèn các công tác đang chọn vào bảng dự toán Excel (Enter)");
            btnInsert.Click += (s, e) => InsertSelectedCongTacIntoExcel();

            chkAutoClose = new CheckBox {
                Text = "Đóng sau khi chèn",
                Location = new Point(540, 20),
                AutoSize = true,
                Font = new Font("Be Vietnam Pro", 9F),
                ForeColor = Color.FromArgb(47, 54, 64),
                Checked = false
            };
            tt.SetToolTip(chkAutoClose, "Nếu chọn: Tự động đóng bảng tra sau khi chèn công tác");

            var lblPhuLuc = new Label { 
                Text = "Bộ định mức:", 
                AutoSize = true, 
                Font = new Font("Be Vietnam Pro SemiBold", 9.5F),
                ForeColor = Color.FromArgb(47, 54, 64)
            };

            btnPhuLuc = new Button {
                Text = "Chọn Phụ Lục ▼",
                Location = new Point(130, 14),
                Width = 160,
                Height = 32,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                Cursor = Cursors.Hand
            };
            btnPhuLuc.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);

            btnClose = new Button {
                Text = "❌ Đóng (Esc)",
                Width = 110,
                Height = 32,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(47, 54, 64),
                Font = new Font("Be Vietnam Pro", 9F),
                Cursor = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
            btnClose.Click += (s, e) => this.Close();

            menuPhuLuc = new ContextMenuStrip();
            string[] phuLucs = { 
                "Phụ lục I - Khảo sát (Mã C)", 
                "Phụ lục II - Xây dựng (Mã A)", 
                "Phụ lục III - Lắp đặt hệ thống (Mã B)", 
                "Phụ lục IV - Lắp đặt máy và TB (Mã M)", 
                "Phụ lục V - Thí nghiệm (Mã D)", 
                "Phụ lục VI - Sửa chữa (Mã S)" 
            };
            
            foreach (var pl in phuLucs)
            {
                var item = new ToolStripMenuItem(pl) { CheckOnClick = true, Checked = true };
                item.CheckedChanged += (s, e) => FilterData();
                menuPhuLuc.Items.Add(item);
            }
            
            menuPhuLuc.Items.Add(new ToolStripSeparator());
            var itemBoChon = new ToolStripMenuItem("Bỏ chọn toàn bộ (Clear all)");
            itemBoChon.Click += (s, e) => {
                foreach (ToolStripItem it in menuPhuLuc.Items)
                {
                    if (it is ToolStripMenuItem tsi && tsi != itemBoChon)
                    {
                        tsi.Checked = false;
                    }
                }
                FilterData();
            };
            menuPhuLuc.Items.Add(itemBoChon);
            
            btnPhuLuc.Click += (s, e) => menuPhuLuc.Show(btnPhuLuc, new Point(0, btnPhuLuc.Height));
            
            pnlTop.Controls.Add(lblSearch);
            pnlTop.Controls.Add(txtSearch);
            pnlTop.Controls.Add(btnSearch);
            pnlTop.Controls.Add(btnInsert);
            pnlTop.Controls.Add(chkAutoClose);
            pnlTop.Controls.Add(lblPhuLuc);
            pnlTop.Controls.Add(btnPhuLuc);
            pnlTop.Controls.Add(btnClose);

            // Tính vị trí PhuLuc & Close controls bám sát bên phải
            Action layoutSearch = () => {
                int right = pnlTop.ClientSize.Width - 15;
                btnClose.Location = new Point(right - btnClose.Width, 14);
                btnPhuLuc.Location = new Point(btnClose.Left - btnPhuLuc.Width - 10, 14);
                lblPhuLuc.Location = new Point(btnPhuLuc.Left - lblPhuLuc.Width - 10, 20);
            };
            pnlTop.Resize += (s, e) => layoutSearch();
            pnlTop.Layout += (s, e) => layoutSearch();

            // --- SPLIT CONTAINER ---
            var splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                FixedPanel = FixedPanel.None,
                SplitterDistance = (int)((workingArea.Height * 0.9 - 95) * 0.6),
                SplitterWidth = 6,
                BackColor = Color.FromArgb(200, 200, 200)
            };

            // DGV Công tác
            dgvCongTac = CreateModernGrid();
            dgvCongTac.SelectionChanged += DgvCongTac_SelectionChanged;
            dgvCongTac.CellDoubleClick += DgvCongTac_CellDoubleClick;
            dgvCongTac.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    InsertSelectedCongTacIntoExcel();
                }
            };
            SetupCongTacColumns();
            
            var pnlGrid1 = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10), BackColor = Color.FromArgb(245, 246, 250) };
            pnlGrid1.Controls.Add(dgvCongTac);
            splitContainer.Panel1.Controls.Add(pnlGrid1);

            // DGV Hao phí
            dgvHaoPhi = CreateModernGrid();
            SetupHaoPhiColumns();
            
            var pnlGrid2 = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10), BackColor = Color.FromArgb(245, 246, 250) };
            pnlGrid2.Controls.Add(dgvHaoPhi);
            splitContainer.Panel2.Controls.Add(pnlGrid2);

            // --- BOTTOM PANEL (STATUS) ---
            var pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 32, BackColor = Color.FromArgb(240, 243, 246) };
            pnlBottom.Paint += (s, e) => {
                ControlPaint.DrawBorder(e.Graphics, pnlBottom.ClientRectangle, Color.FromArgb(220, 221, 225), ButtonBorderStyle.Solid);
            };
            lblStatus = new Label {
                Text = "💡 Mẹo: Giữ Shift hoặc Ctrl để chọn nhiều công tác rồi bấm 'Chèn vào dự toán'. Bấm Esc để đóng bảng.",
                AutoSize = true,
                Location = new Point(15, 7),
                Font = new Font("Be Vietnam Pro", 8.5F),
                ForeColor = Color.FromArgb(90, 100, 110)
            };
            pnlBottom.Controls.Add(lblStatus);

            this.Controls.Add(splitContainer);
            this.Controls.Add(pnlBottom);
            this.Controls.Add(pnlTop);

            this.Load += TraCuuDinhMucForm_Load;
            this.FormClosing += TraCuuDinhMucForm_FormClosing;
        }

        private DataGridView CreateModernGrid()
        {
            return new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                CellBorderStyle = DataGridViewCellBorderStyle.Single,
                GridColor = Color.FromArgb(200, 200, 200),
                EnableHeadersVisualStyles = false,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single,
                ColumnHeadersHeight = 35,
                RowTemplate = { Height = 30 },
                DefaultCellStyle = new DataGridViewCellStyle { Font = new Font("Be Vietnam Pro", 9F) }
            };
        }

        private void SetupCongTacColumns()
        {
            dgvCongTac.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);
            dgvCongTac.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            dgvCongTac.ColumnHeadersDefaultCellStyle.Font = new Font("Be Vietnam Pro SemiBold", 9.5F);
            dgvCongTac.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            
            // Xám ghi selection
            dgvCongTac.DefaultCellStyle.SelectionBackColor = Color.FromArgb(210, 215, 220);
            dgvCongTac.DefaultCellStyle.SelectionForeColor = Color.Black;

            dgvCongTac.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaHieu", DataPropertyName = "MaHieu", HeaderText = "Mã hiệu", Width = 100 });
            dgvCongTac.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenCongTac", DataPropertyName = "TenCongTac", HeaderText = "Tên công tác xây dựng", Width = 500 });
            var colDonVi = new DataGridViewTextBoxColumn { Name = "DonVi", DataPropertyName = "DonVi", HeaderText = "Đơn vị", Width = 100 };
            colDonVi.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvCongTac.Columns.Add(colDonVi);
        }

        private void DgvCongTac_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                InsertSelectedCongTacIntoExcel();
            }
        }

        private void InsertSelectedCongTacIntoExcel()
        {
            var selectedRows = dgvCongTac.SelectedRows.Cast<DataGridViewRow>()
                                .OrderBy(r => r.Index)
                                .ToList();

            if (selectedRows.Count == 0)
            {
                if (dgvCongTac.CurrentRow != null)
                {
                    selectedRows.Add(dgvCongTac.CurrentRow);
                }
                else
                {
                    MessageBox.Show("Vui lòng chọn ít nhất một công tác để chèn.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }

            var listCongTac = selectedRows
                                .Select(r => r.DataBoundItem as CongTacXayDung)
                                .Where(ct => ct != null)
                                .ToList();

            if (listCongTac.Count == 0) return;

            try
            {
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                var ws = app.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                if (ws == null) return;

                // Tạm thời dừng AutoLookup để không kích hoạt vòng lặp sự kiện
                SheetAutoLookupService.IsSuspended = true;
                bool oldUpdating = app.ScreenUpdating;
                app.ScreenUpdating = false;

                try
                {
                    int curRow = _targetRow;
                    if (curRow <= 0)
                    {
                        var activeCell = app.ActiveCell;
                        curRow = (activeCell != null && activeCell.Row >= 6) ? activeCell.Row : 6;
                    }

                    int targetRow = curRow;
                    int startInsertedRow = targetRow;
                    if (_firstInsertedRow <= 0)
                    {
                        _firstInsertedRow = startInsertedRow;
                    }

                    for (int i = 0; i < listCongTac.Count; i++)
                    {
                        var congTac = listCongTac[i];

                        string cellB = ws.Cells[targetRow, 2]?.Value2?.ToString()?.Trim() ?? "";
                        string cellC = ws.Cells[targetRow, 3]?.Value2?.ToString()?.Trim() ?? "";

                        // Nếu không phải công tác đầu tiên HOẶC dòng hiện tại đã có dữ liệu / tiêu đề
                        if (i > 0 || (!string.IsNullOrEmpty(cellB) || !string.IsNullOrEmpty(cellC)))
                        {
                            if (i > 0 || cellC.StartsWith("TỔNG CỘNG", StringComparison.OrdinalIgnoreCase) || !string.IsNullOrEmpty(cellB))
                            {
                                if (i > 0) targetRow++;
                                Microsoft.Office.Interop.Excel.Range rowToInsert = ws.Rows[targetRow] as Microsoft.Office.Interop.Excel.Range;
                                rowToInsert?.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                            }
                        }

                        // Tính STT tạm
                        int stt = 1;
                        for (int rPrev = targetRow - 1; rPrev >= 5; rPrev--)
                        {
                            object prevCell = ws.Cells[rPrev, 1]?.Value2;
                            if (prevCell != null && int.TryParse(prevCell.ToString(), out int prevStt))
                            {
                                stt = prevStt + 1;
                                break;
                            }
                        }

                        ws.Cells[targetRow, 1].Value2 = stt;
                        ws.Cells[targetRow, 2].Value2 = congTac.MaHieu;
                        ws.Cells[targetRow, 3].Value2 = congTac.TenCongTac;
                        ws.Cells[targetRow, 4].Value2 = congTac.DonVi;

                        // Công thức tính Thành tiền: VL (I), NC (J), Máy (K)
                        ws.Cells[targetRow, 9].Formula = $"=ROUND(E{targetRow}*F{targetRow}, 0)";
                        ws.Cells[targetRow, 10].Formula = $"=ROUND(E{targetRow}*G{targetRow}, 0)";
                        ws.Cells[targetRow, 11].Formula = $"=ROUND(E{targetRow}*H{targetRow}, 0)";

                        // Định dạng dòng công tác
                        var rowRange = ws.Range[ws.Cells[targetRow, 1], ws.Cells[targetRow, 11]];
                        rowRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        rowRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        rowRange.Font.Bold = false;
                        rowRange.Interior.ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone;

                        ws.Cells[targetRow, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        ws.Cells[targetRow, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        ws.Cells[targetRow, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignJustify;
                        ws.Cells[targetRow, 3].WrapText = true;
                        ws.Cells[targetRow, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        ExcelFormatHelper.ApplyQuantityFormat(ws.Cells[targetRow, 5], 2);
                        ExcelFormatHelper.ApplyIntegerFormat(ws.Range[ws.Cells[targetRow, 6], ws.Cells[targetRow, 11]]);
                        ws.Range[ws.Cells[targetRow, 5], ws.Cells[targetRow, 11]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    }

                    // Đóng khung toàn bộ bảng từ dòng 4 đến dòng vừa chèn
                    var wholeTable = ws.Range[ws.Cells[4, 1], ws.Cells[targetRow, 11]];
                    wholeTable.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    // Tự động đánh số thứ tự (STT) lại liên tục, chuẩn xác cho toàn bộ các dòng công tác
                    LapDuToanExcelService.DanhLaiSTTCongTac(ws);

                    // Cập nhật _targetRow chuẩn cho lần chèn tiếp theo
                    _targetRow = targetRow + 1;

                    if (lblStatus != null)
                    {
                        lblStatus.Text = $"✅ Đã chèn {listCongTac.Count} công tác vào dòng {startInsertedRow} - {targetRow}! Tiếp tục chọn hoặc bấm Esc để đóng.";
                        lblStatus.ForeColor = Color.FromArgb(39, 174, 96);
                    }

                    if (chkAutoClose != null && chkAutoClose.Checked)
                    {
                        this.Close();
                    }
                    else
                    {
                        txtSearch.Focus();
                        txtSearch.SelectAll();
                    }
                }
                finally
                {
                    app.ScreenUpdating = oldUpdating;
                    SheetAutoLookupService.IsSuspended = false;
                }
            }
            catch (Exception ex)
            {
                SheetAutoLookupService.IsSuspended = false;
                MessageBox.Show($"Không thể chèn vào Excel: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetupHaoPhiColumns()
        {
            dgvHaoPhi.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);
            dgvHaoPhi.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            dgvHaoPhi.ColumnHeadersDefaultCellStyle.Font = new Font("Be Vietnam Pro SemiBold", 9.5F);
            dgvHaoPhi.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            
            // Xám ghi selection
            dgvHaoPhi.DefaultCellStyle.SelectionBackColor = Color.FromArgb(210, 215, 220);
            dgvHaoPhi.DefaultCellStyle.SelectionForeColor = Color.Black;

            var colLoai = new DataGridViewTextBoxColumn { Name = "LoaiHaoPhiStr", DataPropertyName = "LoaiHaoPhiStr", HeaderText = "Loại hao phí", Width = 120 };
            colLoai.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvHaoPhi.Columns.Add(colLoai);
            
            dgvHaoPhi.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaHieuHP", DataPropertyName = "MaHieuHP", HeaderText = "Mã vật tư", Width = 100 });
            dgvHaoPhi.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenHaoPhi", DataPropertyName = "TenHaoPhi", HeaderText = "Tên vật tư / Nhân công / Máy TC", Width = 500 });
            
            var colDonVi = new DataGridViewTextBoxColumn { Name = "DonVi", DataPropertyName = "DonVi", HeaderText = "Đơn vị", Width = 100 };
            colDonVi.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvHaoPhi.Columns.Add(colDonVi);
            
            var colDinhMuc = new DataGridViewTextBoxColumn { Name = "DinhMuc", DataPropertyName = "DinhMuc", HeaderText = "Định mức hao phí", Width = 150 };
            colDinhMuc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            colDinhMuc.DefaultCellStyle.Format = "N4";
            dgvHaoPhi.Columns.Add(colDinhMuc);
        }

        private void TraCuuDinhMucForm_Load(object sender, EventArgs e)
        {
            try
            {
                _allCongTac = _repo.GetAllCongTac().OrderBy(x => x.MaHieu).ToList();
                _bsCongTac.DataSource = _allCongTac;
                dgvCongTac.DataSource = _bsCongTac;
                
                LoadColumnSettings();

                if (!string.IsNullOrWhiteSpace(_initialSearchKeyword))
                {
                    txtSearch.Text = _initialSearchKeyword;
                    FilterData();
                    if (dgvCongTac.Rows.Count > 0)
                    {
                        dgvCongTac.Focus();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FilterData()
        {
            var keyword = txtSearch.Text.Trim().ToLower();
            
            var allowedPrefixes = new List<string>();
            if (((ToolStripMenuItem)menuPhuLuc.Items[0]).Checked) allowedPrefixes.Add("c");
            if (((ToolStripMenuItem)menuPhuLuc.Items[1]).Checked) allowedPrefixes.Add("a");
            if (((ToolStripMenuItem)menuPhuLuc.Items[2]).Checked) allowedPrefixes.Add("b");
            if (((ToolStripMenuItem)menuPhuLuc.Items[3]).Checked) allowedPrefixes.Add("m");
            if (((ToolStripMenuItem)menuPhuLuc.Items[4]).Checked) allowedPrefixes.Add("d");
            if (((ToolStripMenuItem)menuPhuLuc.Items[5]).Checked) allowedPrefixes.Add("s");

            var filtered = _allCongTac.Where(x => 
                (x.MaHieu != null && allowedPrefixes.Any(p => x.MaHieu.ToLower().StartsWith(p))) &&
                (string.IsNullOrEmpty(keyword) || 
                 MatchAllKeywords(keyword, x.MaHieu, x.TenCongTac))
            ).ToList();

            _bsCongTac.DataSource = filtered;
        }

        /// <summary>
        /// Tìm kiếm thông minh: tách từ khoá thành các từ rời, yêu cầu TẤT CẢ các từ đều
        /// xuất hiện trong chuỗi ghép (MaHieu + TenCongTac), không phân biệt thứ tự.
        /// Ví dụ: "máy đào bánh xích" sẽ tìm thấy "Máy đào 1 gầu, bánh xích - dung tích gầu: 0,4m3"
        /// </summary>
        private static bool MatchAllKeywords(string keyword, params string[] fields)
        {
            var combined = string.Join(" ", fields.Where(f => f != null)).ToLower();
            var words = keyword.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return words.All(w => combined.Contains(w));
        }

        private void DgvCongTac_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvCongTac.SelectedRows.Count > 0)
            {
                var ct = dgvCongTac.SelectedRows[0].DataBoundItem as CongTacXayDung;
                if (ct != null && ct.DanhSachHaoPhi != null)
                {
                    var viewList = ct.DanhSachHaoPhi.Select(hp => new {
                        LoaiHaoPhiStr = hp.LoaiHaoPhi == AIE.Core.Enums.LoaiHaoPhi.VL ? "VL" : 
                                       (hp.LoaiHaoPhi == AIE.Core.Enums.LoaiHaoPhi.NC ? "NC" : "MTC"),
                        hp.MaHieuHP,
                        hp.TenHaoPhi,
                        hp.DonVi,
                        hp.DinhMuc
                    }).ToList();
                    
                    dgvHaoPhi.DataSource = viewList;
                    dgvHaoPhi.AutoFit();
                }
                else
                {
                    dgvHaoPhi.DataSource = null;
                }
            }
            else
            {
                dgvHaoPhi.DataSource = null;
            }
        }

        private void TraCuuDinhMucForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveColumnSettings();
            if (_firstInsertedRow > 0)
            {
                try
                {
                    var app = (Microsoft.Office.Interop.Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                    var ws = app.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                    ws?.Activate();
                    var klCell = ws?.Cells[_firstInsertedRow, 5] as Microsoft.Office.Interop.Excel.Range;
                    klCell?.Select();
                }
                catch { }
            }
        }

        // --- SAVE / LOAD COLUMN WIDTHS ---
        private void SaveColumnSettings()
        {
            try
            {
                var lines = new List<string>();
                foreach (DataGridViewColumn col in dgvCongTac.Columns)
                    if (col.AutoSizeMode == DataGridViewAutoSizeColumnMode.NotSet || col.AutoSizeMode == DataGridViewAutoSizeColumnMode.None)
                        lines.Add($"C|{col.Name}|{col.Width}");

                foreach (DataGridViewColumn col in dgvHaoPhi.Columns)
                    if (col.AutoSizeMode == DataGridViewAutoSizeColumnMode.NotSet || col.AutoSizeMode == DataGridViewAutoSizeColumnMode.None)
                        lines.Add($"H|{col.Name}|{col.Width}");

                File.WriteAllLines(_settingsPath, lines);
            }
            catch { /* ignore */ }
        }

        private void LoadColumnSettings()
        {
            try
            {
                if (File.Exists(_settingsPath))
                {
                    var lines = File.ReadAllLines(_settingsPath);
                    foreach (var line in lines)
                    {
                        var parts = line.Split('|');
                        if (parts.Length == 3 && int.TryParse(parts[2], out int width))
                        {
                            if (parts[0] == "C" && dgvCongTac.Columns[parts[1]] != null)
                                dgvCongTac.Columns[parts[1]].Width = width;
                            else if (parts[0] == "H" && dgvHaoPhi.Columns[parts[1]] != null)
                                dgvHaoPhi.Columns[parts[1]].Width = width;
                        }
                    }
                }
            }
            catch { /* ignore */ }
        }
    }
}
