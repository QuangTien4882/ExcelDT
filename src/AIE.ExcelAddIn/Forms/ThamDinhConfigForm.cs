using AIE.Core.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace AIE.ExcelAddIn.Forms;

public class ThamDinhConfigForm : Form
{
    private ComboBox cbSheetName;
    private ComboBox cbColMaHieu;
    private ComboBox cbColTen;
    private ComboBox cbColDonVi;
    private ComboBox cbColDinhMuc;
    private ComboBox cbColDonGia;
    private ComboBox cbColThanhTien;
    private NumericUpDown nudDongBatDau;
    private Button btnOk;
    private Button btnCancel;
    private ProgressBar progressBar;
    private Label lblStatus;

    private ComboBox cbBoDonGia;
    private ComboBox cbVung;
    public ThamDinhConfig ResultConfig { get; private set; }
    public int? SelectedBoDonGiaId { get; private set; }

    #region TCVN3 to Unicode Conversion

    /// <summary>
    /// Bảng chuyển đổi ký tự TCVN3 (.VnTime, .VnArial, ...) sang Unicode.
    /// Khi Excel dùng font TCVN3, giá trị ô đọc qua COM sẽ trả về ký tự Latin sai lệch.
    /// Ví dụ: "®¬n vÞ" thay vì "Đơn vị".
    /// </summary>
    private static readonly Dictionary<char, char> _tcvn3Map = new Dictionary<char, char>
    {
        // a + dấu
        {'\u00B8', 'á'}, {'\u00B5', 'à'}, {'\u00B6', 'ả'}, {'\u00B7', 'ã'}, {'\u00B9', 'ạ'},
        // â + dấu
        {'\u00A9', 'â'}, {'\u00CA', 'ấ'}, {'\u00C7', 'ầ'}, {'\u00C8', 'ẩ'}, {'\u00C9', 'ẫ'}, {'\u00CB', 'ậ'},
        // ă + dấu
        {'\u00A8', 'ă'}, {'\u00BE', 'ắ'}, {'\u00BB', 'ằ'}, {'\u00BC', 'ẳ'}, {'\u00BD', 'ẵ'}, {'\u00C6', 'ặ'},
        // e + dấu
        {'\u00D0', 'é'}, {'\u00CC', 'è'}, {'\u00CE', 'ẻ'}, {'\u00CF', 'ẽ'}, {'\u00D1', 'ẹ'},
        // ê + dấu
        {'\u00AA', 'ê'}, {'\u00D5', 'ế'}, {'\u00D2', 'ề'}, {'\u00D3', 'ể'}, {'\u00D4', 'ễ'}, {'\u00D6', 'ệ'},
        // i + dấu
        {'\u00DD', 'í'}, {'\u00D7', 'ì'}, {'\u00D8', 'ỉ'}, {'\u00DC', 'ĩ'}, {'\u00DE', 'ị'},
        // o + dấu
        {'\u00E3', 'ó'}, {'\u00DF', 'ò'}, {'\u00E1', 'ỏ'}, {'\u00E2', 'õ'}, {'\u00E4', 'ọ'},
        // ô + dấu
        {'\u00AB', 'ô'}, {'\u00E8', 'ố'}, {'\u00E5', 'ồ'}, {'\u00E6', 'ổ'}, {'\u00E7', 'ỗ'}, {'\u00E9', 'ộ'},
        // ơ + dấu
        {'\u00AC', 'ơ'}, {'\u00ED', 'ớ'}, {'\u00EA', 'ờ'}, {'\u00EB', 'ở'}, {'\u00EC', 'ỡ'}, {'\u00EE', 'ợ'},
        // u + dấu
        {'\u00F3', 'ú'}, {'\u00EF', 'ù'}, {'\u00F1', 'ủ'}, {'\u00F2', 'ũ'}, {'\u00F4', 'ụ'},
        // ư + dấu
        {'\u00B0', 'ư'}, {'\u00F8', 'ứ'}, {'\u00F5', 'ừ'}, {'\u00F6', 'ử'}, {'\u00F7', 'ữ'}, {'\u00F9', 'ự'},
        // y + dấu
        {'\u00FD', 'ý'}, {'\u00FA', 'ỳ'}, {'\u00FB', 'ỷ'}, {'\u00FC', 'ỹ'}, {'\u00FE', 'ỵ'},
        // đ
        {'\u00AE', 'đ'}, {'\u00A7', 'đ'},
    };

    /// <summary>
    /// Kiểm tra chuỗi có phải mã hóa TCVN3 hay không.
    /// Dựa vào các ký tự chỉ xuất hiện trong TCVN3, không bao giờ xuất hiện trong text Unicode tiếng Việt thông thường.
    /// </summary>
    private static bool IsTcvn3(string text)
    {
        foreach (char c in text)
        {
            // ® (đ), ¬ (ơ), « (ô), Þ (ị), ¹ (ạ), ¨ (ă), ª (ê), © (â) - đều là ký tự TCVN3 đặc trưng
            if (c == '\u00AE' || c == '\u00AC' || c == '\u00AB' ||
                c == '\u00DE' || c == '\u00B9' || c == '\u00A8' ||
                c == '\u00AA' || c == '\u00A9' || c == '\u00A7')
                return true;
        }
        return false;
    }

    /// <summary>
    /// Chuyển đổi chuỗi TCVN3 sang Unicode.
    /// </summary>
    private static string ConvertTcvn3(string input)
    {
        if (string.IsNullOrEmpty(input)) return input;
        var sb = new StringBuilder(input.Length);
        foreach (char c in input)
        {
            if (_tcvn3Map.TryGetValue(c, out char u))
                sb.Append(u);
            else
                sb.Append(c);
        }
        return sb.ToString();
    }

    /// <summary>
    /// Chuẩn hóa text ô Excel: xử lý xuống dòng, khoảng trắng thừa, chuyển TCVN3 nếu cần, rồi lowercase.
    /// </summary>
    private static string NormalizeCell(string raw)
    {
        if (string.IsNullOrEmpty(raw)) return "";
        string text = raw.Replace("\n", " ").Replace("\r", " ").Trim();
        if (IsTcvn3(text))
            text = ConvertTcvn3(text);
        text = text.ToLower();
        while (text.Contains("  ")) text = text.Replace("  ", " ");
        return text;
    }

    #endregion

    public ThamDinhConfigForm(List<string> sheetNames, int? defaultBoDonGiaId = null)
    {
        InitializeComponent();
        AIE.ExcelAddIn.Helpers.FormStateHelper.Attach(this);
        PopulateColumns();

        // Nạp danh sách sheet
        foreach (var name in sheetNames)
            cbSheetName.Items.Add(name);

        cbSheetName.SelectedIndexChanged += CbSheetName_SelectedIndexChanged;

        LoadBoDonGia(defaultBoDonGiaId);
        SelectBestDefaultSheet(sheetNames);
    }

    private void SelectBestDefaultSheet(List<string> sheetNames)
    {
        try
        {
            var app = (Microsoft.Office.Interop.Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
            string activeSheet = app?.ActiveWorkbook?.ActiveSheet is Microsoft.Office.Interop.Excel.Worksheet ws ? ws.Name : null;
            
            string best = null;
            if (!string.IsNullOrEmpty(activeSheet) && sheetNames.Contains(activeSheet))
            {
                best = activeSheet;
            }
            else
            {
                best = sheetNames.FirstOrDefault(s => {
                    string sl = s.ToLower();
                    return sl.Contains("dutoan") || sl.Contains("dự toán") || sl.Contains("chiphixd") || sl.Contains("chi phí xd") || sl.Contains("dongia") || sl.Contains("đơn giá") || sl.Contains("phân tích") || sl.Contains("phantich") || sl.Contains("congtrinh") || sl.Contains("dt");
                }) ?? sheetNames.FirstOrDefault();
            }

            if (!string.IsNullOrEmpty(best) && cbSheetName.Items.Contains(best))
            {
                cbSheetName.SelectedItem = best;
            }
        }
        catch { }
    }

    private void LoadBoDonGia(int? defaultBoDonGiaId = null)
    {
        var db = new AIE.Data.DatabaseManager();
        var repo = new AIE.Data.Repositories.BoDonGiaRepository(db.Context);
        var list = repo.GetAll().ToList();
        
        list.Insert(0, new AIE.Data.Repositories.BoDonGiaRepository.BoDonGiaInfo { Id = 0, TenBo = "(Mặc định / Không áp dụng)" });
        cbBoDonGia.DataSource = list;

        if (defaultBoDonGiaId.HasValue && defaultBoDonGiaId.Value > 0 && list.Any(x => x.Id == defaultBoDonGiaId.Value))
        {
            cbBoDonGia.SelectedValue = defaultBoDonGiaId.Value;
        }
        else if (list.Count > 1)
        {
            // Tự động chọn Bộ đơn giá mới nhất được lưu trong CSDL
            cbBoDonGia.SelectedIndex = 1;
        }
    }

    private void InitializeComponent()
    {
        this.Text = "Cấu Hình Đọc Dự Toán Cần Thẩm Định";
        this.StartPosition = FormStartPosition.CenterParent;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = true;
        this.AutoSize = true;
        this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        
        var font = new Font("Be Vietnam Pro", 9.5f);
        this.Font = font;

        TableLayoutPanel tlp = new TableLayoutPanel();
        tlp.ColumnCount = 2;
        tlp.RowCount = 15;
        tlp.AutoSize = true;
        tlp.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        tlp.Padding = new Padding(20, 20, 20, 20);
        tlp.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 350F));

        int row = 0;

        void AddRow(string labelText, Control inputControl)
        {
            Label lbl = new Label { Text = labelText, AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top };
            lbl.Margin = new Padding(0, 7, 20, 0);
            inputControl.Width = 330;
            inputControl.Anchor = AnchorStyles.Left | AnchorStyles.Top;
            inputControl.Margin = new Padding(0, 0, 0, 15);
            tlp.Controls.Add(lbl, 0, row);
            tlp.Controls.Add(inputControl, 1, row);
            row++;
        }

        cbSheetName = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        AddRow("Sheet dữ liệu:", cbSheetName);

        var divider = new Label { BorderStyle = BorderStyle.Fixed3D, Height = 2, Anchor = AnchorStyles.Left | AnchorStyles.Right, Margin = new Padding(0, 10, 0, 20) };
        tlp.Controls.Add(divider, 0, row);
        tlp.SetColumnSpan(divider, 2);
        row++;

        cbColMaHieu = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        AddRow("Cột Mã hiệu:", cbColMaHieu);

        cbColTen = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        AddRow("Cột Tên công tác/vật tư:", cbColTen);

        cbColDonVi = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        AddRow("Cột Đơn vị:", cbColDonVi);

        cbColDinhMuc = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        AddRow("Cột Định mức:", cbColDinhMuc);

        cbColDonGia = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        AddRow("Cột Đơn giá:", cbColDonGia);
        
        cbColThanhTien = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        AddRow("Cột Thành tiền:", cbColThanhTien);

        nudDongBatDau = new NumericUpDown { Minimum = 1, Maximum = 10000, Value = 1 };
        AddRow("Dòng bắt đầu (Data):", nudDongBatDau);

        var divider2 = new Label { BorderStyle = BorderStyle.Fixed3D, Height = 2, Anchor = AnchorStyles.Left | AnchorStyles.Right, Margin = new Padding(0, 10, 0, 20) };
        tlp.Controls.Add(divider2, 0, row);
        tlp.SetColumnSpan(divider2, 2);
        row++;

        cbBoDonGia = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, DisplayMember = "TenBo", ValueMember = "Id", Width = 260 };
        
        var btnDeleteBoDonGia = new Button { Text = "Xóa", Width = 60, Height = cbBoDonGia.Height + 2, BackColor = Color.LightCoral, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
        btnDeleteBoDonGia.FlatAppearance.BorderSize = 0;
        btnDeleteBoDonGia.Click += BtnDeleteBoDonGia_Click;
        
        var panelBoDonGia = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
        panelBoDonGia.Controls.Add(cbBoDonGia);
        panelBoDonGia.Controls.Add(btnDeleteBoDonGia);

        Label lblBoDonGia = new Label { Text = "Chọn Bộ Đơn Giá:", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top, Margin = new Padding(0, 7, 20, 0) };
        tlp.Controls.Add(lblBoDonGia, 0, row);
        tlp.Controls.Add(panelBoDonGia, 1, row);
        row++;

        // Vùng áp dụng
        cbVung = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        var vungListTD = new System.Collections.Generic.List<System.Collections.Generic.KeyValuePair<AIE.Core.Enums.Vung, string>>
        {
            new System.Collections.Generic.KeyValuePair<AIE.Core.Enums.Vung, string>(AIE.Core.Enums.Vung.VungII, "Vùng II"),
            new System.Collections.Generic.KeyValuePair<AIE.Core.Enums.Vung, string>(AIE.Core.Enums.Vung.VungIII, "Vùng III"),
            new System.Collections.Generic.KeyValuePair<AIE.Core.Enums.Vung, string>(AIE.Core.Enums.Vung.VungIV, "Vùng IV"),
            new System.Collections.Generic.KeyValuePair<AIE.Core.Enums.Vung, string>(AIE.Core.Enums.Vung.CuLaoCham, "Cù Lao Chàm")
        };
        cbVung.DisplayMember = "Value";
        cbVung.ValueMember = "Key";
        cbVung.DataSource = vungListTD;
        AddRow("Vùng áp dụng:", cbVung);
        
        var lblVungDesc = new Label { AutoSize = true, MaximumSize = new Size(360, 0), Font = new Font("Be Vietnam Pro", 8.5f, FontStyle.Italic), ForeColor = Color.DimGray, Margin = new Padding(0, -10, 0, 15) };
        cbVung.SelectedIndexChanged += (s, e) =>
        {
            var selectedVung = ((System.Collections.Generic.KeyValuePair<AIE.Core.Enums.Vung, string>)cbVung.SelectedItem).Key;
            switch (selectedVung)
            {
                case AIE.Core.Enums.Vung.VungII: lblVungDesc.Text = "Gồm các phường: Hải Châu, Hòa Cường, Thanh Khê, An Khê, An Hải, Sơn Trà, Ngũ Hành Sơn, Hòa Khánh, Hải Vân, Liên Chiểu, Cẩm Lệ, Hòa Xuân, Tam Kỳ, Quảng Phú, Hương Trà, Bàn Thạch, Hội An, Hội An Đông, Hội An Tây và các xã: Hòa Vang, Hòa Tiến, Bà Nà."; break;
                case AIE.Core.Enums.Vung.VungIII: lblVungDesc.Text = "Gồm các phường: Điện Bàn, Điện Bàn Đông, An Thắng, Điện Bàn Bắc và các xã: Núi Thành, Tam Mỹ, Tam Anh, Đức Phú, Tam Xuân, Tam Hải, Tây Hồ, Chiên Đàn, Phú Ninh, Thăng Bình, Thăng An, Thăng Trường, Thăng Điền, Thăng Phú, Đồng Dương, Quế Sơn Trung, Quế Sơn, Xuân Phú, Nông Sơn, Quế Phước, Duy Nghĩa, Nam Phước, Duy Xuyên, Thu Bồn, Điện Bàn Tây, Gò Nổi, Đại Lộc, Hà Nha, Thượng Đức, Vu Gia, Phú Thuận."; break;
                case AIE.Core.Enums.Vung.VungIV: lblVungDesc.Text = "Gồm các xã: Lãnh Ngọc, Tiên Phước, Xã Thạnh Bình, Sơn Cẩm Hà, Trà Liên, Trà Giáp, Trà Tân, Trà Đốc, Trà My, Nam Trà My, Trà Tập, Trà Vân, Trà Linh, Trà Leng, Thạnh Mỹ, Bến Giằng, Nam Giang, Đắc Pring, La Dêê, La Êê, Sông Vàng, Sông Kôn, Đông Giang, Bến Hiên, Avương, Tây Giang, Hùng Sơn, Hiệp Đức, Việt An, Phước Trà, Khâm Đức, Phước Năng, Phước Chánh, Phước Thành, Phước Hiệp."; break;
                case AIE.Core.Enums.Vung.CuLaoCham: lblVungDesc.Text = "Khu vực xã đảo Tân Hiệp (Cù Lao Chàm)"; break;
            }
        };
        
        tlp.Controls.Add(lblVungDesc, 1, row);
        row++;

        ToolTip toolTip = new ToolTip();

        // Buttons
        FlowLayoutPanel flp = new FlowLayoutPanel();
        flp.FlowDirection = FlowDirection.LeftToRight;
        flp.AutoSize = true;
        flp.Anchor = AnchorStyles.None;
        flp.Margin = new Padding(0, 20, 0, 0);
        flp.WrapContents = false;

        btnOk = new Button
        {
            Text = "Bắt đầu thẩm định (kiểm tra)",
            AutoSize = true,
            MinimumSize = new Size(270, 42),
            Height = 42,
            Padding = new Padding(15, 0, 15, 0),
            BackColor = Color.FromArgb(0, 120, 215),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        btnOk.Font = new Font("Segoe UI", 10f, FontStyle.Bold);
        btnOk.FlatAppearance.BorderSize = 0;
        btnOk.Click += BtnOk_Click;
        toolTip.SetToolTip(btnOk, "Bắt đầu kiểm tra thẩm định dự toán");

        btnCancel = new Button { Text = "Hủy bỏ", Width = 110, Height = 42, FlatStyle = FlatStyle.Flat };
        btnCancel.Font = new Font("Segoe UI", 10f);
        btnCancel.Margin = new Padding(15, 0, 0, 0);
        btnCancel.Click += (s, e) => this.DialogResult = DialogResult.Cancel;
        toolTip.SetToolTip(btnCancel, "Hủy bỏ thao tác");

        flp.Controls.Add(btnOk);
        flp.Controls.Add(btnCancel);
        tlp.Controls.Add(flp, 0, row);
        tlp.SetColumnSpan(flp, 2);
        row++;

        // Progress bar (Marquee)
        progressBar = new ProgressBar();
        progressBar.Style = ProgressBarStyle.Marquee;
        progressBar.MarqueeAnimationSpeed = 30;
        progressBar.Height = 6;
        progressBar.Dock = DockStyle.Fill;
        progressBar.Visible = false;
        progressBar.Margin = new Padding(0, 10, 0, 0);
        tlp.Controls.Add(progressBar, 0, row);
        tlp.SetColumnSpan(progressBar, 2);
        row++;

        // Status label
        lblStatus = new Label();
        lblStatus.Text = "";
        lblStatus.AutoSize = true;
        lblStatus.ForeColor = Color.FromArgb(0, 120, 215);
        lblStatus.Font = new Font("Segoe UI", 9f, FontStyle.Italic);
        lblStatus.Anchor = AnchorStyles.None;
        lblStatus.Visible = false;
        lblStatus.Margin = new Padding(0, 5, 0, 0);
        tlp.Controls.Add(lblStatus, 0, row);
        tlp.SetColumnSpan(lblStatus, 2);

        this.Controls.Add(tlp);
        this.AcceptButton = btnOk;
        this.CancelButton = btnCancel;
    }

    private void BtnDeleteBoDonGia_Click(object sender, EventArgs e)
    {
        if (cbBoDonGia.SelectedValue == null) return;
        int val = (int)cbBoDonGia.SelectedValue;
        if (val == 0)
        {
            MessageBox.Show("Không thể xóa bộ đơn giá mặc định.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var result = MessageBox.Show($"Bạn có chắc chắn muốn xóa Bộ đơn giá '{cbBoDonGia.Text}' không?", "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (result == DialogResult.Yes)
        {
            try
            {
                var db = new AIE.Data.DatabaseManager();
                var repo = new AIE.Data.Repositories.BoDonGiaRepository(db.Context);
                repo.Delete(val);
                MessageBox.Show("Đã xóa bộ đơn giá thành công.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadBoDonGia();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa bộ đơn giá: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private void BtnOk_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(cbSheetName.Text) || string.IsNullOrEmpty(cbColMaHieu.Text))
        {
            MessageBox.Show("Vui lòng chọn Sheet và cột Mã hiệu.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (cbBoDonGia.SelectedItem is AIE.Data.Repositories.BoDonGiaRepository.BoDonGiaInfo boInfo)
        {
            SelectedBoDonGiaId = boInfo.Id > 0 ? boInfo.Id : (int?)null;
        }
        else if (cbBoDonGia.SelectedValue is int val)
        {
            SelectedBoDonGiaId = val > 0 ? val : (int?)null;
        }
        else
        {
            SelectedBoDonGiaId = null;
        }

        btnOk.Enabled = false;
        btnCancel.Enabled = false;
        progressBar.Visible = true;
        lblStatus.Text = "⏳ Đang thực hiện thẩm định dự toán...";
        lblStatus.Visible = true;
        Application.DoEvents();

        try
        {
            ResultConfig = new ThamDinhConfig
            {
                SheetName = cbSheetName.SelectedItem?.ToString(),
                ColMaHieu = cbColMaHieu.SelectedItem?.ToString() ?? "B",
                ColTenCT = cbColTen.SelectedItem?.ToString() ?? "C",
                ColDonVi = cbColDonVi.SelectedItem?.ToString() ?? "D",
                ColDinhMuc = cbColDinhMuc.SelectedItem?.ToString() ?? "E",
                ColDonGia = cbColDonGia.SelectedItem?.ToString() ?? "F",
                ColThanhTien = cbColThanhTien.SelectedItem?.ToString() ?? "G",
                DongBatDau = (int)nudDongBatDau.Value,
                Vung = ((System.Collections.Generic.KeyValuePair<AIE.Core.Enums.Vung, string>)cbVung.SelectedItem).Key
            };
            this.DialogResult = DialogResult.OK;
        }
        catch (Exception ex)
        {
            progressBar.Visible = false;
            lblStatus.Visible = false;
            btnOk.Enabled = true;
            btnCancel.Enabled = true;
            MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void PopulateColumns()
    {
        string[] cols = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
        cbColMaHieu.Items.AddRange(cols);
        cbColTen.Items.AddRange(cols);
        cbColDonVi.Items.AddRange(cols);
        cbColDinhMuc.Items.AddRange(cols);
        cbColDonGia.Items.AddRange(cols);
        cbColThanhTien.Items.AddRange(cols);
        // Không chọn mặc định — chờ user chọn sheet rồi auto-map
    }

    private void CbSheetName_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (cbSheetName.SelectedItem == null) return;
        string sheetName = cbSheetName.SelectedItem.ToString();

        // Reset tất cả
        cbColMaHieu.SelectedIndex = -1;
        cbColTen.SelectedIndex = -1;
        cbColDonVi.SelectedIndex = -1;
        cbColDinhMuc.SelectedIndex = -1;
        cbColDonGia.SelectedIndex = -1;
        cbColThanhTien.SelectedIndex = -1;
        nudDongBatDau.Value = 1;

        try
        {
            var app = (Microsoft.Office.Interop.Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
            if (app == null || app.ActiveWorkbook == null) return;

            Microsoft.Office.Interop.Excel.Worksheet targetSheet = null;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in app.ActiveWorkbook.Worksheets)
            {
                if (sheet.Name == sheetName) { targetSheet = sheet; break; }
            }
            if (targetSheet == null) return;

            // Đọc 20 dòng x 20 cột đầu tiên
            var range = targetSheet.Range[targetSheet.Cells[1, 1], targetSheet.Cells[20, 20]];
            object rawVal = range.Value2;
            if (rawVal == null) return;

            object[,] values = (object[,])rawVal;
            int rStart = values.GetLowerBound(0);
            int rEnd = values.GetUpperBound(0);
            int cStart = values.GetLowerBound(1);
            int cEnd = values.GetUpperBound(1);

            string bestMaHieu = "", bestTen = "", bestDonVi = "", bestDinhMuc = "", bestDonGia = "", bestThanhTien = "";
            int bestRow = -1;
            int bestMatchCount = 0;

            for (int r = rStart; r <= rEnd; r++)
            {
                int matchCount = 0;
                string tmpMaHieu = "", tmpTen = "", tmpDonVi = "", tmpDinhMuc = "", tmpDonGia = "", tmpThanhTien = "";

                for (int c = cStart; c <= cEnd; c++)
                {
                    if (values[r, c] == null) continue;

                    // Chuẩn hóa text: xử lý TCVN3 → Unicode nếu cần, rồi lowercase
                    string cellText = NormalizeCell(values[r, c].ToString());
                    if (string.IsNullOrEmpty(cellText)) continue;

                    int colIdx = c - cStart + 1;
                    string colLetter = GetColumnLetter(colIdx);

                    // === Nhận diện cột ===

                    // Mã hiệu / SH Định mức
                    if (cellText.Contains("mã hiệu") || cellText.Contains("mã công tác") ||
                        cellText.Contains("sh định mức") || cellText.Contains("sh đm") ||
                        cellText.Contains("số hiệu") || cellText.Contains("mã đm") ||
                        cellText.Contains("mã định mức") || cellText == "mã cv" || 
                        (cellText.StartsWith("mã") && cellText.Length < 20))
                    {
                        tmpMaHieu = colLetter;
                        matchCount++;
                    }
                    // Tên công tác / hạng mục
                    else if (cellText.Contains("hạng mục công tác") || cellText.Contains("tên công tác") ||
                             cellText.Contains("hạng mục công việc") || cellText.Contains("nội dung công việc") ||
                             cellText.Contains("tên công việc") || cellText.Contains("danh mục") || 
                             cellText.Contains("tên vật tư") || cellText.Contains("nội dung") || 
                             cellText.Contains("diễn giải"))
                    {
                        tmpTen = colLetter;
                        matchCount++;
                    }
                    // Đơn vị
                    else if (cellText == "đơn vị" || cellText == "đvt" || cellText.Contains("đơn vị tính") || cellText == "đ.vị")
                    {
                        tmpDonVi = colLetter;
                        matchCount++;
                    }
                    // Định mức / Khối lượng / Hao phí
                    else if (cellText == "định mức" || cellText == "khối lượng" || cellText == "kl" || cellText == "hao phí" || cellText == "hệ số" ||
                             (cellText.Contains("định mức") && !cellText.Contains("sh")) ||
                             cellText.Contains("khối lượng") || cellText.Contains("hao phí"))
                    {
                        tmpDinhMuc = colLetter;
                        matchCount++;
                    }
                    // Đơn giá
                    else if (cellText.Contains("đơn giá") || cellText.Contains("giá dự toán") || cellText.Contains("giá vật tư") || cellText == "giá")
                    {
                        tmpDonGia = colLetter;
                        matchCount++;
                    }
                    // Thành tiền
                    else if (cellText.Contains("thành tiền") || cellText == "tiền")
                    {
                        tmpThanhTien = colLetter;
                        matchCount++;
                    }
                }

                // Lưu lại kết quả tốt nhất
                if (matchCount > bestMatchCount)
                {
                    bestMatchCount = matchCount;
                    bestMaHieu = tmpMaHieu;
                    bestTen = tmpTen;
                    bestDonVi = tmpDonVi;
                    bestDinhMuc = tmpDinhMuc;
                    bestDonGia = tmpDonGia;
                    bestThanhTien = tmpThanhTien;
                    bestRow = r - rStart + 1; // Chuyển sang số dòng Excel (1-based)
                }
            }

            // Áp dụng kết quả mapping
            if (bestMatchCount >= 2)
            {
                if (!string.IsNullOrEmpty(bestMaHieu)) cbColMaHieu.SelectedItem = bestMaHieu;
                if (!string.IsNullOrEmpty(bestTen)) cbColTen.SelectedItem = bestTen;
                if (!string.IsNullOrEmpty(bestDonVi)) cbColDonVi.SelectedItem = bestDonVi;
                if (!string.IsNullOrEmpty(bestDinhMuc)) cbColDinhMuc.SelectedItem = bestDinhMuc;
                if (!string.IsNullOrEmpty(bestDonGia)) cbColDonGia.SelectedItem = bestDonGia;
                if (!string.IsNullOrEmpty(bestThanhTien)) cbColThanhTien.SelectedItem = bestThanhTien;

                // Tìm dòng data thực sự: quét từ headerRow + 1 trở đi,
                // tìm dòng đầu tiên có chứa Mã hiệu (pattern: chữ + số, VD: AB.21131)
                int dataRow = bestRow + 1;
                int maHieuColIdx = -1;
                if (!string.IsNullOrEmpty(bestMaHieu))
                    maHieuColIdx = GetColumnIndex(bestMaHieu) + cStart - 1;

                for (int r = bestRow + rStart; r <= rEnd; r++)
                {
                    // Kiểm tra cột Mã hiệu
                    if (maHieuColIdx >= cStart && maHieuColIdx <= cEnd && values[r, maHieuColIdx] != null)
                    {
                        string val = values[r, maHieuColIdx].ToString().Trim();
                        // Nếu là TCVN3 thì convert trước
                        if (IsTcvn3(val)) val = ConvertTcvn3(val);
                        // Mã hiệu thường có dạng: chữ + dấu chấm + số (VD: AB.21131, AF.12345)
                        if (Regex.IsMatch(val, @"^[A-Za-z]{2,4}[\.\s]?\d{3,6}"))
                        {
                            dataRow = r - rStart + 1;
                            break;
                        }
                    }

                    // Hoặc Kiểm tra nếu có ít nhất 3 ô có giá trị trong dòng này (dòng data thực sự)
                    int filledCount = 0;
                    for (int c = cStart; c <= cEnd && filledCount < 4; c++)
                    {
                        if (values[r, c] != null && values[r, c].ToString().Trim().Length > 0)
                            filledCount++;
                    }
                    if (filledCount >= 4 && (r - rStart + 1) > bestRow)
                    {
                        dataRow = r - rStart + 1;
                        break;
                    }
                }

                nudDongBatDau.Value = Math.Max(1, dataRow);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi quét sheet '{sheetName}':\n{ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private string GetColumnLetter(int colIndex)
    {
        int div = colIndex;
        string colLetter = string.Empty;
        while (div > 0)
        {
            int mod = (div - 1) % 26;
            colLetter = (char)(65 + mod) + colLetter;
            div = (div - mod) / 26;
        }
        return colLetter;
    }

    private int GetColumnIndex(string colLetter)
    {
        int idx = 0;
        foreach (char c in colLetter.ToUpper())
        {
            idx = idx * 26 + (c - 'A' + 1);
        }
        return idx;
    }
}
