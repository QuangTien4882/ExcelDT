using System;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Font = System.Drawing.Font;

namespace AIE.ExcelAddIn.Helpers;

public static class ExcelFormatHelper
{
    public static void GetSeparators(Microsoft.Office.Interop.Excel.Application? app, out string thouSep, out string decSep)
    {
        try
        {
            thouSep = app?.International[XlApplicationInternational.xlThousandsSeparator]?.ToString() ?? ".";
            decSep = app?.International[XlApplicationInternational.xlDecimalSeparator]?.ToString() ?? ",";
        }
        catch
        {
            thouSep = ".";
            decSep = ",";
        }
    }

    public static string GetIntegerFormat(Microsoft.Office.Interop.Excel.Application? app)
    {
        GetSeparators(app, out string thouSep, out _);
        return $"#{thouSep}##0;-#{thouSep}##0;\"-\"";
    }

    public static string GetQuantityFormat(Microsoft.Office.Interop.Excel.Application? app, int decimals = 2)
    {
        GetSeparators(app, out string thouSep, out string decSep);
        string zeros = new string('0', decimals);
        return $"#{thouSep}##0{decSep}{zeros};-#{thouSep}##0{decSep}{zeros};\"-\"";
    }

    public static string GetRateFormat(Microsoft.Office.Interop.Excel.Application? app, int decimals = 4)
    {
        GetSeparators(app, out string thouSep, out string decSep);
        string zeros = new string('0', decimals);
        return $"#{thouSep}##0{decSep}{zeros};-#{thouSep}##0{decSep}{zeros};\"-\"";
    }

    public static void ApplyIntegerFormat(Range range)
    {
        if (range == null) return;
        try
        {
            range.NumberFormatLocal = GetIntegerFormat(range.Application);
        }
        catch
        {
            try { range.NumberFormatLocal = "#.##0;-#.##0;\"-\""; } catch { }
        }
    }

    public static void ApplyQuantityFormat(Range range, int decimals = 2)
    {
        if (range == null) return;
        try
        {
            range.NumberFormatLocal = GetQuantityFormat(range.Application, decimals);
        }
        catch
        {
            try { range.NumberFormatLocal = "#.##0,00;-#.##0,00;\"-\""; } catch { }
        }
    }

    public static void ApplyRateFormat(Range range, int decimals = 4)
    {
        if (range == null) return;
        try
        {
            range.NumberFormatLocal = GetRateFormat(range.Application, decimals);
        }
        catch
        {
            try { range.NumberFormatLocal = "#.##0,0000;-#.##0,0000;\"-\""; } catch { }
        }
    }
}

public static class UIHelper
{
    private static readonly PrivateFontCollection _pfc = new PrivateFontCollection();
    private static FontFamily? _beVietnamProFamily = null;

    static UIHelper()
    {
        try
        {
            // 1. Kiểm tra xem Font "Be Vietnam Pro" đã có sẵn trong Windows chưa
            using (var test = new Font("Be Vietnam Pro", 10f))
            {
                if (test.Name == "Be Vietnam Pro")
                {
                    _beVietnamProFamily = test.FontFamily;
                    return;
                }
            }
        }
        catch { }

        try
        {
            // 2. Tìm kiếm font TTF trong thư mục Fonts của User hoặc Project Resources
            string userFontDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Microsoft\Windows\Fonts");
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            
            string[] candidateDirs = new string[]
            {
                userFontDir,
                Path.Combine(appDir, "Resources", "Fonts"),
                @"E:\Soft\Personal-app\AIE\src\AIE.ExcelAddIn\Resources\Fonts"
            };

            foreach (var dir in candidateDirs)
            {
                if (Directory.Exists(dir))
                {
                    string reg = Path.Combine(dir, "BeVietnamPro-Regular.ttf");
                    string bold = Path.Combine(dir, "BeVietnamPro-Bold.ttf");
                    string semi = Path.Combine(dir, "BeVietnamPro-SemiBold.ttf");

                    if (File.Exists(reg)) _pfc.AddFontFile(reg);
                    if (File.Exists(bold)) _pfc.AddFontFile(bold);
                    if (File.Exists(semi)) _pfc.AddFontFile(semi);

                    if (_pfc.Families.Length > 0)
                    {
                        _beVietnamProFamily = _pfc.Families.FirstOrDefault(f => f.Name.IndexOf("Vietnam", StringComparison.OrdinalIgnoreCase) >= 0) ?? _pfc.Families[0];
                        break;
                    }
                }
            }
        }
        catch { }
    }

    public static Font GetFont(float size, FontStyle style = FontStyle.Regular)
    {
        if (_beVietnamProFamily != null)
        {
            try
            {
                return new Font(_beVietnamProFamily, size, style);
            }
            catch { }
        }

        try
        {
            using (var test = new Font("Be Vietnam Pro", size, style))
            {
                if (test.Name == "Be Vietnam Pro")
                    return new Font("Be Vietnam Pro", size, style);
            }
        }
        catch { }

        return new Font("Segoe UI", size, style);
    }

    public static void ApplyStyle(TabControl tabControl)
    {
        tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;
        tabControl.DrawItem -= TabControl_DrawItem;
        tabControl.DrawItem += TabControl_DrawItem;
    }

    private static void TabControl_DrawItem(object? sender, DrawItemEventArgs e)
    {
        if (sender is not TabControl tabControl) return;

        var g = e.Graphics;
        var tabArea = tabControl.GetTabRect(e.Index);

        if (e.Index == tabControl.SelectedIndex)
        {
            using (var brush = new SolidBrush(Color.FromArgb(0, 120, 215)))
            {
                g.FillRectangle(brush, tabArea);
            }
            TextRenderer.DrawText(g, tabControl.TabPages[e.Index].Text, tabControl.Font, tabArea, Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
        else
        {
            g.FillRectangle(Brushes.White, tabArea);
            TextRenderer.DrawText(g, tabControl.TabPages[e.Index].Text, tabControl.Font, tabArea, Color.Black, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
    }

    public static void ApplyStyle(DataGridView dgv)
    {
        // 1. DoubleBuffering chống giật lag
        typeof(DataGridView).InvokeMember(
            "DoubleBuffered",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty,
            null,
            dgv,
            new object[] { true });

        // 2. Tiêu đề cột chuẩn đẹp, chữ to rõ ràng
        dgv.EnableHeadersVisualStyles = false;
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(230, 236, 245);
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(0, 51, 102);
        dgv.ColumnHeadersDefaultCellStyle.Font = GetFont(10.5f, FontStyle.Bold);
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        dgv.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
        dgv.ColumnHeadersDefaultCellStyle.Padding = new Padding(2, 6, 2, 6);
        dgv.ColumnHeadersHeight = 44;
        dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;

        // 3. Định dạng dòng
        dgv.DefaultCellStyle.Font = GetFont(10f);
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
        dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
        dgv.RowHeadersVisible = false;
        dgv.AllowUserToResizeRows = true;
        dgv.RowTemplate.Resizable = DataGridViewTriState.True;
        dgv.RowTemplate.Height = 35;
        dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 254);
        dgv.GridColor = Color.FromArgb(220, 226, 235);
    }

    public static readonly System.Globalization.CultureInfo ViCulture = new System.Globalization.CultureInfo("vi-VN")
    {
        NumberFormat = {
            NumberGroupSeparator = ".",
            NumberDecimalSeparator = ","
        }
    };

    public static string FormatTien(decimal value) => value.ToString("#,##0", ViCulture);

    public static string FormatTyLe(decimal value) => value.ToString("0.####", ViCulture);

    public static decimal ParseFlexibleDecimal(string text, bool isPercentage = false)
    {
        if (string.IsNullOrWhiteSpace(text)) return 0m;
        string clean = text.Trim().Replace("%", "").Replace("đ", "").Replace("Đ", "").Replace(" ", "");
        if (string.IsNullOrWhiteSpace(clean)) return 0m;

        int dotCount = clean.Count(c => c == '.');
        int commaCount = clean.Count(c => c == ',');

        if (dotCount > 0 && commaCount > 0)
        {
            int lastDot = clean.LastIndexOf('.');
            int lastComma = clean.LastIndexOf(',');
            if (lastComma > lastDot)
            {
                // Chuẩn Việt Nam: 1.500.000,25 -> chấm là hàng nghìn, phẩy là thập phân
                clean = clean.Replace(".", "").Replace(",", ".");
            }
            else
            {
                // Chuẩn quốc tế: 1,500,000.25 -> phẩy là hàng nghìn, chấm là thập phân
                clean = clean.Replace(",", "");
            }
        }
        else if (commaCount > 0)
        {
            if (commaCount > 1)
            {
                // Nhiều dấu phẩy: 1,000,000 -> phẩy là hàng nghìn
                clean = clean.Replace(",", "");
            }
            else
            {
                // Một dấu phẩy: 1500,25 hoặc 1,5 -> phẩy là thập phân
                clean = clean.Replace(",", ".");
            }
        }
        else if (dotCount > 0)
        {
            if (dotCount > 1)
            {
                // Nhiều dấu chấm: 15.000.000 -> chấm là hàng nghìn
                clean = clean.Replace(".", "");
            }
            else
            {
                // Một dấu chấm duy nhất
                if (isPercentage)
                {
                    // Tỷ lệ %: 1.25 hoặc 0.8 -> chấm luôn là thập phân
                }
                else
                {
                    int dotIdx = clean.IndexOf('.');
                    int digitsAfterDot = clean.Length - dotIdx - 1;
                    // Nếu số chữ số sau dấu chấm khác 3 (ví dụ: 1500.25, 12.5): là số thập phân
                    // Nếu đúng 3 chữ số và phần đầu <= 3 chữ số (ví dụ: 1.500, 20.000): là phân cách hàng nghìn
                    if (digitsAfterDot == 3 && dotIdx <= 3)
                    {
                        clean = clean.Replace(".", "");
                    }
                }
            }
        }

        return decimal.TryParse(clean, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal val) ? val : 0m;
    }

    public static decimal ParseTien(string text) => ParseFlexibleDecimal(text, isPercentage: false);

    public static decimal ParseTyLe(string text) => ParseFlexibleDecimal(text, isPercentage: true);

    public static string ChuanHoaChuThuong(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return string.Empty;
        text = text.Trim();
        bool hasLetters = text.Any(char.IsLetter);
        bool isAllUpper = hasLetters && text.Where(char.IsLetter).All(char.IsUpper);
        if (isAllUpper)
        {
            var textInfo = new System.Globalization.CultureInfo("vi-VN", false).TextInfo;
            return textInfo.ToTitleCase(text.ToLower());
        }
        return text;
    }

    private static readonly string[] ChuSo = { "không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín" };

    public static string DocSoThanhChu(decimal number)
    {
        long n = (long)Math.Round(number, MidpointRounding.AwayFromZero);
        if (n == 0) return "Không đồng";

        bool isNegative = n < 0;
        if (isNegative) n = -n;

        string s = n.ToString();
        var groups = new System.Collections.Generic.List<int>();
        while (s.Length > 0)
        {
            int len = Math.Min(3, s.Length);
            string part = s.Substring(s.Length - len, len);
            groups.Add(int.Parse(part));
            s = s.Substring(0, s.Length - len);
        }

        var groupPhrases = new System.Collections.Generic.List<string>();
        for (int i = groups.Count - 1; i >= 0; i--)
        {
            int g = groups[i];
            if (g == 0) continue;

            bool dayDu = (i < groups.Count - 1);
            string t = DocBlock3So(g, dayDu);
            if (!string.IsNullOrEmpty(t))
            {
                string unit = "";
                if (i == 1) unit = "nghìn";
                else if (i == 2) unit = "triệu";
                else if (i == 3) unit = "tỷ";
                else if (i == 4) unit = "nghìn tỷ";
                else if (i == 5) unit = "triệu tỷ";
                else if (i == 6) unit = "tỷ tỷ";

                if (!string.IsNullOrEmpty(unit))
                    groupPhrases.Add($"{t} {unit}");
                else
                    groupPhrases.Add(t);
            }
        }

        string res = string.Join(", ", groupPhrases).Trim();
        if (res.Length > 0)
        {
            res = char.ToUpper(res[0]) + res.Substring(1);
        }
        if (isNegative) res = "Âm " + res;
        return res + " đồng";
    }

    private static string DocBlock3So(int n, bool dayDu)
    {
        int tram = n / 100;
        int chuc = (n % 100) / 10;
        int dv = n % 10;
        var res = new System.Collections.Generic.List<string>();

        if (tram > 0 || dayDu)
        {
            res.Add(ChuSo[tram] + " trăm");
        }

        if (chuc > 1)
        {
            res.Add(ChuSo[chuc] + " mươi");
            if (dv == 1) res.Add("mốt");
            else if (dv == 4) res.Add("tư");
            else if (dv == 5) res.Add("lăm");
            else if (dv > 0) res.Add(ChuSo[dv]);
        }
        else if (chuc == 1)
        {
            res.Add("mười");
            if (dv == 5) res.Add("lăm");
            else if (dv > 0) res.Add(ChuSo[dv]);
        }
        else if (chuc == 0 && dv > 0)
        {
            if (res.Count > 0) res.Add("lẻ");
            res.Add(ChuSo[dv]);
        }

        return string.Join(" ", res);
    }

    /// <summary>
    /// Gắn hành vi tự động chuyển đổi ký tự '.' trên bàn phím số Numpad thành dấu phân cách thập phân thực tế (',')
    /// giúp người dùng gõ số thực trên TextBox/Control luôn đúng chuẩn và không bị lỗi.
    /// </summary>
    public static void AttachDecimalInputBehavior(Control control)
    {
        if (control == null) return;
        control.KeyPress += (s, e) =>
        {
            if (e.KeyChar == '.' || e.KeyChar == ',')
            {
                string decSep = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
                if (!string.IsNullOrEmpty(decSep))
                {
                    e.KeyChar = decSep[0];
                }
            }
        };
    }
}
