using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using AIE.Data;
using AIE.Data.Repositories;
using AIE.ExcelAddIn.Forms;
using AIE.ExcelAddIn.Helpers;

namespace AIE.ExcelAddIn.Services
{
    /// <summary>
    /// Dịch vụ tra cứu định mức siêu tốc tại cột "Mã hiệu" (cột B) của sheet DuToan:
    /// 1. Bắt phím tức thì khi gõ ký tự đầu tiên: Ngay khi người dùng nhấn phím ký tự bất kỳ (A-Z, 0-9, F2...),
    ///    bảng Tra cứu định mức hiển thị ngay tức thì mà không cần bấm Enter!
    /// 2. Bắt khi người dùng gõ từ khóa rồi Enter hoặc dán dữ liệu vào cột B.
    /// 3. Double-Click: Nhấp đúp chuột vào ô cột B mở ngay bảng Tra cứu định mức.
    /// </summary>
    public static class SheetAutoLookupService
    {
        private static bool _isHooked = false;
        private static bool _isFormShowing = false;

        // Windows Keyboard Hook definitions
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_SYSKEYDOWN = 0x0104;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private static readonly LowLevelKeyboardProc _hookProc = HookCallback;
        private static IntPtr _hookId = IntPtr.Zero;

        private static Worksheet _activeWorksheet = null;
        private static int _activeRow = -1;

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr LoadLibrary(string lpFileName);

        public static void Register()
        {
            if (_isHooked) return;
            try
            {
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                if (app != null)
                {
                    app.SheetSelectionChange += OnAppSheetSelectionChange;
                    app.SheetBeforeDoubleClick += OnAppSheetBeforeDoubleClick;
                    app.SheetChange += OnAppSheetChange;
                    _isHooked = true;
                }
            }
            catch { }
        }

        public static void Unregister()
        {
            UninstallKeyboardHook();
            if (!_isHooked) return;
            try
            {
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                if (app != null)
                {
                    app.SheetSelectionChange -= OnAppSheetSelectionChange;
                    app.SheetBeforeDoubleClick -= OnAppSheetBeforeDoubleClick;
                    app.SheetChange -= OnAppSheetChange;
                    _isHooked = false;
                }
            }
            catch { }
        }

        private static void InstallKeyboardHook(Worksheet ws, int row)
        {
            if (_hookId != IntPtr.Zero) return;
            try
            {
                _activeWorksheet = ws;
                _activeRow = row;
                IntPtr hMod = LoadLibrary("user32.dll");
                _hookId = SetWindowsHookEx(WH_KEYBOARD_LL, _hookProc, hMod, 0);
            }
            catch { }
        }

        private static void UninstallKeyboardHook()
        {
            if (_hookId == IntPtr.Zero) return;
            try
            {
                UnhookWindowsHookEx(_hookId);
            }
            catch { }
            finally
            {
                _hookId = IntPtr.Zero;
                _activeWorksheet = null;
                _activeRow = -1;
            }
        }

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
            {
                int vkCode = Marshal.ReadInt32(lParam);
                Keys key = (Keys)vkCode;

                // Nếu nhấn F2: mở ngay bảng định mức
                if (key == Keys.F2)
                {
                    var targetWs = _activeWorksheet;
                    int targetRow = _activeRow;
                    UninstallKeyboardHook();
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        ShowLookupDialog(targetWs, targetRow, "");
                    });
                    return (IntPtr)1;
                }

                if (!IsIgnoredKey(key))
                {
                    string charStr = GetCharFromKey(key);
                    if (!string.IsNullOrEmpty(charStr))
                    {
                        var targetWs = _activeWorksheet;
                        int targetRow = _activeRow;

                        // Gỡ hook ngay để tránh bắt lặp phím
                        UninstallKeyboardHook();

                        // Mở ngay form Tra cứu định mức với ký tự đầu tiên vừa gõ
                        ExcelAsyncUtil.QueueAsMacro(() =>
                        {
                            ShowLookupDialog(targetWs, targetRow, charStr);
                        });

                        // Chặn phím gửi vào Excel để Excel không rơi vào in-cell edit mode
                        return (IntPtr)1;
                    }
                }
            }
            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        private static bool IsIgnoredKey(Keys key)
        {
            switch (key)
            {
                case Keys.Escape:
                case Keys.Tab:
                case Keys.Enter:
                case Keys.Up:
                case Keys.Down:
                case Keys.Left:
                case Keys.Right:
                case Keys.Home:
                case Keys.End:
                case Keys.PageUp:
                case Keys.PageDown:
                case Keys.Insert:
                case Keys.Delete:
                case Keys.Back:
                case Keys.ControlKey:
                case Keys.ShiftKey:
                case Keys.Menu:
                case Keys.LWin:
                case Keys.RWin:
                case Keys.Apps:
                case Keys.Capital:
                case Keys.NumLock:
                case Keys.Scroll:
                case Keys.PrintScreen:
                case Keys.Pause:
                    return true;
            }

            if (key >= Keys.F1 && key <= Keys.F24) return true;

            // Bỏ qua các tổ hợp phím tắt như Ctrl+C, Ctrl+V, Ctrl+Z, Alt+...
            if ((Control.ModifierKeys & Keys.Control) != 0 || (Control.ModifierKeys & Keys.Alt) != 0)
            {
                return true;
            }

            return false;
        }

        private static string GetCharFromKey(Keys key)
        {
            if (key >= Keys.A && key <= Keys.Z)
            {
                bool isShift = (Control.ModifierKeys & Keys.Shift) != 0;
                bool isCaps = Console.CapsLock;
                bool isUpper = isShift ^ isCaps;
                char c = (char)('a' + (key - Keys.A));
                return isUpper ? char.ToUpper(c).ToString() : c.ToString();
            }
            if (key >= Keys.D0 && key <= Keys.D9)
            {
                return ((char)('0' + (key - Keys.D0))).ToString();
            }
            if (key >= Keys.NumPad0 && key <= Keys.NumPad9)
            {
                return ((char)('0' + (key - Keys.NumPad0))).ToString();
            }
            if (key == Keys.OemPeriod || key == Keys.Decimal)
            {
                return ".";
            }
            if (key == Keys.OemMinus || key == Keys.Subtract)
            {
                return "-";
            }
            return null;
        }

        private static void OnAppSheetSelectionChange(object sh, Range target)
        {
            try
            {
                var ws = sh as Worksheet;
                if (ws == null || target == null)
                {
                    UninstallKeyboardHook();
                    return;
                }

                string sheetName = ws.Name;
                if (sheetName.StartsWith("DuToan", StringComparison.OrdinalIgnoreCase) &&
                    target.Columns.Count == 1 && target.Rows.Count == 1 &&
                    target.Column == 2 && target.Row >= 6)
                {
                    string cText = ws.Cells[target.Row, 3]?.Value2?.ToString()?.Trim() ?? "";
                    if (!cText.StartsWith("TỔNG CỘNG", StringComparison.OrdinalIgnoreCase) &&
                        !cText.Equals("CỘNG", StringComparison.OrdinalIgnoreCase))
                    {
                        InstallKeyboardHook(ws, target.Row);
                        return;
                    }
                }

                UninstallKeyboardHook();
            }
            catch
            {
                UninstallKeyboardHook();
            }
        }

        private static void OnAppSheetBeforeDoubleClick(object sh, Range target, ref bool cancel)
        {
            try
            {
                var ws = sh as Worksheet;
                if (ws == null || target == null) return;
                if (!ws.Name.StartsWith("DuToan", StringComparison.OrdinalIgnoreCase)) return;

                if (target.Column == 2 && target.Row >= 6)
                {
                    string cText = ws.Cells[target.Row, 3]?.Value2?.ToString()?.Trim() ?? "";
                    if (!cText.StartsWith("TỔNG CỘNG", StringComparison.OrdinalIgnoreCase))
                    {
                        cancel = true; // Ngăn không cho Excel vào cell-edit mode
                        int row = target.Row;
                        UninstallKeyboardHook();
                        ExcelAsyncUtil.QueueAsMacro(() =>
                        {
                            ShowLookupDialog(ws, row, "");
                        });
                    }
                }
            }
            catch { }
        }

        private static void OnAppSheetChange(object sh, Range target)
        {
            if (_isFormShowing) return;

            try
            {
                var ws = sh as Worksheet;
                if (ws == null || target == null) return;

                string sheetName = ws.Name;
                if (!sheetName.StartsWith("DuToan", StringComparison.OrdinalIgnoreCase)) return;

                if (target.Columns.Count > 1 || target.Rows.Count > 1) return;
                if (target.Column != 2 || target.Row < 6) return;

                string val = target.Value2?.ToString()?.Trim() ?? "";
                if (string.IsNullOrWhiteSpace(val)) return;

                string cText = ws.Cells[target.Row, 3]?.Value2?.ToString()?.Trim() ?? "";
                if (cText.StartsWith("TỔNG CỘNG", StringComparison.OrdinalIgnoreCase) || cText.Equals("CỘNG", StringComparison.OrdinalIgnoreCase)) return;

                int row = target.Row;
                UninstallKeyboardHook();
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    HandleExactOrLookup(ws, row, val);
                });
            }
            catch { }
        }

        private static void HandleExactOrLookup(Worksheet ws, int row, string keyword)
        {
            if (_isFormShowing) return;
            try
            {
                var db = new DatabaseManager();
                var repo = new CongTacRepository(db.Context);

                var exactMatch = repo.GetByMaHieu(keyword.ToUpperInvariant());
                if (exactMatch != null)
                {
                    ws.Cells[row, 2].Value2 = exactMatch.MaHieu;
                    ws.Cells[row, 3].Value2 = exactMatch.TenCongTac;
                    ws.Cells[row, 4].Value2 = exactMatch.DonVi;

                    ws.Cells[row, 9].Formula = $"=ROUND(E{row}*F{row}, 0)";
                    ws.Cells[row, 10].Formula = $"=ROUND(E{row}*G{row}, 0)";
                    ws.Cells[row, 11].Formula = $"=ROUND(E{row}*H{row}, 0)";

                    ExcelFormatHelper.ApplyQuantityFormat(ws.Cells[row, 5], 2);
                    ExcelFormatHelper.ApplyIntegerFormat(ws.Range[ws.Cells[row, 6], ws.Cells[row, 11]]);
                    var rowRange = ws.Range[ws.Cells[row, 1], ws.Cells[row, 11]];
                    rowRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                    rowRange.VerticalAlignment = XlVAlign.xlVAlignCenter;

                    LapDuToanExcelService.DanhLaiSTTCongTac(ws);

                    ((Range)ws.Cells[row, 5]).Select();
                    return;
                }

                // Không khớp 100% -> mở form Tra cứu định mức với từ khóa
                ShowLookupDialog(ws, row, keyword);
            }
            catch { }
            finally
            {
                RecheckHook(ws);
            }
        }

        private static void ShowLookupDialog(Worksheet ws, int row, string initialKeyword)
        {
            if (_isFormShowing) return;
            _isFormShowing = true;
            try
            {
                using (var form = new TraCuuDinhMucForm(initialKeyword, row))
                {
                    form.ShowDialog();
                }
            }
            catch { }
            finally
            {
                _isFormShowing = false;
                RecheckHook(ws);
            }
        }

        private static void RecheckHook(Worksheet ws)
        {
            try
            {
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                var curWs = app.ActiveSheet as Worksheet;
                var curCell = app.ActiveCell;
                if (curWs != null && curCell != null &&
                    curWs.Name.StartsWith("DuToan", StringComparison.OrdinalIgnoreCase) &&
                    curCell.Column == 2 && curCell.Row >= 6)
                {
                    string cText = curWs.Cells[curCell.Row, 3]?.Value2?.ToString()?.Trim() ?? "";
                    if (!cText.StartsWith("TỔNG CỘNG", StringComparison.OrdinalIgnoreCase) &&
                        !cText.Equals("CỘNG", StringComparison.OrdinalIgnoreCase))
                    {
                        InstallKeyboardHook(curWs, curCell.Row);
                    }
                }
            }
            catch { }
        }
    }
}
