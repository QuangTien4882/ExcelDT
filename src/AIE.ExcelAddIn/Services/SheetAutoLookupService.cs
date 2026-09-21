using System;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using AIE.ExcelAddIn.Forms;

namespace AIE.ExcelAddIn.Services
{
    /// <summary>
    /// Dịch vụ hỗ trợ mở bảng tra cứu định mức khi người dùng nhấp đúp chuột vào cột "Mã hiệu" (cột B) của sheet DuToan.
    /// Không can thiệp vào việc nhập liệu bình thường của người dùng tại cột B.
    /// </summary>
    public static class SheetAutoLookupService
    {
        private static bool _isHooked = false;
        private static bool _isFormShowing = false;

        /// <summary>
        /// Cờ tạm dừng dịch vụ khi đang thực hiện các thao tác ghi dữ liệu hàng loạt (mở file, áp giá...).
        /// </summary>
        public static bool IsSuspended { get; set; } = false;

        public static void Register()
        {
            if (_isHooked) return;
            try
            {
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                if (app != null)
                {
                    app.SheetBeforeDoubleClick += OnAppSheetBeforeDoubleClick;
                    _isHooked = true;
                }
            }
            catch { }
        }

        public static void Unregister()
        {
            if (!_isHooked) return;
            try
            {
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                if (app != null)
                {
                    app.SheetBeforeDoubleClick -= OnAppSheetBeforeDoubleClick;
                    _isHooked = false;
                }
            }
            catch { }
        }

        private static void OnAppSheetBeforeDoubleClick(object sh, Range target, ref bool cancel)
        {
            if (IsSuspended || _isFormShowing) return;

            try
            {
                var ws = sh as Worksheet;
                if (ws == null || target == null) return;
                if (!ws.Name.StartsWith("DuToan", StringComparison.OrdinalIgnoreCase)) return;

                // Khi click đúp chuột vào ô tại cột B (Mã hiệu), từ dòng 6 trở đi
                if (target.Column == 2 && target.Row >= 6)
                {
                    string cText = ws.Cells[target.Row, 3]?.Value2?.ToString()?.Trim() ?? "";
                    if (!cText.StartsWith("TỔNG CỘNG", StringComparison.OrdinalIgnoreCase) &&
                        !cText.Equals("CỘNG", StringComparison.OrdinalIgnoreCase))
                    {
                        cancel = true; // Ngăn không cho Excel vào in-cell edit mode
                        int row = target.Row;
                        ExcelAsyncUtil.QueueAsMacro(() =>
                        {
                            ShowLookupDialog(ws, row);
                        });
                    }
                }
            }
            catch { }
        }

        private static void ShowLookupDialog(Worksheet ws, int row)
        {
            if (_isFormShowing) return;
            _isFormShowing = true;
            try
            {
                using (var form = new TraCuuDinhMucForm("", row))
                {
                    form.ShowDialog();
                }
            }
            catch { }
            finally
            {
                _isFormShowing = false;
            }
        }
    }
}
