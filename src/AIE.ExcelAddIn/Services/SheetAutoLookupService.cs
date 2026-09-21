using System;
using System.Linq;
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
    /// Dịch vụ lắng nghe sự kiện nhập liệu trực tiếp tại cột "Mã hiệu" (cột B) của sheet DuToan:
    /// - Nếu gõ đúng mã hiệu chuẩn: tự động điền Tên, ĐVT, công thức Thành tiền và Auto-focus ô Khối lượng.
    /// - Nếu gõ tiền tố / từ khóa: tự động hiển thị bảng Tra cứu định mức tương ứng để người dùng lựa chọn nhanh.
    /// </summary>
    public static class SheetAutoLookupService
    {
        private static bool _isHooked = false;
        private static bool _isProcessing = false;

        public static void Register()
        {
            if (_isHooked) return;
            try
            {
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                if (app != null)
                {
                    app.SheetChange += OnAppSheetChange;
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
                    app.SheetChange -= OnAppSheetChange;
                    _isHooked = false;
                }
            }
            catch { }
        }

        private static void OnAppSheetChange(object sh, Range target)
        {
            if (_isProcessing) return;

            try
            {
                var ws = sh as Worksheet;
                if (ws == null || target == null) return;

                // Chỉ bắt trên các sheet Dự toán (DuToan, DuToan_...)
                string sheetName = ws.Name;
                if (!sheetName.StartsWith("DuToan", StringComparison.OrdinalIgnoreCase)) return;

                // Chỉ bắt khi thay đổi đúng 1 ô tại Cột B (Cột 2 - Mã hiệu) từ dòng 6 trở đi
                if (target.Columns.Count > 1 || target.Rows.Count > 1) return;
                if (target.Column != 2 || target.Row < 6) return;

                string val = target.Value2?.ToString()?.Trim() ?? "";
                if (string.IsNullOrWhiteSpace(val)) return;

                // Bỏ qua nếu dòng này là dòng TỔNG CỘNG
                string cText = ws.Cells[target.Row, 3]?.Value2?.ToString()?.Trim() ?? "";
                if (cText.StartsWith("TỔNG CỘNG", StringComparison.OrdinalIgnoreCase) || cText.Equals("CỘNG", StringComparison.OrdinalIgnoreCase)) return;

                // Chạy macro bất đồng bộ để tránh COM busy
                int row = target.Row;
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    HandleLookup(ws, row, val);
                });
            }
            catch { }
        }

        private static void HandleLookup(Worksheet ws, int row, string keyword)
        {
            if (_isProcessing) return;
            _isProcessing = true;
            try
            {
                var db = new DatabaseManager();
                var repo = new CongTacRepository(db.Context);

                // Kiểm tra xem có khớp chính xác 100% 1 mã hiệu không (ví dụ AA.11124)
                var exactMatch = repo.GetByMaHieu(keyword.ToUpperInvariant());
                if (exactMatch != null)
                {
                    ws.Cells[row, 2].Value2 = exactMatch.MaHieu;
                    ws.Cells[row, 3].Value2 = exactMatch.TenCongTac;
                    ws.Cells[row, 4].Value2 = exactMatch.DonVi;

                    // Công thức Thành tiền
                    ws.Cells[row, 9].Formula = $"=ROUND(E{row}*F{row}, 0)";
                    ws.Cells[row, 10].Formula = $"=ROUND(E{row}*G{row}, 0)";
                    ws.Cells[row, 11].Formula = $"=ROUND(E{row}*H{row}, 0)";

                    // Định dạng số và đường viền
                    ExcelFormatHelper.ApplyQuantityFormat(ws.Cells[row, 5], 2);
                    ExcelFormatHelper.ApplyIntegerFormat(ws.Range[ws.Cells[row, 6], ws.Cells[row, 11]]);
                    var rowRange = ws.Range[ws.Cells[row, 1], ws.Cells[row, 11]];
                    rowRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                    rowRange.VerticalAlignment = XlVAlign.xlVAlignCenter;

                    // Đánh lại STT tự động cho toàn bảng
                    LapDuToanExcelService.DanhLaiSTTCongTac(ws);

                    // Auto-Focus vào ô Khối lượng
                    ((Range)ws.Cells[row, 5]).Select();
                    return;
                }

                // Nếu không khớp chính xác 100% (ví dụ gõ "AA", "AF", "bê tông"...):
                // Hiển thị form Tra cứu định mức với từ khóa keyword đã được nạp sẵn
                using var form = new TraCuuDinhMucForm(keyword, row);
                form.ShowDialog();
            }
            catch { }
            finally
            {
                _isProcessing = false;
            }
        }
    }
}
