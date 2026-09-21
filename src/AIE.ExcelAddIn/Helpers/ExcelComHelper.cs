using System;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace AIE.ExcelAddIn.Helpers
{
    /// <summary>
    /// Tiện ích quản lý và giải phóng an toàn tài nguyên COM Interop của Excel.
    /// Giúp ngăn ngừa hiện tượng rò rỉ bộ nhớ và tiến trình EXCEL.EXE bị treo ngầm.
    /// Cung cấp giải pháp đọc/ghi theo mảng khối (Bulk Range) tăng tốc độ gấp 50-100 lần.
    /// </summary>
    public static class ExcelComHelper
    {
        /// <summary>
        /// Giải phóng an toàn một hoặc nhiều đối tượng COM (Worksheet, Range, Workbook, v.v.)
        /// </summary>
        public static void SafeRelease(params object?[]? comObjects)
        {
            if (comObjects == null) return;

            foreach (var obj in comObjects)
            {
                if (obj != null && Marshal.IsComObject(obj))
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(obj);
                    }
                    catch
                    {
                        // Bỏ qua lỗi giải phóng nếu đối tượng đã bị giải phóng trước đó
                    }
                }
            }
        }

        /// <summary>
        /// Thu gom rác bộ nhớ và dọn sạch các tham chiếu COM còn sót lại.
        /// Thường gọi sau các tác vụ xuất bảng biểu quy mô lớn hoặc khi đóng form.
        /// </summary>
        public static void CollectGarbage()
        {
            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch { }
        }

        /// <summary>
        /// Đọc toàn bộ giá trị trong một vùng ô Excel vào mảng 2 chiều trong RAM (1 lần gọi COM duy nhất).
        /// </summary>
        public static object[,] ReadRangeData(Worksheet ws, int startRow, int startCol, int endRow, int endCol)
        {
            if (ws == null || startRow > endRow || startCol > endCol) return null;
            Range rng = null;
            try
            {
                rng = ws.Range[ws.Cells[startRow, startCol], ws.Cells[endRow, endCol]];
                object val = rng.Value2;
                if (val is object[,] arr) return arr;
                if (val != null)
                {
                    object[,] single = new object[2, 2];
                    single[1, 1] = val;
                    return single;
                }
                return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                SafeRelease(rng);
            }
        }

        /// <summary>
        /// Ghi đồng loạt mảng giá trị 2 chiều xuống vùng ô Excel trong 1 lần gọi COM duy nhất.
        /// </summary>
        public static void WriteRangeData(Worksheet ws, int startRow, int startCol, object[,] data)
        {
            if (ws == null || data == null) return;
            int numRows = data.GetLength(0);
            int numCols = data.GetLength(1);
            if (numRows <= 0 || numCols <= 0) return;

            Range rng = null;
            try
            {
                rng = ws.Range[ws.Cells[startRow, startCol], ws.Cells[startRow + numRows - 1, startCol + numCols - 1]];
                rng.Value2 = data;
            }
            catch { }
            finally
            {
                SafeRelease(rng);
            }
        }

        /// <summary>
        /// Ghi đồng loạt mảng công thức 2 chiều xuống vùng ô Excel trong 1 lần gọi COM duy nhất.
        /// </summary>
        public static void WriteRangeFormulas(Worksheet ws, int startRow, int startCol, object[,] formulas)
        {
            if (ws == null || formulas == null) return;
            int numRows = formulas.GetLength(0);
            int numCols = formulas.GetLength(1);
            if (numRows <= 0 || numCols <= 0) return;

            Range rng = null;
            try
            {
                rng = ws.Range[ws.Cells[startRow, startCol], ws.Cells[startRow + numRows - 1, startCol + numCols - 1]];
                rng.Formula = formulas;
            }
            catch { }
            finally
            {
                SafeRelease(rng);
            }
        }
    }

    /// <summary>
    /// Trình đọc dữ liệu Excel siêu tốc (Bulk Range Reader) trong bộ nhớ RAM.
    /// Nạp toàn bộ dữ liệu của sheet hoặc vùng chọn vào mảng RAM chỉ trong 1 lần gọi COM,
    /// sau đó tất cả các thao tác truy xuất ô đều diễn ra tức thì trong RAM (tốc độ vi giây).
    /// </summary>
    public class FastRangeReader
    {
        private readonly object[,] _data;
        private readonly int _startRow;
        private readonly int _startCol;
        private readonly int _rows;
        private readonly int _cols;
        private readonly int _lower0;
        private readonly int _lower1;

        public FastRangeReader(Worksheet ws, int startRow, int startCol, int endRow, int endCol)
        {
            _startRow = startRow;
            _startCol = startCol;
            if (ws != null && startRow <= endRow && startCol <= endCol)
            {
                Range rng = null;
                try
                {
                    rng = ws.Range[ws.Cells[startRow, startCol], ws.Cells[endRow, endCol]];
                    object val = rng.Value2;
                    if (val is object[,] arr)
                    {
                        _data = arr;
                        _rows = arr.GetLength(0);
                        _cols = arr.GetLength(1);
                        _lower0 = arr.GetLowerBound(0);
                        _lower1 = arr.GetLowerBound(1);
                    }
                    else if (val != null)
                    {
                        _data = new object[2, 2];
                        _data[1, 1] = val;
                        _rows = 1;
                        _cols = 1;
                        _lower0 = 1;
                        _lower1 = 1;
                    }
                }
                catch { }
                finally
                {
                    ExcelComHelper.SafeRelease(rng);
                }
            }
        }

        public bool HasData => _data != null;
        public int StartRow => _startRow;
        public int StartCol => _startCol;
        public int RowCount => _rows;
        public int ColCount => _cols;

        public object GetValue(int row, int col)
        {
            if (_data == null) return null;
            int rIdx = row - _startRow + _lower0;
            int cIdx = col - _startCol + _lower1;
            if (rIdx >= _lower0 && rIdx < _lower0 + _rows && cIdx >= _lower1 && cIdx < _lower1 + _cols)
            {
                return _data[rIdx, cIdx];
            }
            return null;
        }

        public string GetString(int row, int col)
        {
            object val = GetValue(row, col);
            return val?.ToString()?.Trim() ?? string.Empty;
        }

        public decimal GetDecimal(int row, int col)
        {
            object val = GetValue(row, col);
            if (val == null) return 0m;
            if (val is double d) return (decimal)d;
            if (val is decimal m) return m;
            if (val is int i) return i;
            if (val is long l) return l;
            string s = val.ToString().Trim();
            if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal res))
                return res;
            if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out res))
                return res;
            return 0m;
        }

        public double GetDouble(int row, int col)
        {
            object val = GetValue(row, col);
            if (val == null) return 0.0;
            if (val is double d) return d;
            if (val is decimal m) return (double)m;
            if (val is int i) return i;
            if (val is long l) return l;
            string s = val.ToString().Trim();
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double res))
                return res;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out res))
                return res;
            return 0.0;
        }
    }
}
