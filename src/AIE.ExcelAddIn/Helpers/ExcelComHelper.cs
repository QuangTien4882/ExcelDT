using System;
using System.Runtime.InteropServices;

namespace AIE.ExcelAddIn.Helpers
{
    /// <summary>
    /// Tiện ích quản lý và giải phóng an toàn tài nguyên COM Interop của Excel.
    /// Giúp ngăn ngừa hiện tượng rò rỉ bộ nhớ và tiến trình EXCEL.EXE bị treo ngầm.
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
    }
}
