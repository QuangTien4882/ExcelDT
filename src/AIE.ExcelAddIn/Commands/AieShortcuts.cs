using System;
using System.Windows.Forms;
using ExcelDna.Integration;
using AIE.ExcelAddIn.Forms;
using AIE.ExcelAddIn.Ribbon;

namespace AIE.ExcelAddIn.Commands
{
    /// <summary>
    /// Hệ thống phím tắt tiện dụng (Keyboard Shortcuts) của AIE Add-In
    /// </summary>
    public static class AieShortcuts
    {
        /// <summary>
        /// Phím tắt Ctrl + Shift + S: Đồng bộ dữ liệu tức thì từ Sheet Excel vào mô hình Dự toán
        /// </summary>
        [ExcelCommand(ShortCut = "^+S", Description = "Đồng bộ tức thì từ Sheet Excel vào Dự toán")]
        public static void ShortcutDongBoExcel()
        {
            try
            {
                AieRibbon.DongBoDuLieuTuExcel();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi thực thi phím tắt Ctrl+Shift+S: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Phím tắt Ctrl + Shift + K: Mở bảng Tổng hợp kinh phí (Chi phí XD, Dự toán công trình & TMĐT)
        /// </summary>
        [ExcelCommand(ShortCut = "^+K", Description = "Mở bảng Tổng hợp kinh phí")]
        public static void ShortcutTongHopKinhPhi()
        {
            try
            {
                // Gọi handler mở Tổng hợp kinh phí
                AieRibbon.MoTongHopKinhPhiTuShortcut();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi thực thi phím tắt Ctrl+Shift+K: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Phím tắt Ctrl + Shift + F: Mở cửa sổ Tra cứu định mức nhanh
        /// </summary>
        [ExcelCommand(ShortCut = "^+F", Description = "Mở cửa sổ Tra cứu định mức nhanh")]
        public static void ShortcutTraCuuDinhMuc()
        {
            try
            {
                var form = new TraCuuDinhMucForm();
                form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi thực thi phím tắt Ctrl+Shift+F: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
