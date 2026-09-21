using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using AIE.Core.Models;
using AIE.ExcelAddIn.Ribbon;

namespace AIE.ExcelAddIn.Services
{
    /// <summary>
    /// Quản lý tự động lưu (Auto-Save) và phục hồi dự toán khi gặp sự cố crash.
    /// Định kỳ lưu bản sao lưu (.autosave.dt) vào %AppData%\AIE_DuToan\AutoSave.
    /// </summary>
    public static class AutoSaveManager
    {
        private static System.Windows.Forms.Timer _timer;
        private static readonly object _lock = new object();

        /// <summary>Thư mục lưu các bản tự động sao lưu</summary>
        public static string AutoSaveDirectory
        {
            get
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "AIE_DuToan",
                    "AutoSave");
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                return dir;
            }
        }

        /// <summary>
        /// Khởi động tiến trình tự động lưu ngầm.
        /// </summary>
        /// <param name="intervalMinutes">Chu kỳ tự động lưu (mặc định 5 phút)</param>
        public static void Start(int intervalMinutes = 5)
        {
            lock (_lock)
            {
                if (_timer != null)
                {
                    _timer.Stop();
                    _timer.Dispose();
                    _timer = null;
                }

                _timer = new System.Windows.Forms.Timer();
                _timer.Interval = Math.Max(1, intervalMinutes) * 60 * 1000;
                _timer.Tick += (s, e) => ExecuteAutoSave();
                _timer.Start();
            }
        }

        /// <summary>
        /// Dừng tiến trình tự động lưu.
        /// </summary>
        public static void Stop()
        {
            lock (_lock)
            {
                if (_timer != null)
                {
                    _timer.Stop();
                    _timer.Dispose();
                    _timer = null;
                }
            }
        }

        /// <summary>
        /// Thực hiện tự động lưu dự toán hiện tại.
        /// </summary>
        public static void ExecuteAutoSave()
        {
            try
            {
                var dt = AieRibbon.CurrentDuToan;
                if (dt == null) return;

                // Chỉ lưu nếu dự toán có dữ liệu
                bool hasData = (dt.DanhSachHangMuc != null && dt.DanhSachHangMuc.Count > 0) ||
                               !string.IsNullOrEmpty(dt.TenCongTrinh) ||
                               !string.IsNullOrEmpty(dt.TenDuAn);
                if (!hasData) return;

                string wbKey = AieRibbon.GetActiveWorkbookKey();
                string safeName = SanitizeFileName(wbKey);
                if (string.IsNullOrEmpty(safeName) || safeName == "__default__")
                {
                    safeName = !string.IsNullOrEmpty(dt.TenCongTrinh)
                        ? SanitizeFileName(dt.TenCongTrinh)
                        : "DuToan_ChuaDatTen";
                }

                string filePath = Path.Combine(AutoSaveDirectory, $"{safeName}.autosave.dt");
                string tempPath = filePath + ".tmp";

                // Lưu ra file tạm trước để tránh file bị hỏng nếu crash đúng lúc đang ghi
                DuToanFileService.Save(dt, tempPath, maxBackups: 1);

                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
                File.Move(tempPath, filePath);
            }
            catch
            {
                // Bỏ qua lỗi trong quá trình tự động lưu ngầm để không làm phiền người dùng
            }
        }

        /// <summary>
        /// Xóa bản sao lưu tự động khi người dùng đã bấm lưu chính thức hoặc đóng file bình thường.
        /// </summary>
        public static void DeleteAutoSave(string wbKey)
        {
            try
            {
                if (string.IsNullOrEmpty(wbKey)) return;
                string safeName = SanitizeFileName(wbKey);
                string filePath = Path.Combine(AutoSaveDirectory, $"{safeName}.autosave.dt");
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
            catch { }
        }

        /// <summary>
        /// Kiểm tra xem có bản sao lưu tự động nào chưa được lưu chính thức không.
        /// Nếu có, hiển thị hộp thoại hỏi người dùng có muốn khôi phục không.
        /// </summary>
        public static void CheckAndPromptRecovery()
        {
            try
            {
                if (!Directory.Exists(AutoSaveDirectory)) return;

                var files = Directory.GetFiles(AutoSaveDirectory, "*.autosave.dt")
                    .Select(f => new FileInfo(f))
                    .OrderByDescending(f => f.LastWriteTime)
                    .ToList();

                if (files.Count == 0) return;

                foreach (var fi in files)
                {
                    // Nếu bản sao lưu quá cũ (hơn 14 ngày), dọn dẹp
                    if ((DateTime.Now - fi.LastWriteTime).TotalDays > 14)
                    {
                        try { fi.Delete(); } catch { }
                        continue;
                    }

                    string tenHienThi = Path.GetFileNameWithoutExtension(Path.GetFileNameWithoutExtension(fi.Name));
                    var result = MessageBox.Show(
                        $"AIE Dự Toán phát hiện bản tự động sao lưu (Auto-Save) từ phiên làm việc trước:\n\n" +
                        $"• Dự án/File: {tenHienThi}\n" +
                        $"• Thời điểm lưu: {fi.LastWriteTime:dd/MM/yyyy HH:mm:ss}\n\n" +
                        $"Bạn có muốn khôi phục lại bản dự toán này không?",
                        "Phục Hồi Dự Toán (Crash Recovery)",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        // Mở file dự toán
                        AieRibbon.MoDuToanTuDuongDan(fi.FullName);
                        // Đổi tên hoặc xóa file sau khi đã nạp
                        try { fi.Delete(); } catch { }
                        break;
                    }
                    else
                    {
                        // Người dùng từ chối khôi phục -> xóa file sao lưu để không hỏi lại
                        try { fi.Delete(); } catch { }
                    }
                }
            }
            catch { }
        }

        private static string SanitizeFileName(string name)
        {
            return AIE.Core.Services.Shared.TextHelper.SanitizeFileName(name, "Unknown");
        }
    }
}
