using ExcelDna.Integration;

namespace AIE.ExcelAddIn;

/// <summary>
/// Lớp khởi tạo Add-in khi Excel tải.
/// </summary>
public class AieAddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        try
        {
            System.Windows.Forms.Application.SetUnhandledExceptionMode(System.Windows.Forms.UnhandledExceptionMode.CatchException);
            System.Windows.Forms.Application.ThreadException += (s, e) =>
            {
                LogCrash("ThreadException", e.Exception);
            };
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                LogCrash("UnhandledException", e.ExceptionObject as System.Exception);
            };
        }
        catch { }

        // Khởi tạo DB và seed dữ liệu nhân công nếu chưa có
        try
        {
            var db = new AIE.Data.DatabaseManager();
            db.SeedNhanCong();
        }
        catch { /* Bỏ qua lỗi seed */ }

        // Khởi động trình tự động sao lưu (Auto-Save 5 phút) và kiểm tra khôi phục crash
        try
        {
            AIE.ExcelAddIn.Services.AutoSaveManager.Start(5);
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                AIE.ExcelAddIn.Services.AutoSaveManager.CheckAndPromptRecovery();
            });
        }
        catch { }

        // Đăng ký dịch vụ tự động tra cứu định mức khi người dùng nhập liệu cột Mã hiệu trên sheet DuToan
        try
        {
            AIE.ExcelAddIn.Services.SheetAutoLookupService.Register();
        }
        catch { }
    }

    private static void LogCrash(string source, System.Exception? ex)
    {
        if (ex == null) return;
        try
        {
            string dir = System.IO.Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "AIE_DuToan");
            if (!System.IO.Directory.Exists(dir)) System.IO.Directory.CreateDirectory(dir);
            string file = System.IO.Path.Combine(dir, "crash_log.txt");
            string entry = $"[{System.DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{source}] {ex.GetType().FullName}: {ex.Message}\r\n{ex.StackTrace}\r\n\r\n";
            System.IO.File.AppendAllText(file, entry);
        }
        catch { }
    }

    public void AutoClose()
    {
        try
        {
            AIE.ExcelAddIn.Services.AutoSaveManager.Stop();
        }
        catch { }

        try
        {
            AIE.ExcelAddIn.Services.SheetAutoLookupService.Unregister();
        }
        catch { }
    }
}
