using AIE.Core.Models;
using Newtonsoft.Json;
using System;
using System.IO;

namespace AIE.ExcelAddIn.Services
{
    /// <summary>
    /// Dịch vụ lưu/đọc file dự toán (.dt).
    /// File .dt bản chất là JSON chứa toàn bộ đối tượng DuToan.
    /// </summary>
    public static class DuToanFileService
    {
        private static readonly JsonSerializerSettings _settings = new JsonSerializerSettings
        {
            Formatting = Formatting.Indented,
            NullValueHandling = NullValueHandling.Ignore,
            TypeNameHandling = TypeNameHandling.Auto, // Cần thiết để deserialize đúng kiểu kế thừa (VatLieuHienTruong, NhanCongHienTruong, MayThiCongHienTruong)
            ReferenceLoopHandling = ReferenceLoopHandling.Ignore
        };

        /// <summary>
        /// Lưu đối tượng DuToan ra file .dt (JSON).
        /// Tự động tạo bản lưu .bak trước khi ghi đè (luân chuyển tối đa <paramref name="maxBackups"/> bản).
        /// </summary>
        public static void Save(DuToan duToan, string filePath, int maxBackups = 5)
        {
            if (duToan == null) throw new ArgumentNullException(nameof(duToan));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("Đường dẫn file không hợp lệ.", nameof(filePath));

            if (File.Exists(filePath)) TaoBanLuuTruocKhiGhiDe(filePath, maxBackups);

            var json = JsonConvert.SerializeObject(duToan, _settings);
            File.WriteAllText(filePath, json);
        }

        /// <summary>
        /// Phục hồi file .dt từ bản lưu .bak gần nhất (nếu có).
        /// Trả về true nếu thành công.
        /// </summary>
        public static bool KhoiPhucTuBanLuu(string filePath)
        {
            string bak = filePath + ".bak";
            if (!File.Exists(bak)) return false;

            File.Copy(bak, filePath, true);
            return true;
        }

        /// <summary>
        /// Tạo bản lưu .bak trước khi ghi đè file .dt.
        /// Luân chuyển: .bak → .bak.1 → .bak.2 → … → .bak.(n-1). Giữ bản mới nhất tại .bak.
        /// </summary>
        private static void TaoBanLuuTruocKhiGhiDe(string filePath, int maxBackups)
        {
            try
            {
                // Luân chuyển các bản cũ: .bak.(n-1) → .bak.n
                for (int i = maxBackups - 1; i >= 1; i--)
                {
                    string src = filePath + (i == 1 ? ".bak" : $".bak.{i - 1}");
                    string dst = filePath + $".bak.{i}";
                    if (File.Exists(src))
                    {
                        if (File.Exists(dst)) File.Delete(dst);
                        File.Move(src, dst);
                    }
                }

                // Bản lưu mới nhất = .bak
                string newestBak = filePath + ".bak";
                if (File.Exists(newestBak)) File.Delete(newestBak);
                File.Copy(filePath, newestBak);
            }
            catch { }
        }

        /// <summary>
        /// Đọc file .dt và trả về đối tượng DuToan
        /// </summary>
        public static DuToan Load(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Không tìm thấy file dự toán.", filePath);

            var json = File.ReadAllText(filePath);
            var duToan = JsonConvert.DeserializeObject<DuToan>(json, _settings);

            if (duToan == null)
                throw new InvalidOperationException("Không thể đọc dữ liệu dự toán từ file.");

            return duToan;
        }
    }
}
