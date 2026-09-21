using System;
using System.Text.RegularExpressions;

namespace AIE.Core.Services.Shared
{
    public static class TextHelper
    {
        /// <summary>
        /// Chuẩn hóa tên hạng mục cho tiêu đề bảng biểu: chỉ in hoa chữ cái đầu tiên,
        /// bảo toàn tiền tố La Mã (I., II., III., v.v.) nếu có.
        /// Ví dụ: "NỀN, MẶT ĐƯỜNG" -> "Nền, mặt đường"
        ///        "HỆ THỐNG THOÁT NƯỚC" -> "Hệ thống thoát nước"
        ///        "I. NỀN, MẶT ĐƯỜNG" -> "I. Nền, mặt đường"
        /// </summary>
        public static string FormatTenHangMucHeader(string? input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            input = input.Trim();

            var match = Regex.Match(input, @"^([IVXLCDM]+\.)\s*(.*)$", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                string roman = match.Groups[1].Value.ToUpper();
                string rest = match.Groups[2].Value.Trim();
                if (rest.Length > 0)
                {
                    rest = char.ToUpper(rest[0]) + rest.Substring(1).ToLower();
                    return $"{roman} {rest}";
                }
                return roman;
            }

            if (input.Length == 1) return input.ToUpper();
            return char.ToUpper(input[0]) + input.Substring(1).ToLower();
        }

        /// <summary>
        /// Làm sạch tên file, loại bỏ tất cả các ký tự cấm của Windows file system.
        /// </summary>
        public static string SanitizeFileName(string? name, string fallback = "DuToan_Moi")
        {
            if (string.IsNullOrWhiteSpace(name)) return fallback;
            name = name.Trim();
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            name = name.Replace(" ", "_");
            while (name.Contains("__")) name = name.Replace("__", "_");
            name = name.Trim('_');
            return string.IsNullOrEmpty(name) ? fallback : name;
        }
    }
}
