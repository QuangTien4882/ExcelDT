using Xunit;
using AIE.Core.Services.Shared;

namespace AIE.Core.Tests
{
    public class TextHelperTests
    {
        [Theory]
        [InlineData("NỀN, MẶT ĐƯỜNG", "Nền, mặt đường")]
        [InlineData("HỆ THỐNG THOÁT NƯỚC", "Hệ thống thoát nước")]
        [InlineData("I. NỀN, MẶT ĐƯỜNG", "I. Nền, mặt đường")]
        [InlineData("II. HỆ THỐNG THOÁT NƯỚC", "II. Hệ thống thoát nước")]
        [InlineData("III. XÂY DỰNG NHÀ ĐIỀU HÀNH", "III. Xây dựng nhà điều hành")]
        [InlineData("Hạng mục 1", "Hạng mục 1")]
        [InlineData("cầu bê tông", "Cầu bê tông")]
        [InlineData("", "")]
        [InlineData(null, "")]
        public void FormatTenHangMucHeader_DungQuyTacSentenceCase(string? input, string expected)
        {
            string actual = TextHelper.FormatTenHangMucHeader(input);
            Assert.Equal(expected, actual);
        }

        [Theory]
        [InlineData("CongTrinh/2026:KhuA", "CongTrinh_2026_KhuA")]
        [InlineData("Công trình * Cầu ? Nhà | Trạm", "Công_trình_Cầu_Nhà_Trạm")]
        [InlineData("   ", "DuToan")]
        [InlineData(null, "DuToan")]
        [InlineData("DuToanChuan", "DuToanChuan")]
        public void SanitizeFileName_LoaiBoKyTuCamVaDungFallback(string? input, string expected)
        {
            string actual = TextHelper.SanitizeFileName(input, "DuToan");
            Assert.Equal(expected, actual);
        }
    }
}
