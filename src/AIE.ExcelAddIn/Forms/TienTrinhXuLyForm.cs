using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using AIE.ExcelAddIn.Helpers;

namespace AIE.ExcelAddIn.Forms
{
    /// <summary>
    /// Form hiển thị thanh tiến trình trực quan khi thực hiện các tác vụ nặng:
    /// Tra cứu dữ liệu lớn, tính toán thẩm định, xuất toàn bộ bảng biểu Excel.
    /// </summary>
    public class TienTrinhXuLyForm : Form
    {
        private readonly Label lblTieuDe;
        private readonly Label lblTrangThai;
        private readonly ProgressBar progressBar;
        private readonly Button btnHuy;
        private readonly CancellationTokenSource? _cts;

        public bool IsCancelled => _cts?.IsCancellationRequested ?? false;

        public TienTrinhXuLyForm(string tieuDe, bool choPhepChiHuy = false, CancellationTokenSource? cts = null)
        {
            _cts = cts;

            this.Text = "AIE - Đang xử lý dữ liệu";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = true;
            this.TopMost = true;
            this.Size = new Size(520, 220);
            this.BackColor = Color.FromArgb(248, 249, 250);
            this.Font = UIHelper.GetFont(9.5f);

            // 1. Tiêu đề
            lblTieuDe = new Label
            {
                Text = tieuDe,
                Location = new Point(25, 20),
                Size = new Size(460, 28),
                Font = UIHelper.GetFont(11f, FontStyle.Bold),
                ForeColor = Color.FromArgb(16, 124, 65)
            };
            this.Controls.Add(lblTieuDe);

            // 2. Nhãn trạng thái chi tiết
            lblTrangThai = new Label
            {
                Text = "Đang khởi tạo tác vụ...",
                Location = new Point(25, 55),
                Size = new Size(460, 24),
                Font = UIHelper.GetFont(9.5f),
                ForeColor = Color.FromArgb(70, 70, 70)
            };
            this.Controls.Add(lblTrangThai);

            // 3. ProgressBar
            progressBar = new ProgressBar
            {
                Location = new Point(25, 88),
                Size = new Size(455, 26),
                Minimum = 0,
                Maximum = 100,
                Value = 0,
                Style = ProgressBarStyle.Continuous
            };
            this.Controls.Add(progressBar);

            // 4. Nút Hủy
            btnHuy = new Button
            {
                Text = "Hủy bỏ",
                Location = new Point(380, 130),
                Size = new Size(100, 34),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(9.5f),
                Cursor = Cursors.Hand,
                Visible = choPhepChiHuy
            };
            btnHuy.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
            btnHuy.Click += (s, e) =>
            {
                if (_cts != null && !_cts.IsCancellationRequested)
                {
                    _cts.Cancel();
                    lblTrangThai.Text = "Đang yêu cầu dừng tác vụ...";
                    btnHuy.Enabled = false;
                }
            };
            this.Controls.Add(btnHuy);
        }

        /// <summary>
        /// Cập nhật % tiến trình và nhãn trạng thái từ thread nền một cách an toàn (Thread-safe).
        /// </summary>
        public void CapNhatTienTrinh(int phanTram, string trangThai)
        {
            if (this.InvokeRequired)
            {
                try
                {
                    this.BeginInvoke(new Action(() => CapNhatTienTrinh(phanTram, trangThai)));
                }
                catch { }
                return;
            }

            if (phanTram < 0)
            {
                progressBar.Style = ProgressBarStyle.Marquee;
            }
            else
            {
                progressBar.Style = ProgressBarStyle.Continuous;
                progressBar.Value = Math.Max(0, Math.Min(100, phanTram));
            }

            if (!string.IsNullOrEmpty(trangThai))
            {
                lblTrangThai.Text = trangThai;
            }
        }

        /// <summary>
        /// Phương thức tiện ích để thực thi tác vụ nền với hộp thoại hiển thị tiến trình.
        /// </summary>
        /// <param name="tieuDe">Tiêu đề tác vụ hiển thị</param>
        /// <param name="taskFunc">Hàm thực thi (nhận delegate báo tiến trình: (phanTram, thongBao))</param>
        public static void ChayTacVu(string tieuDe, Action<Action<int, string>> taskFunc)
        {
            var form = new TienTrinhXuLyForm(tieuDe);
            form.Shown += async (s, e) =>
            {
                try
                {
                    await Task.Run(() =>
                    {
                        taskFunc((pct, msg) => form.CapNhatTienTrinh(pct, msg));
                    });
                }
                catch (Exception ex)
                {
                    form.Invoke(new Action(() =>
                    {
                        MessageBox.Show($"Lỗi trong quá trình xử lý: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
                finally
                {
                    form.Invoke(new Action(() => form.Close()));
                }
            };

            form.ShowDialog();
        }
    }
}
