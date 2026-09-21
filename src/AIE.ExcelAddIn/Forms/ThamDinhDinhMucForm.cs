using AIE.Core.Enums;
using AIE.Core.Models;
using AIE.ExcelAddIn.Helpers;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace AIE.ExcelAddIn.Forms;

/// <summary>
/// Bước 2: Giao diện Thẩm tra Định mức & Khối lượng theo Thông tư 38/2026/TT-BXD.
/// Hiển thị đối chiếu trực quan song song giữa Dự toán tư vấn đệ trình và Định mức chuẩn.
/// </summary>
public class ThamDinhDinhMucForm : Form
{
    private readonly List<KetQuaCongTacThamDinh> _results;
    private readonly Vung _vung;
    private readonly ThamDinhConfig _config;
    private readonly int? _boDonGiaId;
    private readonly Action<List<KetQuaCongTacThamDinh>> _onTiepTuc;

    private DataGridView dgvCongTac;
    private DataGridView dgvHaoPhi;
    private Label lblDetailTitle;
    private Label lblStatTong;
    private Label lblStatKhop;
    private Label lblStatSai;
    private Label lblStatTamTinh;

    private readonly CultureInfo ViVn = new CultureInfo("vi-VN");

    public ThamDinhDinhMucForm(
        List<KetQuaCongTacThamDinh> results,
        Vung vung,
        ThamDinhConfig config,
        int? boDonGiaId,
        Action<List<KetQuaCongTacThamDinh>> onTiepTuc)
    {
        _results = results;
        _vung = vung;
        _config = config;
        _boDonGiaId = boDonGiaId;
        _onTiepTuc = onTiepTuc;

        InitializeComponent();
        FormStateHelper.Attach(this);
        LoadDataToGrids();
    }

    private void InitializeComponent()
    {
        this.Text = "Bước 2: Thẩm Tra Định Mức & Khối Lượng (Thông tư 38/2026/TT-BXD)";
        var workingArea = Screen.PrimaryScreen.WorkingArea;
        this.Size = new Size((int)(workingArea.Width * 0.95), (int)(workingArea.Height * 0.92));
        this.StartPosition = FormStartPosition.CenterScreen;
        this.MinimizeBox = true;
        this.MaximizeBox = true;
        this.ShowInTaskbar = true;
        this.Font = new Font("Be Vietnam Pro", 9.5f);

        // Header Summary Panel
        var topPanel = new Panel { Dock = DockStyle.Top, Height = 65, Padding = new Padding(12, 8, 12, 8), BackColor = Color.FromArgb(245, 247, 250) };
        var summaryLayout = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };

        lblStatTong = CreateStatBadge("Tổng số công tác: 0", Color.FromArgb(52, 73, 94), Color.White);
        lblStatKhop = CreateStatBadge("🟢 Khớp chuẩn TT38: 0", Color.FromArgb(40, 167, 69), Color.White);
        lblStatSai = CreateStatBadge("🔴 Sai lệch định mức: 0", Color.FromArgb(220, 53, 69), Color.White);
        lblStatTamTinh = CreateStatBadge("🟡 Tạm tính (TT): 0", Color.FromArgb(243, 156, 18), Color.White);

        summaryLayout.Controls.Add(lblStatTong);
        summaryLayout.Controls.Add(lblStatKhop);
        summaryLayout.Controls.Add(lblStatSai);
        summaryLayout.Controls.Add(lblStatTamTinh);
        topPanel.Controls.Add(summaryLayout);

        // Bottom Action Panel
        var bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 55, Padding = new Padding(12, 8, 12, 8), BackColor = Color.FromArgb(245, 247, 250) };
        var btnLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1 };
        btnLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40f));
        btnLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20f));
        btnLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40f));

        var btnDongBo = new Button
        {
            Text = "🔄 Đồng bộ tất cả theo Định mức chuẩn TT38",
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold),
            Margin = new Padding(0, 2, 10, 2)
        };
        btnDongBo.FlatAppearance.BorderSize = 0;
        btnDongBo.Click += BtnDongBo_Click;

        var btnDong = new Button
        {
            Text = "Đóng",
            Dock = DockStyle.Fill,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Be Vietnam Pro", 10f),
            Margin = new Padding(5, 2, 5, 2)
        };
        btnDong.FlatAppearance.BorderColor = Color.LightGray;
        btnDong.Click += (s, e) => this.Close();

        var btnTiepTuc = new Button
        {
            Text = "Tiếp tục sang Bước 3: Thẩm tra Đơn giá ➔",
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(0, 120, 215),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold),
            Margin = new Padding(10, 2, 0, 2)
        };
        btnTiepTuc.FlatAppearance.BorderSize = 0;
        btnTiepTuc.Click += BtnTiepTuc_Click;

        btnLayout.Controls.Add(btnDongBo, 0, 0);
        btnLayout.Controls.Add(btnDong, 1, 0);
        btnLayout.Controls.Add(btnTiepTuc, 2, 0);
        bottomPanel.Controls.Add(btnLayout);

        // Split Container
        var splitContainer = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Horizontal,
            SplitterDistance = (int)(this.Height * 0.45),
            SplitterWidth = 6
        };

        // Top Split: Master Cong Tac
        var pnlMaster = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12, 6, 12, 6) };
        var lblMasterTitle = new Label
        {
            Text = "📋 DANH SÁCH CÔNG TÁC TRONG DỰ TOÁN (Click chọn dòng để xem chi tiết đối chiếu hao phí bên dưới):",
            Dock = DockStyle.Top,
            Height = 26,
            Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59)
        };
        dgvCongTac = CreateGrid();
        dgvCongTac.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dgvCongTac.MultiSelect = false;
        SetupCongTacColumns();
        dgvCongTac.SelectionChanged += DgvCongTac_SelectionChanged;

        pnlMaster.Controls.Add(dgvCongTac);
        pnlMaster.Controls.Add(lblMasterTitle);
        splitContainer.Panel1.Controls.Add(pnlMaster);

        // Bottom Split: Detail Hao Phi
        var pnlDetail = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12, 6, 12, 6) };
        lblDetailTitle = new Label
        {
            Text = "🔍 CHI TIẾT ĐỐI CHIẾU HAO PHÍ ĐỊNH MỨC:",
            Dock = DockStyle.Top,
            Height = 26,
            Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 102, 204)
        };
        dgvHaoPhi = CreateGrid();
        dgvHaoPhi.ReadOnly = true;
        SetupHaoPhiColumns();

        pnlDetail.Controls.Add(dgvHaoPhi);
        pnlDetail.Controls.Add(lblDetailTitle);
        splitContainer.Panel2.Controls.Add(pnlDetail);

        this.Controls.Add(splitContainer);
        this.Controls.Add(topPanel);
        this.Controls.Add(bottomPanel);
    }

    private Label CreateStatBadge(string text, Color bg, Color fg)
    {
        return new Label
        {
            Text = text,
            AutoSize = true,
            BackColor = bg,
            ForeColor = fg,
            Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Bold),
            Padding = new Padding(10, 7, 10, 7),
            Margin = new Padding(0, 4, 12, 0)
        };
    }

    private DataGridView CreateGrid()
    {
        var dgv = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = false,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            BackgroundColor = Color.White,
            RowTemplate = { Height = 28 },
            Font = new Font("Be Vietnam Pro", 9f),
            EnableHeadersVisualStyles = false,
            ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single,
            RowHeadersVisible = false,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            ColumnHeadersHeight = 35
        };
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(235, 238, 242);
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(30, 41, 59);
        dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Be Vietnam Pro", 9f, FontStyle.Bold);
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        return dgv;
    }

    private void SetupCongTacColumns()
    {
        dgvCongTac.Columns.Add(new DataGridViewTextBoxColumn { Name = "STT", HeaderText = "STT", Width = 45, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        dgvCongTac.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaHieuDT", HeaderText = "Mã dự toán", Width = 95 });
        dgvCongTac.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaHieuChuan", HeaderText = "Mã chuẩn TT38", Width = 105 });
        dgvCongTac.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenCongTacDT", HeaderText = "Tên công việc dự toán", Width = 280 });
        dgvCongTac.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenCongTacChuan", HeaderText = "Tên chuẩn TT38", Width = 280 });
        dgvCongTac.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", Width = 60, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        dgvCongTac.Columns.Add(new DataGridViewTextBoxColumn { Name = "KhoiLuong", HeaderText = "Khối lượng", Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "N3", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvCongTac.Columns.Add(new DataGridViewTextBoxColumn { Name = "TrangThai", HeaderText = "Đánh giá định mức", Width = 150, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, Font = new Font(this.Font, FontStyle.Bold) } });
        dgvCongTac.Columns.Add(new DataGridViewTextBoxColumn { Name = "SoSaiLech", HeaderText = "Số lỗi", Width = 65, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
    }

    private void SetupHaoPhiColumns()
    {
        dgvHaoPhi.Columns.Add(new DataGridViewTextBoxColumn { Name = "Loai", HeaderText = "Loại", Width = 65, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, Font = new Font(this.Font, FontStyle.Bold) } });
        dgvHaoPhi.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaHP", HeaderText = "Mã hiệu HP", Width = 95 });
        dgvHaoPhi.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenHPDT", HeaderText = "Tên hao phí (Dự toán)", Width = 260 });
        dgvHaoPhi.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenHPChuan", HeaderText = "Tên hao phí (Chuẩn TT38)", Width = 260 });
        dgvHaoPhi.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", Width = 60, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        dgvHaoPhi.Columns.Add(new DataGridViewTextBoxColumn { Name = "DinhMucDT", HeaderText = "Hao phí DT", Width = 95, DefaultCellStyle = new DataGridViewCellStyle { Format = "0.####", Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvHaoPhi.Columns.Add(new DataGridViewTextBoxColumn { Name = "DinhMucChuan", HeaderText = "Định mức TT38", Width = 105, DefaultCellStyle = new DataGridViewCellStyle { Format = "0.####", Alignment = DataGridViewContentAlignment.MiddleRight, ForeColor = Color.FromArgb(0, 102, 204), Font = new Font(this.Font, FontStyle.Bold) } });
        dgvHaoPhi.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChenhLech", HeaderText = "Chênh lệch", Width = 95, DefaultCellStyle = new DataGridViewCellStyle { Format = "0.####", Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvHaoPhi.Columns.Add(new DataGridViewTextBoxColumn { Name = "DanhGia", HeaderText = "Đánh giá chi tiết", Width = 180 });
    }

    private void LoadDataToGrids()
    {
        dgvCongTac.Rows.Clear();

        int countKhop = 0;
        int countSai = 0;
        int countTamTinh = 0;

        for (int i = 0; i < _results.Count; i++)
        {
            var kq = _results[i];
            int rowIndex = dgvCongTac.Rows.Add();
            var row = dgvCongTac.Rows[rowIndex];

            row.Cells["STT"].Value = i + 1;
            row.Cells["MaHieuDT"].Value = kq.DuToan.MaHieu;
            row.Cells["MaHieuChuan"].Value = kq.DinhMucChuan?.MaHieu ?? "(Không có)";
            row.Cells["TenCongTacDT"].Value = kq.DuToan.TenCongTac;
            row.Cells["TenCongTacChuan"].Value = kq.DinhMucChuan?.TenCongTac ?? "(Công tác tạm tính / Báo giá riêng)";
            row.Cells["DonVi"].Value = kq.DuToan.DonVi;
            row.Cells["KhoiLuong"].Value = kq.DuToan.KhoiLuong;

            int soLoi = kq.DanhSachSaiLech.Count(s => !string.IsNullOrEmpty(s.LoaiLoi));
            row.Cells["SoSaiLech"].Value = soLoi;

            if (kq.DinhMucChuan == null)
            {
                countTamTinh++;
                row.Cells["TrangThai"].Value = "🟡 Tạm tính (TT)";
                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 252, 235);
                row.Cells["TrangThai"].Style.ForeColor = Color.FromArgb(180, 115, 0);
            }
            else if (soLoi > 0)
            {
                countSai++;
                row.Cells["TrangThai"].Value = $"🔴 Sai {soLoi} hao phí";
                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 240, 240);
                row.Cells["TrangThai"].Style.ForeColor = Color.FromArgb(200, 30, 30);
            }
            else
            {
                countKhop++;
                row.Cells["TrangThai"].Value = "🟢 Khớp chuẩn TT38";
                row.Cells["TrangThai"].Style.ForeColor = Color.FromArgb(40, 167, 69);
            }

            row.Tag = kq;
        }

        lblStatTong.Text = $"Tổng số công tác: {_results.Count}";
        lblStatKhop.Text = $"🟢 Khớp chuẩn TT38: {countKhop}";
        lblStatSai.Text = $"🔴 Sai lệch định mức: {countSai}";
        lblStatTamTinh.Text = $"🟡 Tạm tính (TT): {countTamTinh}";

        if (dgvCongTac.Rows.Count > 0)
        {
            dgvCongTac.Rows[0].Selected = true;
            HienThiChiTietHaoPhi(_results[0]);
        }
    }

    private void DgvCongTac_SelectionChanged(object? sender, EventArgs e)
    {
        if (dgvCongTac.SelectedRows.Count == 0) return;
        var selectedRow = dgvCongTac.SelectedRows[0];
        if (selectedRow.Tag is KetQuaCongTacThamDinh kq)
        {
            HienThiChiTietHaoPhi(kq);
        }
    }

    private void HienThiChiTietHaoPhi(KetQuaCongTacThamDinh kq)
    {
        lblDetailTitle.Text = $"🔍 CHI TIẾT ĐỐI CHIẾU HAO PHÍ ĐỊNH MỨC: [{kq.DuToan.MaHieu}] {kq.DuToan.TenCongTac}";
        dgvHaoPhi.Rows.Clear();

        if (kq.DanhSachSaiLech.Count == 0 && kq.DinhMucChuan != null)
        {
            // Hiển thị danh sách chuẩn nếu không có sai lệch
            foreach (var hp in kq.DinhMucChuan.DanhSachHaoPhi)
            {
                int r = dgvHaoPhi.Rows.Add();
                var row = dgvHaoPhi.Rows[r];
                row.Cells["Loai"].Value = hp.LoaiHaoPhi.ToString();
                row.Cells["MaHP"].Value = hp.MaHieuHP;
                row.Cells["TenHPDT"].Value = hp.TenHaoPhi;
                row.Cells["TenHPChuan"].Value = hp.TenHaoPhi;
                row.Cells["DonVi"].Value = hp.DonVi;
                row.Cells["DinhMucDT"].Value = hp.DinhMuc;
                row.Cells["DinhMucChuan"].Value = hp.DinhMuc;
                row.Cells["ChenhLech"].Value = 0m;
                row.Cells["DanhGia"].Value = "🟢 Trùng khớp";
                row.Cells["DanhGia"].Style.ForeColor = Color.FromArgb(40, 167, 69);
            }
            return;
        }

        foreach (var sl in kq.DanhSachSaiLech)
        {
            int r = dgvHaoPhi.Rows.Add();
            var row = dgvHaoPhi.Rows[r];

            string loaiStr = sl.LoaiHP?.ToString() ?? (sl.HaoPhiChuan?.LoaiHaoPhi.ToString() ?? "VL");
            row.Cells["Loai"].Value = loaiStr;
            row.Cells["MaHP"].Value = sl.HaoPhiChuan?.MaHieuHP ?? sl.HaoPhiDuToan?.MaHieuHP ?? "";
            row.Cells["TenHPDT"].Value = sl.HaoPhiDuToan?.TenHaoPhi ?? "(Không có trong dự toán)";
            row.Cells["TenHPChuan"].Value = sl.HaoPhiChuan?.TenHaoPhi ?? "(Không có trong TT38)";
            row.Cells["DonVi"].Value = sl.HaoPhiChuan?.DonVi ?? sl.HaoPhiDuToan?.DonVi ?? "";

            decimal dmDt = sl.HaoPhiDuToan?.DinhMuc ?? 0;
            decimal dmChuan = sl.HaoPhiChuan?.DinhMuc ?? 0;
            row.Cells["DinhMucDT"].Value = dmDt;
            row.Cells["DinhMucChuan"].Value = dmChuan;
            row.Cells["ChenhLech"].Value = sl.ChenhLechDinhMuc;

            if (string.IsNullOrEmpty(sl.LoaiLoi))
            {
                row.Cells["DanhGia"].Value = "🟢 Trùng khớp";
                row.Cells["DanhGia"].Style.ForeColor = Color.FromArgb(40, 167, 69);
            }
            else
            {
                row.Cells["DanhGia"].Value = "🔴 " + sl.LoaiLoi;
                row.Cells["DanhGia"].Style.ForeColor = Color.FromArgb(220, 53, 69);
                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 245, 245);
                if (Math.Abs(sl.ChenhLechDinhMuc) > 0.0001m)
                {
                    row.Cells["ChenhLech"].Style.ForeColor = Color.Red;
                    row.Cells["ChenhLech"].Style.Font = new Font(dgvHaoPhi.Font, FontStyle.Bold);
                }
            }
        }
    }

    private void BtnDongBo_Click(object? sender, EventArgs e)
    {
        int dongBoCount = 0;
        foreach (var kq in _results)
        {
            if (kq.DinhMucChuan != null && kq.DanhSachSaiLech.Any(s => !string.IsNullOrEmpty(s.LoaiLoi)))
            {
                // Đồng bộ hao phí dự toán sang chuẩn
                kq.DuToan.DanhSachHaoPhi.Clear();
                foreach (var hpChuan in kq.DinhMucChuan.DanhSachHaoPhi)
                {
                    kq.DuToan.DanhSachHaoPhi.Add(new HaoPhiThamDinh
                    {
                        Loai = hpChuan.LoaiHaoPhi,
                        MaHieuHP = hpChuan.MaHieuHP,
                        TenHaoPhi = hpChuan.TenHaoPhi,
                        DonVi = hpChuan.DonVi,
                        DinhMuc = hpChuan.DinhMuc,
                        DonGia = 0,
                        SoDongExcel = kq.DuToan.SoDongExcel
                    });
                }
                kq.DanhSachSaiLech.Clear();
                dongBoCount++;
            }
        }

        if (dongBoCount > 0)
        {
            MessageBox.Show($"Đã đồng bộ {dongBoCount} công tác về định mức chuẩn Thông tư 38/2026/TT-BXD thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            LoadDataToGrids();
        }
        else
        {
            MessageBox.Show("Tất cả công tác chuẩn đã trùng khớp hoặc là công tác tạm tính, không cần đồng bộ.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void BtnTiepTuc_Click(object? sender, EventArgs e)
    {
        this.DialogResult = DialogResult.OK;
        this.Close();
        _onTiepTuc?.Invoke(_results);
    }
}
