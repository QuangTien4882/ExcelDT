using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using AIE.Core.Models;
using AIE.Core.Services.LapDuToan;
using AIE.Data;
using AIE.Data.ImportExport;
using AIE.Data.Repositories;
using AIE.ExcelAddIn.Forms;
using AIE.ExcelAddIn.Helpers;
using AIE.ExcelAddIn.Services;
using Dapper;

namespace AIE.ExcelAddIn.Ribbon
{
    public class WindowWrapper : System.Windows.Forms.IWin32Window
    {
        public WindowWrapper(IntPtr handle)
        {
            Handle = handle;
        }
        public IntPtr Handle { get; }
    }

    [ComVisible(true)]
    public class AieRibbon : ExcelRibbon
    {
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, AIE.Core.Models.DuToan> _duToanByWb = new System.Collections.Concurrent.ConcurrentDictionary<string, AIE.Core.Models.DuToan>(StringComparer.OrdinalIgnoreCase);
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, string> _filePathByWb = new System.Collections.Concurrent.ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, TongHopKinhPhiForm> _tongHopForms = new System.Collections.Concurrent.ConcurrentDictionary<string, TongHopKinhPhiForm>(StringComparer.OrdinalIgnoreCase);
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, TinhGiaHienTruongForm> _tinhGiaForms = new System.Collections.Concurrent.ConcurrentDictionary<string, TinhGiaHienTruongForm>(StringComparer.OrdinalIgnoreCase);
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, ThamDinhDonGiaForm> _thamDinhForms = new System.Collections.Concurrent.ConcurrentDictionary<string, ThamDinhDonGiaForm>(StringComparer.OrdinalIgnoreCase);
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, XacNhanSaiKhacForm> _xacNhanSaiKhacForms = new System.Collections.Concurrent.ConcurrentDictionary<string, XacNhanSaiKhacForm>(StringComparer.OrdinalIgnoreCase);
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, System.Collections.Generic.List<KetQuaCongTacThamDinh>> _lastThamDinhResultsByWb = new System.Collections.Concurrent.ConcurrentDictionary<string, System.Collections.Generic.List<KetQuaCongTacThamDinh>>(StringComparer.OrdinalIgnoreCase);
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, ThamDinhConfig> _lastThamDinhConfigByWb = new System.Collections.Concurrent.ConcurrentDictionary<string, ThamDinhConfig>(StringComparer.OrdinalIgnoreCase);
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, int?> _lastBoDonGiaIdByWb = new System.Collections.Concurrent.ConcurrentDictionary<string, int?>(StringComparer.OrdinalIgnoreCase);

        public static string GetActiveWorkbookKey()
        {
            try
            {
                var app = (Application)ExcelDnaUtil.Application;
                var wb = app?.ActiveWorkbook;
                if (wb != null && !string.IsNullOrEmpty(wb.Name))
                {
                    return wb.Name;
                }
            }
            catch { }
            return "__default__";
        }

        /// <summary>Dự toán đang làm việc (in-memory) theo Workbook hiện tại.</summary>
        public static AIE.Core.Models.DuToan CurrentDuToan
        {
            get
            {
                string key = GetActiveWorkbookKey();
                _duToanByWb.TryGetValue(key, out var dt);
                return dt;
            }
            set
            {
                string key = GetActiveWorkbookKey();
                if (value == null)
                    _duToanByWb.TryRemove(key, out _);
                else
                    _duToanByWb[key] = value;
            }
        }

        /// <summary>Đường dẫn file .dt hiện tại theo Workbook hiện tại (null nếu chưa lưu)</summary>
        public static string CurrentFilePath
        {
            get
            {
                string key = GetActiveWorkbookKey();
                _filePathByWb.TryGetValue(key, out var path);
                return path;
            }
            set
            {
                string key = GetActiveWorkbookKey();
                if (value == null)
                    _filePathByWb.TryRemove(key, out _);
                else
                    _filePathByWb[key] = value;
            }
        }

        // Singleton references theo từng Workbook: tránh mở nhiều instance cùng lúc cho cùng một Workbook
        private static TongHopKinhPhiForm _tongHopForm
        {
            get { _tongHopForms.TryGetValue(GetActiveWorkbookKey(), out var f); return f; }
            set { string k = GetActiveWorkbookKey(); if (value == null) _tongHopForms.TryRemove(k, out _); else _tongHopForms[k] = value; }
        }

        private static TinhGiaHienTruongForm _tinhGiaForm
        {
            get { _tinhGiaForms.TryGetValue(GetActiveWorkbookKey(), out var f); return f; }
            set { string k = GetActiveWorkbookKey(); if (value == null) _tinhGiaForms.TryRemove(k, out _); else _tinhGiaForms[k] = value; }
        }

        private static ThamDinhDonGiaForm _thamDinhForm
        {
            get { _thamDinhForms.TryGetValue(GetActiveWorkbookKey(), out var f); return f; }
            set { string k = GetActiveWorkbookKey(); if (value == null) _thamDinhForms.TryRemove(k, out _); else _thamDinhForms[k] = value; }
        }

        private static XacNhanSaiKhacForm _xacNhanSaiKhacForm
        {
            get { _xacNhanSaiKhacForms.TryGetValue(GetActiveWorkbookKey(), out var f); return f; }
            set { string k = GetActiveWorkbookKey(); if (value == null) _xacNhanSaiKhacForms.TryRemove(k, out _); else _xacNhanSaiKhacForms[k] = value; }
        }

        private static System.Collections.Generic.List<KetQuaCongTacThamDinh> _lastThamDinhResults
        {
            get { _lastThamDinhResultsByWb.TryGetValue(GetActiveWorkbookKey(), out var r); return r; }
            set { string k = GetActiveWorkbookKey(); if (value == null) _lastThamDinhResultsByWb.TryRemove(k, out _); else _lastThamDinhResultsByWb[k] = value; }
        }

        private static ThamDinhConfig _lastThamDinhConfig
        {
            get { _lastThamDinhConfigByWb.TryGetValue(GetActiveWorkbookKey(), out var c); return c; }
            set { string k = GetActiveWorkbookKey(); if (value == null) _lastThamDinhConfigByWb.TryRemove(k, out _); else _lastThamDinhConfigByWb[k] = value; }
        }

        private static int? _lastBoDonGiaId
        {
            get { _lastBoDonGiaIdByWb.TryGetValue(GetActiveWorkbookKey(), out var id); return id; }
            set { string k = GetActiveWorkbookKey(); if (!value.HasValue) _lastBoDonGiaIdByWb.TryRemove(k, out _); else _lastBoDonGiaIdByWb[k] = value; }
        }
        public override string GetCustomUI(string RibbonID)
        {
            return @"
            <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
              <ribbon>
                <tabs>
                  <tab id='tabAIE' label='AIE Dự Toán'>
                    
                    <group id='groupThietLap' label='Thiết lập dữ liệu'>
                      <button id='btnImport' label='Nhập Database' screentip='Nhập Database' size='normal' showImage='false' onAction='OnImportClicked' />
                      <button id='btnTraCuu' label='Tra cứu định mức' screentip='Tra cứu định mức' size='normal' showImage='false' onAction='OnTraCuuClicked' />
                      <button id='btnTaoTemplate' label='Tạo File mẫu' screentip='Tạo File mẫu' size='normal' showImage='false' onAction='OnTaoTemplateClicked' />
                      <button id='btnXuatDonGia' label='Trích xuất Đơn Giá' screentip='Trích xuất Đơn Giá' size='normal' showImage='false' onAction='OnXuatDonGiaClicked' />
                      <button id='btnQuanLyDonGia' label='Quản lý Đơn giá' screentip='Quản lý Đơn giá' size='normal' showImage='false' onAction='OnQuanLyDonGiaClicked' />
                      <button id='btnDeleteDb' label='Xóa Database' screentip='Xóa toàn bộ Database' size='normal' showImage='false' onAction='OnDeleteDbClicked' />
                    </group>

                    <group id='groupThamDinh' label='Thẩm định dự toán'>
                      <button id='btnDonGiaThamDinh' label='Kiểm tra ĐM, ĐG' screentip='Kiểm tra Định mức, Đơn giá dự toán' size='normal' showImage='false' onAction='OnDonGiaThamDinhClicked' />
                      <button id='btnMoDonGiaThamDinh' label='Đơn giá TĐ' screentip='Mở Bộ Đơn giá Thẩm định' size='normal' showImage='false' onAction='OnMoDonGiaThamDinhClicked' />
                      <button id='btnKiemTra' label='Kiểm tra' screentip='Kiểm tra dự toán' size='normal' showImage='false' onAction='OnKiemTraClicked' />
                      <button id='btnBaoCaoTD' label='Kết quả thẩm định' screentip='Kết quả thẩm định dự toán (Bảng 3.8 TT 36 &amp; THDT TT 38)' size='normal' showImage='false' onAction='OnBaoCaoTDClicked' />
                    </group>

                    <group id='groupLapDuToan' label='Lập dự toán'>
                      <button id='btnTaoDuToanMoi' label='Tạo Dự toán mới' screentip='Tạo Dự toán mới' size='normal' showImage='false' onAction='OnTaoDuToanMoiClicked' />
                      <button id='btnDongBoExcel' label='Đồng bộ từ Excel' screentip='Đồng bộ khối lượng, công tác từ bảng tính Excel vào mô hình dự toán (Phím tắt: Ctrl+Shift+S)' size='normal' showImage='false' onAction='OnDongBoExcelClicked' />
                      <button id='btnGoiDonGia' label='Gọi Đơn giá' screentip='Gọi Đơn giá' size='normal' showImage='false' onAction='OnGoiDonGiaClicked' />
                      <button id='btnTinhGiaHienTruong' label='Giá VL, NC, MTC' screentip='Giá VL, NC, MTC' size='normal' showImage='false' onAction='OnTinhGiaHienTruongClicked' />
                      <button id='btnTongHopKinhPhi' label='Tổng hợp kinh phí' screentip='Tổng hợp Chi phí xây dựng, Dự toán &amp; Tổng mức đầu tư (Phím tắt: Ctrl+Shift+K)' size='normal' showImage='false' onAction='OnTongHopKinhPhiClicked' />
                    </group>

                    <group id='groupFile' label='File Dự toán'>
                      <button id='btnLuuDuToan' label='Lưu Dự toán' screentip='Lưu Dự toán (.dt)' size='normal' showImage='false' onAction='OnLuuDuToanClicked' />
                      <button id='btnMoDuToan' label='Mở Dự toán' screentip='Mở Dự toán (.dt)' size='normal' showImage='false' onAction='OnMoDuToanClicked' />
                      <button id='btnKhoiPhuc' label='Khôi phục lưu gần nhất' screentip='Khôi phục .dt từ bản lưu .bak gần nhất' size='normal' showImage='false' onAction='OnKhoiPhucBamLuuClicked' />
                      <button id='btnPhucHoiAutoSave' label='Phục hồi tự động' screentip='Kiểm tra và phục hồi từ bản sao lưu tự động (Auto-Save)' size='normal' showImage='false' onAction='OnPhucHoiAutoSaveClicked' />
                    </group>

                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        public void OnImportClicked(IRibbonControl control)
        {
            try
            {
                var form = new ImportDatabaseForm();
                form.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi mở form Import: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnXuatDonGiaClicked(IRibbonControl control)
        {
            try
            {
                using var fbd = new SaveFileDialog();
                fbd.Filter = "Excel Files|*.xlsx";
                fbd.Title = "Lưu file Đơn giá chuẩn";
                fbd.FileName = "DonGiaChuan_Template.xlsx";
                
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    var db = new DatabaseManager();
                    var repo = new AIE.Data.Repositories.CongTacRepository(db.Context);
                    var haoPhis = repo.GetAllHaoPhi().ToList();

                    if (haoPhis.Count == 0)
                    {
                        MessageBox.Show("Không có dữ liệu Hao phí trong hệ thống. Vui lòng nạp Định mức trước.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    var generator = new ExcelTemplateGenerator();
                    generator.ExportDonGiaTemplate(fbd.FileName, haoPhis);
                    
                    MessageBox.Show("Đã trích xuất danh mục Đơn giá thành công!\nBạn hãy mở file lên, điền giá trị và dùng nó cho các lần thẩm định sau.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xuất Đơn giá: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnTraCuuClicked(IRibbonControl control)
        {
            try
            {
                var form = new TraCuuDinhMucForm();
                // Show floating without owner to allow Alt+Tab and independent minimize
                form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi mở form Tra cứu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnTaoTemplateClicked(IRibbonControl control)
        {
            try
            {
                var dialog = new FolderBrowserDialog();
                dialog.Description = "Chọn thư mục để lưu các file mẫu (Templates):";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var generator = new ExcelTemplateGenerator();
                    generator.GenerateAllTemplates(dialog.SelectedPath);
                    MessageBox.Show(
                        "Đã tạo 6 file mẫu thành công tại:\n" + dialog.SelectedPath,
                        "AIE Dự Toán",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi Tạo File mẫu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnDeleteDbClicked(IRibbonControl control)
        {
            var form = new Forms.DeleteDatabaseForm();
            form.ShowDialog();
        }

        public void OnKiemTraClicked(IRibbonControl control)
        {
            try
            {
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                var wb = app.ActiveWorkbook;
                if (wb == null)
                {
                    MessageBox.Show("Không có file Excel nào đang mở.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var sheetNames = new System.Collections.Generic.List<string>();
                foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in wb.Worksheets)
                {
                    sheetNames.Add(sheet.Name);
                }

                var form = new ThamDinhConfigForm(sheetNames, _lastBoDonGiaId);
                if (form.ShowDialog() == DialogResult.OK)
                {
                    var config = form.ResultConfig;
                    
                    using var progress = new AIE.ExcelAddIn.Forms.TienTrinhXuLyForm("Thẩm Định Dự Toán Chi Tiết");
                    progress.Show();
                    progress.CapNhatTienTrinh(10, "Đang đọc dữ liệu công tác từ sheet Excel...");
                    System.Windows.Forms.Application.DoEvents();

                    // 1. Đọc dữ liệu công tác từ sheet
                    var reader = new AIE.ExcelAddIn.Services.DuToanExcelReader();
                    var danhSachCongTac = reader.Read(config);

                    if (danhSachCongTac.Count == 0)
                    {
                        progress.Close();
                        MessageBox.Show("Không tìm thấy dữ liệu công tác nào. Vui lòng kiểm tra lại cấu hình cột.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    progress.CapNhatTienTrinh(40, $"Đã đọc {danhSachCongTac.Count} công tác. Đang thẩm tra theo định mức...");
                    System.Windows.Forms.Application.DoEvents();

                    // 2. Thẩm tra theo Định mức chuẩn và Bộ đơn giá đã chọn (nếu có)
                    var db = new DatabaseManager();
                    var repo = new AIE.Data.Repositories.CongTacRepository(db.Context);
                    var vlRepo = new AIE.Data.Repositories.VatLieuRepository(db.Context);
                    var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);
                    var mayRepo = new AIE.Data.Repositories.MayThiCongRepository(db.Context);
                    var engine = new AIE.ExcelAddIn.Services.ThamDinhEngine(repo, vlRepo, ncRepo, mayRepo);
                    
                    var ketQua = engine.KiemTra(danhSachCongTac, form.SelectedBoDonGiaId, config.Vung);

                    progress.CapNhatTienTrinh(75, "Đang xuất kết quả thẩm định ra sheet KQ_ThamDinh...");
                    System.Windows.Forms.Application.DoEvents();

                    // 3. Xuất kết quả kiểm tra trực tiếp vào sheet KQ_ThamDinh
                    var writer = new AIE.ExcelAddIn.Services.ThamDinhExcelWriter();
                    writer.ExportResult(config, ketQua, form.SelectedBoDonGiaId);

                    progress.CapNhatTienTrinh(100, "Hoàn tất thẩm định dự toán!");
                    System.Windows.Forms.Application.DoEvents();
                    progress.Close();

                    // Lưu trạng thái phiên làm việc cho Xuất Báo cáo
                    _lastThamDinhConfig = config;
                    _lastBoDonGiaId = form.SelectedBoDonGiaId;
                    _lastThamDinhResults = ketQua;

                    int soLoi = ketQua?.Sum(x => x.DanhSachSaiLech?.Count(s => !string.IsNullOrEmpty(s.LoaiLoi)) ?? 0) ?? 0;
                    MessageBox.Show(
                        $"Thẩm định hoàn tất!\n\n" +
                        $"• Đã kiểm tra: {ketQua.Count} công tác\n" +
                        $"• Phát hiện: {soLoi} sai lệch\n" +
                        $"• Kết quả chi tiết đã được xuất ra sheet 'KQ_ThamDinh'.\n\n" +
                        $"Bạn có thể nhấn nút 'Xuất Báo cáo' để tạo Hồ sơ Báo cáo Thẩm định đầy đủ theo Thông tư 36/2026/TT-BXD.",
                        "AIE Dự Toán - Thẩm Định Hoàn Tất",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi kiểm tra thẩm định: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnDonGiaThamDinhClicked(IRibbonControl control)
        {
            try
            {
                var app = (Application)ExcelDnaUtil.Application;
                var wb = app.ActiveWorkbook;
                if (wb == null)
                {
                    MessageBox.Show("Không có file Excel nào đang mở.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (_thamDinhForm != null && !_thamDinhForm.IsDisposed)
                {
                    _thamDinhForm.WindowState = FormWindowState.Normal;
                    _thamDinhForm.Activate();
                    return;
                }

                var sheetNames = new System.Collections.Generic.List<string>();
                foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in wb.Worksheets)
                {
                    sheetNames.Add(sheet.Name);
                }

                var form = new ThamDinhConfigForm(sheetNames, _lastBoDonGiaId);
                form.Text = "Cấu Hình Đọc Dự Toán - Lấy Danh Mục Đơn Giá Thẩm Định";
                if (form.ShowDialog() == DialogResult.OK)
                {
                    var config = form.ResultConfig;
                    
                    using var progress = new AIE.ExcelAddIn.Forms.TienTrinhXuLyForm("Trích Xuất Đơn Giá Thẩm Định");
                    progress.Show();
                    progress.CapNhatTienTrinh(15, "Đang đọc dữ liệu công tác từ sheet Excel...");
                    System.Windows.Forms.Application.DoEvents();

                    var reader = new AIE.ExcelAddIn.Services.DuToanExcelReader();
                    var danhSachCongTac = reader.Read(config);

                    if (danhSachCongTac.Count == 0)
                    {
                        progress.Close();
                        MessageBox.Show("Không tìm thấy dữ liệu công tác nào. Vui lòng kiểm tra lại cấu hình cột.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    progress.CapNhatTienTrinh(50, $"Đã đọc {danhSachCongTac.Count} công tác. Đang phân tích hao phí định mức...");
                    System.Windows.Forms.Application.DoEvents();

                    var db = new DatabaseManager();
                    var repo = new AIE.Data.Repositories.CongTacRepository(db.Context);
                    var vlRepo = new AIE.Data.Repositories.VatLieuRepository(db.Context);
                    var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);
                    var mayRepo = new AIE.Data.Repositories.MayThiCongRepository(db.Context);
                    var engine = new AIE.ExcelAddIn.Services.ThamDinhEngine(repo, vlRepo, ncRepo, mayRepo);
                    
                    var ketQua = engine.KiemTra(danhSachCongTac, form.SelectedBoDonGiaId, config.Vung);

                    progress.CapNhatTienTrinh(80, "Đang tổng hợp danh mục vật tư VL, NC, Máy...");
                    System.Windows.Forms.Application.DoEvents();

                    var danhSachVatTu = engine.TrichXuatVatTu(ketQua);

                    progress.CapNhatTienTrinh(100, "Mở bảng thẩm định giá...");
                    System.Windows.Forms.Application.DoEvents();
                    progress.Close();

                    _lastThamDinhConfig = config;
                    _lastBoDonGiaId = form.SelectedBoDonGiaId;
                    _lastThamDinhResults = ketQua;

                    // Mở Form Thẩm định đơn giá dạng Modeless độc lập
                    string currentKey = GetActiveWorkbookKey();
                    var donGiaForm = new ThamDinhDonGiaForm(danhSachVatTu, config.Vung, form.SelectedBoDonGiaId);
                    _thamDinhForms[currentKey] = donGiaForm;
                    donGiaForm.FormClosed += (s, ev) =>
                    {
                        _thamDinhForms.TryRemove(currentKey, out _);
                        if (donGiaForm.DialogResult == DialogResult.OK && donGiaForm.SavedBoDonGiaId.HasValue)
                        {
                            _lastBoDonGiaIdByWb[currentKey] = donGiaForm.SavedBoDonGiaId;
                        }
                    };

                    donGiaForm.Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi mở Đơn giá thẩm định: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnMoDonGiaThamDinhClicked(IRibbonControl control)
        {
            try
            {
                if (_thamDinhForm != null && !_thamDinhForm.IsDisposed)
                {
                    _thamDinhForm.WindowState = FormWindowState.Normal;
                    _thamDinhForm.Activate();
                    return;
                }

                var db = new DatabaseManager();
                var repo = new BoDonGiaRepository(db.Context);
                
                using (var form = new ChonBoDonGiaForm(repo))
                {
                    form.Text = "Mở Bộ Đơn Giá Thẩm Định";
                    if (form.ShowDialog() == DialogResult.OK && form.SelectedBoDonGiaId > 0)
                    {
                        int boId = form.SelectedBoDonGiaId;
                        var selectedBo = repo.GetAll().FirstOrDefault(x => x.Id == boId);
                        AIE.Core.Enums.Vung vung = AIE.Core.Enums.Vung.VungII;

                        using var progress = new AIE.ExcelAddIn.Forms.TienTrinhXuLyForm("Mở Bộ Đơn Giá Thẩm Định");
                        progress.Show();
                        progress.CapNhatTienTrinh(30, "Đang tải dữ liệu bộ đơn giá từ CSDL...");
                        System.Windows.Forms.Application.DoEvents();

                        // Sử dụng hàm nạp trực tiếp từ Database với đầy đủ giá gốc, cước VC, bốc xếp, ca máy
                        string currentKey = GetActiveWorkbookKey();
                        var donGiaForm = new ThamDinhDonGiaForm(boId, vung);
                        _thamDinhForms[currentKey] = donGiaForm;
                        if (selectedBo != null)
                        {
                            donGiaForm.SetTenBoDonGia(selectedBo.TenBo, selectedBo.GiaXang, selectedBo.GiaDiezel, selectedBo.GiaDien);
                        }

                        progress.CapNhatTienTrinh(100, "Hoàn tất nạp bộ đơn giá!");
                        System.Windows.Forms.Application.DoEvents();
                        progress.Close();

                        donGiaForm.FormClosed += (s, ev) =>
                        {
                            _thamDinhForms.TryRemove(currentKey, out _);
                            if (donGiaForm.DialogResult == DialogResult.OK && donGiaForm.SavedBoDonGiaId.HasValue)
                            {
                                _lastBoDonGiaIdByWb[currentKey] = donGiaForm.SavedBoDonGiaId;
                            }
                        };

                        // Mở dạng Modeless để không khóa Excel và cho phép Alt+Tab, Minimize
                        donGiaForm.Show();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi mở Bộ Đơn Giá: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnQuanLyDonGiaClicked(IRibbonControl control)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                var db = new AIE.Data.DatabaseManager();
                var vlRepo = new AIE.Data.Repositories.VatLieuRepository(db.Context);
                var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);
                var mayRepo = new AIE.Data.Repositories.MayThiCongRepository(db.Context);
                var mayDmRepo = new AIE.Data.Repositories.DinhMucCaMayRepository(db.Context);

                using var progress = new AIE.ExcelAddIn.Forms.TienTrinhXuLyForm("Quản Lý Đơn Giá");
                progress.Show();
                progress.CapNhatTienTrinh(40, "Đang nạp dữ liệu đơn giá VL, NC, Ca máy...");
                System.Windows.Forms.Application.DoEvents();

                var form = new AIE.ExcelAddIn.Forms.QuanLyDonGiaForm(vlRepo, ncRepo, mayRepo, mayDmRepo);
                Cursor.Current = Cursors.Default;
                
                progress.CapNhatTienTrinh(100, "Mở Quản lý đơn giá...");
                System.Windows.Forms.Application.DoEvents();
                progress.Close();
                
                form.ShowDialog();
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnBaoCaoTDClicked(IRibbonControl control)
        {
            try
            {
                var app = (Application)ExcelDnaUtil.Application;
                var wb = app.ActiveWorkbook;
                if (wb == null)
                {
                    MessageBox.Show("Không có file Excel nào đang mở.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (_lastThamDinhResults == null || _lastThamDinhResults.Count == 0 || _lastThamDinhConfig == null)
                {
                    var choice = MessageBox.Show(
                        "Chưa có dữ liệu kiểm tra định mức, đơn giá trong phiên làm việc hiện tại.\n\nBạn có muốn thực hiện Kiểm tra dự toán trước khi Tổng hợp kinh phí thẩm định không?",
                        "Tổng Hợp Kinh Phí Thẩm Định",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);
                    if (choice == DialogResult.Yes)
                    {
                        OnKiemTraClicked(control);
                    }
                    return;
                }

                if (_xacNhanSaiKhacForm != null && !_xacNhanSaiKhacForm.IsDisposed)
                {
                    _xacNhanSaiKhacForm.WindowState = FormWindowState.Normal;
                    _xacNhanSaiKhacForm.Activate();
                    return;
                }

                string defaultProj = !string.IsNullOrEmpty(CurrentDuToan?.TenCongTrinh) 
                    ? CurrentDuToan.TenCongTrinh 
                    : (!string.IsNullOrEmpty(wb.Name) ? System.IO.Path.GetFileNameWithoutExtension(wb.Name) : "Công trình xây dựng");

                string currentKey = wb.Name ?? GetActiveWorkbookKey();
                var form = new XacNhanSaiKhacForm(_lastThamDinhResults, _lastThamDinhConfig, _lastBoDonGiaId, defaultProj);
                _xacNhanSaiKhacForms[currentKey] = form;

                form.FormClosed += (s, ev) =>
                {
                    _xacNhanSaiKhacForms.TryRemove(currentKey, out _);
                };

                // Mở cửa sổ dạng Modeless: cho phép Alt+Tab chuyển đổi qua lại với Excel, không khóa Excel
                form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi mở Tổng hợp kinh phí thẩm định: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnTaoDuToanMoiClicked(IRibbonControl control)
        {
            try
            {
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                var wb = app.ActiveWorkbook;
                
                // Check if there's an active workbook that is not completely empty
                if (wb != null)
                {
                    var ws = wb.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                    bool isEmpty = false;
                    if (ws != null && ws.Name.StartsWith("Sheet"))
                    {
                        var range = ws.UsedRange;
                        var cell1 = ws.Cells[1, 1] as Microsoft.Office.Interop.Excel.Range;
                        string cellText = cell1 != null ? (cell1.Text?.ToString() ?? "") : "";
                        if (range.Rows.Count <= 1 && range.Columns.Count <= 1 && string.IsNullOrEmpty(cellText))
                        {
                            isEmpty = true;
                        }
                    }

                    if (!isEmpty)
                    {
                        var result = MessageBox.Show(
                            "Đang có dự án mở. Bạn có muốn lưu dự án hiện hành trước khi tạo mới không?\n\n- Chọn Yes để Lưu dự án cũ và Tạo mới\n- Chọn No để Đóng dự án cũ (không lưu) và Tạo mới\n- Chọn Cancel để Hủy thao tác",
                            "Xác nhận đóng dự án",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Question);

                        if (result == DialogResult.Cancel)
                        {
                            return; // Hủy tạo mới
                        }
                        else if (result == DialogResult.Yes)
                        {
                            bool saved = SaveDuToan();
                            if (!saved) return; // Nếu người dùng hủy lưu thì không đóng
                            
                            wb.Close(false);
                        }
                        else if (result == DialogResult.No)
                        {
                            wb.Close(false); // Đóng không lưu
                        }
                    }
                }

                var form = new TaoDuToanMoiForm();
                form.ShowDialog(new WindowWrapper(ExcelDnaUtil.WindowHandle));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khởi tạo dự toán: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnDongBoExcelClicked(IRibbonControl control)
        {
            DongBoDuLieuTuExcel();
        }

        public static void DongBoDuLieuTuExcel()
        {
            try
            {
                if (CurrentDuToan == null)
                {
                    MessageBox.Show("Chưa có dự toán nào đang hoạt động. Vui lòng tạo mới hoặc mở file dự toán!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                TienTrinhXuLyForm.ChayTacVu("Đang đồng bộ dữ liệu từ Excel...", report =>
                {
                    report(20, "Đang đọc khối lượng và danh sách công tác từ Sheet hiện tại...");
                    var excelSvc = new LapDuToanExcelService();
                    var latestBoq = excelSvc.ReadBOQFromActiveSheet();

                    if (latestBoq == null || (latestBoq.DanhSachHangMuc?.Count ?? 0) == 0)
                    {
                        throw new Exception("Không tìm thấy dữ liệu dự toán trên Sheet hiện tại!");
                    }

                    report(60, "Đang gộp và bảo toàn định mức, đơn giá đã tra...");
                    var merged = DuToanSyncHelper.Merge(latestBoq, CurrentDuToan);
                    DuToanSyncHelper.OverwriteInto(CurrentDuToan, merged);

                    report(90, "Đang hoàn tất đồng bộ...");
                    System.Threading.Thread.Sleep(200);
                });

                int soHm = CurrentDuToan.DanhSachHangMuc?.Count ?? 0;
                int soCt = CurrentDuToan.DanhSachHangMuc?.Sum(h => h.DanhSachCongTac?.Count ?? 0) ?? 0;
                MessageBox.Show($"Đồng bộ thành công!\n- Số hạng mục: {soHm}\n- Tổng số công tác: {soCt}", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi đồng bộ từ Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnGoiDonGiaClicked(IRibbonControl control)
        {
            try
            {
                var app = (Application)ExcelDnaUtil.Application;
                var ws = app.ActiveSheet as Worksheet;
                if (ws == null) return;

                Range selection = app.Selection as Range;
                var dbManager = new DatabaseManager();
                var repo = new AIE.Data.Repositories.CongTacRepository(dbManager.Context);

                // Tìm phạm vi dòng dữ liệu trong sheet (từ dòng 6)
                var usedRange = ws.UsedRange;
                int usedMax = usedRange.Rows.Count + usedRange.Row - 1;
                Marshal.ReleaseComObject(usedRange);
                int scanHigh = Math.Max(usedMax, 200);
                int maxRow = Math.Max(usedMax, 6);

                object[,] raw = null;
                try
                {
                    var rawRange = ws.Range[$"A6:D{scanHigh}"];
                    raw = rawRange.Value2 as object[,];
                    Marshal.ReleaseComObject(rawRange);
                }
                catch { }

                // Xác định dòng kết thúc thực sự có dữ liệu
                int dataRows = raw?.GetLength(0) ?? 0;
                if (dataRows > 0)
                {
                    for (int i = 0; i < dataRows; i++)
                    {
                        string a = raw[i + 1, 1]?.ToString()?.Trim() ?? "";
                        string b = raw[i + 1, 2]?.ToString()?.Trim() ?? "";
                        string c = raw[i + 1, 3]?.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(a) || !string.IsNullOrEmpty(b) || !string.IsNullOrEmpty(c))
                        {
                            int r = 6 + i;
                            if (r > maxRow) maxRow = r;
                        }
                    }
                }

                // Tập hợp các dòng được chọn (nếu có)
                var selectedRowIndices = new HashSet<int>();
                if (selection != null)
                {
                    var selRows = selection.Rows;
                    foreach (Range r in selRows)
                    {
                        if (r.Row >= 6) selectedRowIndices.Add(r.Row);
                        Marshal.ReleaseComObject(r);
                    }
                    Marshal.ReleaseComObject(selRows);
                    Marshal.ReleaseComObject(selection);
                    selection = null;
                }

                int updatedCount = 0;
                int currentStt = 0;

                int procRows = maxRow - 6 + 1;
                object[,] writeStt = new object[procRows, 1];
                object[,] writeTen = new object[procRows, 1];
                object[,] writeDonVi = new object[procRows, 1];
                var hangMucIndices = new List<int>();

                // Sao chép giá trị đọc được ban đầu để không xóa dữ liệu các dòng không thay đổi
                for (int i = 0; i < procRows && raw != null && i < dataRows; i++)
                {
                    writeStt[i, 0] = raw[i + 1, 1];
                    writeTen[i, 0] = raw[i + 1, 3];
                    writeDonVi[i, 0] = raw[i + 1, 4];
                }

                for (int i = 0; i < procRows; i++)
                {
                    int rowIndex = 6 + i;
                    string maHieu = raw != null && i < dataRows ? (raw[i + 1, 2]?.ToString()?.Trim() ?? "") : "";
                    string sttText = raw != null && i < dataRows ? (raw[i + 1, 1]?.ToString()?.Trim() ?? "") : "";
                    string tenText = raw != null && i < dataRows ? (raw[i + 1, 3]?.ToString()?.Trim() ?? "") : "";

                    // Kiểm tra xem dòng có phải là dòng Dữ liệu/Hạng mục không
                    if (string.IsNullOrEmpty(maHieu) && string.IsNullOrEmpty(sttText) && string.IsNullOrEmpty(tenText))
                    {
                        continue;
                    }

                    if (!string.IsNullOrEmpty(maHieu))
                    {
                        // Dòng CÔNG TÁC
                        bool shouldLookup = selectedRowIndices.Count <= 1 || selectedRowIndices.Contains(rowIndex) || string.IsNullOrEmpty(tenText);
                        if (shouldLookup)
                        {
                            var congTac = repo.GetByMaHieu(maHieu);
                            if (congTac != null)
                            {
                                writeTen[i, 0] = congTac.TenCongTac;
                                writeDonVi[i, 0] = congTac.DonVi;
                                updatedCount++;
                            }
                        }

                        // Đánh số STT tăng dần nếu chưa có
                        if (string.IsNullOrWhiteSpace(sttText) || !int.TryParse(sttText, out _))
                        {
                            currentStt++;
                            writeStt[i, 0] = currentStt;
                        }
                        else if (int.TryParse(sttText, out int s))
                        {
                            currentStt = s;
                        }
                    }
                    else
                    {
                        // Dòng HẠNG MỤC / tiêu đề: định dạng đậm + nền xanh (trừ dòng TỔNG CỘNG)
                        if (!tenText.StartsWith("TỔNG CỘNG", StringComparison.OrdinalIgnoreCase))
                        {
                            hangMucIndices.Add(i);
                        }
                    }
                }

                // ===== GHI GIÁ TRỊ MỘT LƯỢT (giảm tối đa COM object) =====
                if (procRows > 0)
                {
                    var rngStt = ws.Range[$"A6:A{maxRow}"];
                    rngStt.Value2 = writeStt;
                    Marshal.ReleaseComObject(rngStt);

                    var rngTen = ws.Range[$"C6:C{maxRow}"];
                    rngTen.Value2 = writeTen;
                    Marshal.ReleaseComObject(rngTen);

                    var rngDV = ws.Range[$"D6:D{maxRow}"];
                    rngDV.Value2 = writeDonVi;
                    Marshal.ReleaseComObject(rngDV);
                }

                // ===== ĐỊNH DẠNG TOÀN BẢNG (một lượt, không dùng ws.Cells) =====
                Range full1 = null;
                try
                {
                    full1 = ws.Range[$"A4:K{maxRow}"];
                    var borders = full1.Borders;
                    borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    Marshal.ReleaseComObject(borders);
                    full1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                    // Cột A (STT) canh giữa
                    var colA = ws.Range[$"A6:A{maxRow}"];
                    colA.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    Marshal.ReleaseComObject(colA);

                    // Cột C (Tên) canh đều + wrap
                    var colC = ws.Range[$"C6:C{maxRow}"];
                    colC.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignJustify;
                    colC.WrapText = true;
                    Marshal.ReleaseComObject(colC);

                    // Cột D (Đơn vị) canh giữa
                    var colD = ws.Range[$"D6:D{maxRow}"];
                    colD.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    Marshal.ReleaseComObject(colD);

                    // Cột E..K (số liệu) canh phải + định dạng số
                    var colNum = ws.Range[$"E6:K{maxRow}"];
                    colNum.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    Marshal.ReleaseComObject(colNum);

                    var colE = ws.Range[$"E6:E{maxRow}"];
                    ExcelFormatHelper.ApplyQuantityFormat(colE, 2);
                    Marshal.ReleaseComObject(colE);

                    var colFtoK = ws.Range[$"F6:K{maxRow}"];
                    ExcelFormatHelper.ApplyIntegerFormat(colFtoK);
                    Marshal.ReleaseComObject(colFtoK);

                    // Dòng HẠNG MỤC: gom theo nhóm địa chỉ để định dạng nhanh và không leak COM
                    if (hangMucIndices.Count > 0)
                    {
                        var batch = new System.Text.StringBuilder();
                        int countInBatch = 0;
                        foreach (int idx in hangMucIndices)
                        {
                            int r = 6 + idx;
                            string addr = $"A{r}:K{r}";
                            if (batch.Length > 0) batch.Append(",");
                            batch.Append(addr);
                            countInBatch++;

                            if (batch.Length > 200 || countInBatch >= 15)
                            {
                                var hmBatch = ws.Range[batch.ToString()];
                                hmBatch.Font.Bold = true;
                                hmBatch.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(220, 235, 252));
                                Marshal.ReleaseComObject(hmBatch);
                                batch.Clear();
                                countInBatch = 0;
                            }
                        }
                        if (batch.Length > 0)
                        {
                            var hmBatch = ws.Range[batch.ToString()];
                            hmBatch.Font.Bold = true;
                            hmBatch.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(220, 235, 252));
                            Marshal.ReleaseComObject(hmBatch);
                        }
                    }

                    // Đảm bảo mở lại hiển thị cột 10 (J) nếu vô tình bị ẩn
                    try
                    {
                        var colJ = ws.Range[$"J6:J{maxRow}"];
                        colJ.Hidden = false;
                        Marshal.ReleaseComObject(colJ);
                    }
                    catch { }

                    // Đánh số thứ tự (STT) chuẩn xác, liên tục cho toàn bộ các dòng công tác trên sheet
                    LapDuToanExcelService.DanhLaiSTTCongTac(ws);

                    if (updatedCount == 0 && selectedRowIndices.Count > 1)
                    {
                        MessageBox.Show("Đã chuẩn hóa định dạng bảng và các dòng Hạng mục. Không tìm thấy Mã hiệu mới nào cần tra cứu.", "AIE Dự Toán", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                finally
                {
                    if (full1 != null) Marshal.ReleaseComObject(full1);
                    if (ws != null) Marshal.ReleaseComObject(ws);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi Gọi Đơn giá: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnTongHopKinhPhiClicked(IRibbonControl control)
        {
            try
            {
                var db = new DatabaseManager();
                var ctRepo = new AIE.Data.Repositories.CongTacRepository(db.Context);
                var mayRepo = new AIE.Data.Repositories.MayThiCongRepository(db.Context);
                var dinhMucMayRepo = new AIE.Data.Repositories.DinhMucCaMayRepository(db.Context);
                var vlRepo = new AIE.Data.Repositories.VatLieuRepository(db.Context);
                var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);

                var excelService = new AIE.ExcelAddIn.Services.LapDuToanExcelService();
                var duToan = CurrentDuToan;

                try
                {
                    var latestBoq = excelService.ReadBOQFromActiveSheet();
                    if (latestBoq != null && (latestBoq.DanhSachHangMuc?.Count ?? 0) > 0)
                    {
                        if (duToan == null || duToan.DanhSachHangMuc == null || duToan.DanhSachHangMuc.Count == 0)
                        {
                            duToan = latestBoq;
                        }
                        else
                        {
                            var merged = DuToanSyncHelper.Merge(latestBoq, duToan);
                            DuToanSyncHelper.OverwriteInto(duToan, merged);
                        }
                    }
                }
                catch { }

                if (duToan == null || duToan.DanhSachHangMuc == null || duToan.DanhSachHangMuc.Count == 0 || duToan.DanhSachHangMuc.Sum(hm => hm.DanhSachCongTac.Count) == 0)
                {
                    MessageBox.Show("Không tìm thấy công tác nào trong file Excel hoặc dự toán hiện hành.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Chạy phân tích vật tư nếu chưa có bảng tổng hợp vật tư
                if (duToan.BangTongHop == null || duToan.BangTongHop.DanhSachVatLieu.Count == 0)
                {
                    var service = new AIE.ExcelAddIn.Services.PhanTichVatTuService(ctRepo, mayRepo, dinhMucMayRepo, vlRepo, ncRepo);
                    service.PhanTich(duToan);
                }

                // Tính toán đơn giá chi tiết
                var phanTichDonGiaService = new AIE.ExcelAddIn.Services.PhanTichDonGiaService(ctRepo);
                phanTichDonGiaService.TinhDonGiaChiTiet(duToan);
                
                // Mở Form Tổng hợp kinh phí (Modeless: người dùng có thể click vào Excel)
                if (_tongHopForm != null && !_tongHopForm.IsDisposed)
                {
                    _tongHopForm.Activate();
                    return;
                }
                CurrentDuToan = duToan;
                string currentKey = GetActiveWorkbookKey();
                var form = new TongHopKinhPhiForm(duToan);
                _tongHopForms[currentKey] = form;
                form.FormClosed += (s, ev) => { _tongHopForms.TryRemove(currentKey, out _); };
                form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnTinhTongHopClicked(IRibbonControl control)
        {
            OnTongHopKinhPhiClicked(control);
        }

        public void OnXuatBangTHClicked(IRibbonControl control)
        {
            OnTongHopKinhPhiClicked(control);
        }

        public void OnTongHopDuToanClicked(IRibbonControl control)
        {
            OnTongHopKinhPhiClicked(control);
        }

        public void OnTongMucDauTuClicked(IRibbonControl control)
        {
            OnTongHopKinhPhiClicked(control);
        }

        private void MoFormTongHopKinhPhi(bool macDinhTongMucDauTu)
        {
            try
            {
                var duToan = CurrentDuToan;
                if (duToan == null)
                {
                    try
                    {
                        var excelService = new AIE.ExcelAddIn.Services.LapDuToanExcelService();
                        duToan = excelService.ReadBOQFromActiveSheet();
                    }
                    catch { }

                    if (duToan == null)
                    {
                        duToan = new DuToan();
                    }
                }

                if (_tongHopForm != null && !_tongHopForm.IsDisposed)
                {
                    _tongHopForm.Activate();
                    return;
                }
                CurrentDuToan = duToan;
                string currentKey = GetActiveWorkbookKey();
                var form = new TongHopKinhPhiForm(duToan, macDinhTongMucDauTu);
                _tongHopForms[currentKey] = form;
                form.FormClosed += (s, ev) => { _tongHopForms.TryRemove(currentKey, out _); };
                form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi mở Bảng Tổng hợp kinh phí: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void MoTongHopKinhPhiTuShortcut()
        {
            try
            {
                var duToan = CurrentDuToan;
                if (duToan == null)
                {
                    try
                    {
                        var excelService = new AIE.ExcelAddIn.Services.LapDuToanExcelService();
                        duToan = excelService.ReadBOQFromActiveSheet();
                    }
                    catch { }

                    if (duToan == null)
                    {
                        duToan = new DuToan();
                    }
                }

                if (_tongHopForm != null && !_tongHopForm.IsDisposed)
                {
                    _tongHopForm.Activate();
                    return;
                }
                CurrentDuToan = duToan;
                string currentKey = GetActiveWorkbookKey();
                var form = new TongHopKinhPhiForm(duToan);
                _tongHopForms[currentKey] = form;
                form.FormClosed += (s, ev) => { _tongHopForms.TryRemove(currentKey, out _); };
                form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi mở Tổng hợp kinh phí: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnTinhGiaHienTruongClicked(IRibbonControl control)
        {
            try
            {
                var db = new DatabaseManager();
                var ctRepo = new AIE.Data.Repositories.CongTacRepository(db.Context);
                var mayRepo = new AIE.Data.Repositories.MayThiCongRepository(db.Context);
                var dinhMucMayRepo = new AIE.Data.Repositories.DinhMucCaMayRepository(db.Context);
                var vlRepo = new AIE.Data.Repositories.VatLieuRepository(db.Context);
                var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);

                var excelService = new AIE.ExcelAddIn.Services.LapDuToanExcelService();
                var duToan = excelService.ReadBOQFromActiveSheet();

                if (duToan == null || duToan.DanhSachHangMuc == null || duToan.DanhSachHangMuc.Sum(hm => hm.DanhSachCongTac.Count) == 0)
                {
                    MessageBox.Show("Không tìm thấy công tác nào trong file Excel.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Giữ lại BoDonGiaId nếu đã mở từ file .dt trước đó
                if (CurrentDuToan != null && CurrentDuToan.BoDonGiaId.HasValue)
                {
                    duToan.BoDonGiaId = CurrentDuToan.BoDonGiaId;
                }
                else
                {
                    var boDonGiaRepo = new AIE.Data.Repositories.BoDonGiaRepository(db.Context);
                    using var frmBoDonGia = new Forms.ChonBoDonGiaForm(boDonGiaRepo);
                    if (frmBoDonGia.ShowDialog() == DialogResult.OK)
                    {
                        duToan.BoDonGiaId = frmBoDonGia.SelectedBoDonGiaId;
                    }
                    else
                    {
                        return; // Hủy tính giá nếu không chọn bộ đơn giá
                    }
                }

                var service = new AIE.ExcelAddIn.Services.PhanTichVatTuService(ctRepo, mayRepo, dinhMucMayRepo, vlRepo, ncRepo);
                service.PhanTich(duToan);

                // Bảo toàn thông tin cước và cấu hình đã tính từ phiên trước nếu có
                if (CurrentDuToan?.BangTongHop?.DanhSachVatLieu != null)
                {
                    var oldVlDict = CurrentDuToan.BangTongHop.DanhSachVatLieu
                        .GroupBy(x => x.MaVatTu)
                        .ToDictionary(g => g.Key, g => g.First());

                    foreach (var vl in duToan.BangTongHop.DanhSachVatLieu)
                    {
                        if (oldVlDict.TryGetValue(vl.MaVatTu, out var oldVl))
                        {
                            vl.GiaGoc = oldVl.GiaGoc;
                            vl.ChiPhiBocXep = oldVl.ChiPhiBocXep;
                            vl.CuocVCOTo = oldVl.CuocVCOTo;
                            vl.CuocVCBo = oldVl.CuocVCBo;
                            vl.MaDinhMucBocXep = oldVl.MaDinhMucBocXep;
                            vl.PhamViBocXep = oldVl.PhamViBocXep;
                            vl.DmNCBocXep = oldVl.DmNCBocXep;
                            vl.DmMayBocXep = oldVl.DmMayBocXep;
                            vl.MaMayBocXep = oldVl.MaMayBocXep;
                            vl.MaDinhMucVCOTo = oldVl.MaDinhMucVCOTo;
                            vl.MaMayVCOTo = oldVl.MaMayVCOTo;
                            vl.MaDinhMucVCBo = oldVl.MaDinhMucVCBo;
                        }
                    }
                }

                // Load giá từ Bộ Đơn Giá vừa chọn (Ghi đè giá gốc từ thư viện chung)
                var boDonGiaRepo2 = new AIE.Data.Repositories.BoDonGiaRepository(db.Context);
                var giaVl = boDonGiaRepo2.GetGiaVL(duToan.BoDonGiaId.Value).GroupBy(x => x.MaVL).ToDictionary(g => g.Key, g => g.First());
                var giaNc = boDonGiaRepo2.GetGiaNC(duToan.BoDonGiaId.Value).GroupBy(x => x.MaNC).ToDictionary(g => g.Key, g => g.First());
                var giaMay = boDonGiaRepo2.GetGiaMay(duToan.BoDonGiaId.Value).GroupBy(x => x.MaMay).ToDictionary(g => g.Key, g => g.First());

                foreach (var vl in duToan.BangTongHop.DanhSachVatLieu)
                {
                    if (giaVl.TryGetValue(vl.MaVatTu, out var g))
                    {
                        vl.GiaGoc = g.GiaGoc;
                        if (g.ChiPhiBocXep > 0 || g.CuocVCOTo > 0 || g.CuocVCBo > 0)
                        {
                            vl.ChiPhiBocXep = g.ChiPhiBocXep;
                            vl.CuocVCOTo = g.CuocVCOTo;
                            vl.CuocVCBo = g.CuocVCBo;
                        }
                        else if (vl.ChiPhiBocXep == 0 && vl.CuocVCOTo == 0 && vl.CuocVCBo == 0 && g.CuocVC > 0)
                        {
                            vl.CuocVanChuyen = g.CuocVC;
                        }
                    }
                }
                
                foreach (var nc in duToan.BangTongHop.DanhSachNhanCong)
                    if (giaNc.TryGetValue(nc.MaVatTu, out var g)) nc.GiaGoc = g.DonGia;
                
                foreach (var may in duToan.BangTongHop.DanhSachMay)
                    if (giaMay.TryGetValue(may.MaVatTu, out var g)) may.GiaGoc = g.DonGia;

                var phanTichDonGiaService = new AIE.ExcelAddIn.Services.PhanTichDonGiaService(ctRepo);

                // Lưu vào bộ nhớ chung để có thể Save ra file .dt
                CurrentDuToan = duToan;

                using var progress = new AIE.ExcelAddIn.Forms.TienTrinhXuLyForm("Giá Hiện Trường VL, NC, MTC");
                progress.Show();
                progress.CapNhatTienTrinh(50, "Đang khởi tạo bảng tính Giá VL, NC, MTC...");
                System.Windows.Forms.Application.DoEvents();
                
                if (_tinhGiaForm != null && !_tinhGiaForm.IsDisposed)
                {
                    progress.Close();
                    _tinhGiaForm.Activate();
                    return;
                }
                string currentKey = GetActiveWorkbookKey();
                var form = new TinhGiaHienTruongForm(duToan, service, phanTichDonGiaService);
                _tinhGiaForms[currentKey] = form;
                form.FormClosed += (s, ev) => { _tinhGiaForms.TryRemove(currentKey, out _); };
                
                progress.CapNhatTienTrinh(100, "Mở bảng giá...");
                System.Windows.Forms.Application.DoEvents();
                progress.Close();
                
                form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ==================== LƯU / MỞ FILE DỰ TOÁN (.dt) ====================

        public void OnLuuDuToanClicked(IRibbonControl control)
        {
            SaveDuToan();
        }

        public bool SaveDuToan()
        {
            try
            {
                var excelService = new AIE.ExcelAddIn.Services.LapDuToanExcelService();
                var latestDuToan = excelService.ReadBOQFromActiveSheet();

                if (CurrentDuToan != null)
                {
                    // Gộp dữ liệu: BOQ từ sheet (latestDuToan) + toàn bộ dữ liệu kinh phí/cấu hình từ bộ nhớ (CurrentDuToan).
                    // GIỮ NGUYÊN instance CurrentDuToan để form Tổng hợp kinh phí đang mở không bị lệch tham chiếu.
                    var merged = DuToanSyncHelper.Merge(latestDuToan, CurrentDuToan);
                    DuToanSyncHelper.OverwriteInto(CurrentDuToan, merged);
                }
                else
                {
                    CurrentDuToan = latestDuToan;
                }

                // Đọc lại hệ số điều chỉnh (VL/NC/M) theo hạng mục do người dùng nhập tay trên sheet HeSo_DieuChinh
                excelService.DocHeSoDieuChinhTuSheet(CurrentDuToan);

                string filePath = CurrentFilePath;

                // Nếu chưa có đường dẫn, hỏi người dùng
                if (string.IsNullOrEmpty(filePath))
                {
                    using var sfd = new SaveFileDialog();
                    sfd.Filter = "File Dự Toán (*.dt)|*.dt";
                    sfd.Title = "Lưu Dự Toán";
                    string safeName = AIE.Core.Services.Shared.TextHelper.SanitizeFileName(CurrentDuToan.TenCongTrinh, "DuToan_Moi");
                    sfd.FileName = safeName + ".dt";

                    if (sfd.ShowDialog() != DialogResult.OK) return false;
                    filePath = sfd.FileName;
                }

                AIE.ExcelAddIn.Services.DuToanFileService.Save(CurrentDuToan, filePath);
                CurrentFilePath = filePath;
                AIE.ExcelAddIn.Services.AutoSaveManager.DeleteAutoSave(GetActiveWorkbookKey());

                MessageBox.Show(
                    $"Đã lưu dự toán thành công!\n\nFile: {filePath}",
                    "Lưu Dự Toán", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lưu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public void OnMoDuToanClicked(IRibbonControl control)
        {
            try
            {
                using var ofd = new OpenFileDialog();
                ofd.Filter = "File Dự Toán (*.dt)|*.dt|Tất cả file (*.*)|*.*";
                ofd.Title = "Mở Dự Toán";

                if (ofd.ShowDialog() != DialogResult.OK) return;

                MoDuToanTuFile(ofd.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi mở: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnPhucHoiAutoSaveClicked(IRibbonControl control)
        {
            try
            {
                AIE.ExcelAddIn.Services.AutoSaveManager.CheckAndPromptRecovery();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi kiểm tra phục hồi tự động: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Mở file .dt: tạo workbook mới, ghi BOQ ra sheet, xuất các bảng thành phần nếu đã có ChiPhiXD.
        /// KHÔNG xuất TH_ChiPhiXD / TongMucDauTu / TH_DuToan / HeSo_DieuChinh — các sheet này chỉ do
        /// modal Tổng hợp kinh phí xuất, tránh ghi dữ liệu cũ.
        /// </summary>
        public static bool MoDuToanTuDuongDan(string filePath)
        {
            var duToan = AIE.ExcelAddIn.Services.DuToanFileService.Load(filePath);
            var app = (Application)ExcelDnaUtil.Application;
            var wb = app.Workbooks.Add(); // Tạo workbook mới
            string wbKey = wb.Name;
            _duToanByWb[wbKey] = duToan;
            _filePathByWb[wbKey] = filePath;

            bool oldEvents = app.EnableEvents;
            AIE.ExcelAddIn.Services.SheetAutoLookupService.IsSuspended = true;
            app.EnableEvents = false;
            try
            {
                // Ghi dữ liệu ra Excel (sheet hiện tại) để người dùng có thể xem/chỉnh sửa
                var excelService = new AIE.ExcelAddIn.Services.LapDuToanExcelService();
                excelService.WriteBOQToActiveSheet(duToan);

                // Luôn tái tính đơn giá + liên kết lại sheet DuToan theo Bảng tổng hợp hiện tại,
                // tránh sheet giữ đơn giá cũ (lệch với modal Tổng hợp kinh phí khi xuất).
                var hasCongTac = duToan.DanhSachHangMuc != null &&
                                 duToan.DanhSachHangMuc.Sum(hm => hm.DanhSachCongTac?.Count ?? 0) > 0;
                var bth = duToan.BangTongHop;
                var hasBangTongHop = bth != null &&
                                     ((bth.DanhSachVatLieu?.Count ?? 0) + (bth.DanhSachNhanCong?.Count ?? 0) + (bth.DanhSachMay?.Count ?? 0)) > 0;
                if (hasCongTac && (duToan.ChiPhiXD != null || hasBangTongHop))
                {
                    var xuatService = new AIE.ExcelAddIn.Services.XuatBangBieuService();
                    xuatService.ApGiaVaLienKetDuToan(duToan);

                    if (xuatService.LastDonGiaDrift > 1000m)
                    {
                        MessageBox.Show(
                            $"Đơn giá trong file lưu đã cũ so với Bảng tổng hợp hiện tại.\n" +
                            $"Đã tự động cập nhật lại đơn giá và các bảng liên quan.\n\n" +
                            $"Tổng chênh lệch tạm tính: {xuatService.LastDonGiaDrift:N0} đồng.",
                            "Cập nhật đơn giá", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            finally
            {
                app.EnableEvents = oldEvents;
                AIE.ExcelAddIn.Services.SheetAutoLookupService.IsSuspended = false;
            }

            MessageBox.Show(
                $"Đã mở dự toán thành công!\n\nCông trình: {duToan.TenCongTrinh}\nLoại: {duToan.LoaiCongTrinh}\nFile: {filePath}",
                "Mở Dự Toán", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return true;
        }

        private bool MoDuToanTuFile(string filePath)
        {
            return MoDuToanTuDuongDan(filePath);
        }

        /// <summary>
        /// Khôi phục file .dt hiện tại từ bản lưu .bak gần nhất (được tạo tự động mỗi lần Lưu),
        /// sau đó tải lại dự toán vào workbook mới.
        /// </summary>
        public void OnKhoiPhucBamLuuClicked(IRibbonControl control)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(CurrentFilePath) || !System.IO.File.Exists(CurrentFilePath))
                {
                    MessageBox.Show("Chưa có dự toán nào được mở/lưu trong phiên làm việc này.", "Khôi phục bản lưu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string bak = CurrentFilePath + ".bak";
                if (!System.IO.File.Exists(bak))
                {
                    MessageBox.Show($"Không tìm thấy bản lưu gần nhất:\n{bak}\n\n(Lưu lại dự toán ít nhất 1 lần để tạo bản lưu).", "Khôi phục bản lưu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var confirm = MessageBox.Show(
                    "Phục hồi dự toán hiện tại từ bản lưu .bak gần nhất?\n\n" + CurrentFilePath + "\n\nCác thay đổi sau lần lưu gần nhất sẽ bị mất.",
                    "Khôi phục bản lưu",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (confirm != DialogResult.Yes) return;

                bool restored = AIE.ExcelAddIn.Services.DuToanFileService.KhoiPhucTuBanLuu(CurrentFilePath);
                if (!restored)
                {
                    MessageBox.Show("Không thể phục hồi bản lưu.", "Khôi phục bản lưu", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                MessageBox.Show("Đã phục hồi file từ bản lưu .bak. Đang tải lại dự toán...", "Khôi phục bản lưu", MessageBoxButtons.OK, MessageBoxIcon.Information);
                MoDuToanTuFile(CurrentFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi khôi phục: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
