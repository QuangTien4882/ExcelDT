using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using AIE.Core.Models;
using AIE.Core.Services;
using AIE.Core.Services.LapDuToan;
using AIE.Core.Services.Shared;
using AIE.Data;
using AIE.Data.Repositories;
using AIE.ExcelAddIn.Helpers;
using AIE.ExcelAddIn.Services;
using ExcelDna.Integration;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelWb = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Font = System.Drawing.Font;
using TextBox = System.Windows.Forms.TextBox;
using CheckBox = System.Windows.Forms.CheckBox;
using Label = System.Windows.Forms.Label;
using Button = System.Windows.Forms.Button;
using GroupBox = System.Windows.Forms.GroupBox;
using Panel = System.Windows.Forms.Panel;
using ComboBox = System.Windows.Forms.ComboBox;
using RadioButton = System.Windows.Forms.RadioButton;

namespace AIE.ExcelAddIn.Forms
{
    public class TongHopKinhPhiForm : Form
    {
        private DuToan _duToan;
        private BangTongHopKinhPhiModel _model;
        private XuatBangBieuService _xuatService;
        private bool _dangKiemTraDinhMuc = false;

        // TabControl chính
        private TabControl tabMain;
        private TabPage tabChiPhiXD;
        private TabPage tabTMDT;

        // Controls Tab 1: Chi phí Xây dựng (Bảng 3.8 TT 36)
        private ComboBox cboHangMucSelector;
        private HangMuc _currentHangMuc;
        private bool _isChangingHangMuc = false;
        private ComboBox cboGiaiDoanXD;
        private ComboBox cboLoaiCongTrinhXD;
        private ComboBox cboPhanLoaiPhuXD;
        private TextBox txtQuyMoXD;
        private string _lastCustomQuyMo = "15";
        private bool _isUpdatingQuyMoXD = false;
        private ComboBox cboLoaiNhaTamXD;
        private CheckBox chkVungSauXaXD;
        private TextBox txtCPCXD;
        private TextBox txtTTXD;
        private TextBox txtTNCTTTXD;
        private TextBox txtGTGTXD;
        private TextBox txtNhaTamXD;
        private DataGridView dgvPreviewChiPhiXD;
        private Button btnChuyenSangTab2;

        private ChiPhiXayDungCalc _calcService;
        private DinhMucCPCRepository _cpcRepo;
        private DinhMucTTRepository _ttRepo;
        private System.Data.IDbConnection _dbConn;
        private DatabaseManager _dbManager;
        private decimal _tongT = 0;
        private decimal _tongNC = 0;
        private decimal _tongVL = 0;
        private decimal _tongMay = 0;
        private Label lblBangChu;
        private ToolTip _toolTip = new ToolTip();

        // Controls Tab 2: Chế độ bảng tính (Yêu cầu 2)
        private RadioButton radBangTHDT;
        private RadioButton radBangTMDT;
        public bool LaCheDoTongMucDauTu => radBangTMDT != null && radBangTMDT.Checked;

        // Controls thông số công trình
        private ComboBox cboLoaiCongTrinh;
        private ComboBox cboCapCongTrinh;
        private ComboBox cboSoBuocThietKe;

        private TextBox txtChiPhiXD;
        private TextBox txtChiPhiNT;
        private TextBox txtChiPhiTB;
        private TextBox txtChiPhiBT;

        // CheckBoxes điều kiện áp dụng hệ số (TT 38 & BTC)
        private CheckBox chkThietBi50;
        private CheckBox chkDaKiemToan;
        private CheckBox chkThueThamTra;
        private CheckBox chkCdtTuQL;
        private CheckBox chkVungKhoKhan;
        private CheckBox chkTuyenNhieuTinh;
        private CheckBox chkCaiTao;
        private CheckBox chkLapLai;

        // Thuế VAT chung (Yêu cầu 7)
        private ComboBox cboVATChung;

        // Grid chi phí (Yêu cầu 1, 3, 4, 6, 7)
        private DataGridView dgvChiPhi;

        // Labels tổng cộng
        private Label lblTongTruocThue;
        private Label lblTongGTGT;
        private Label lblTongSauThue;
        private Label lblTieuDeTongSauThue;

        // Buttons thao tác
        private Button btnThemThuVien;
        private Button btnThemTuyBien;
        private Button btnXoaChiPhi;
        private Button btnKhoiPhuc;
        private Button btnTraLaiDinhMuc;

        // Controls Bottom Action Bar
        private TableLayoutPanel pnlSummary;
        private Button btnXuatExcelChinh;
        private Button btnLuu;
        private Button btnDong;

        private bool _isUpdating = false;
        private bool _isSyncingLoaiCT = false;
        private bool _isSyncingVAT = false;
        private bool _isFromThamDinh = false;
        private System.Windows.Forms.Timer _tyLeDebounce;

        public TongHopKinhPhiForm(DuToan duToan, bool macDinhTongMucDauTu = false, bool isFromThamDinh = false)
        {
            _duToan = duToan ?? new DuToan();
            _isFromThamDinh = isFromThamDinh;
            _xuatService = new XuatBangBieuService();

            KhoiTaoDuLieu();
            InitializeComponent();

            if (_isFromThamDinh)
            {
                this.Text = "Bảng Tổng Hợp Kinh Phí Thẩm Định Dự Toán (Thông tư 36/2026/TT-BXD & TT 38/2026/TT-BXD)";
            }

            if (macDinhTongMucDauTu)
            {
                if (radBangTMDT != null) radBangTMDT.Checked = true;
                if (tabMain != null && tabTMDT != null) tabMain.SelectedTab = tabTMDT;
            }

            NapDuLieuLenGiaoDien();

            FormStateHelper.Attach(this);

            this.Shown += (s, e) =>
            {
                dgvPreviewChiPhiXD?.AutoFit();
                dgvChiPhi?.AutoFit();
            };

            if (tabMain != null)
            {
                tabMain.SelectedIndexChanged += (s, e) =>
                {
                    if (tabMain.SelectedTab == tabChiPhiXD) dgvPreviewChiPhiXD?.AutoFit();
                    else if (tabMain.SelectedTab == tabTMDT) dgvChiPhi?.AutoFit();
                };
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _dbConn?.Dispose();
                _dbConn = null;
                _dbManager?.Dispose();
                _dbManager = null;
            }
            base.Dispose(disposing);
        }

        private void KhoiTaoDuLieu()
        {
            _dbManager = new DatabaseManager();
            _dbConn = _dbManager.Context.GetConnection();
            _cpcRepo = new DinhMucCPCRepository(_dbConn);
            _ttRepo = new DinhMucTTRepository(_dbConn);
            _calcService = new ChiPhiXayDungCalc();

            _tongVL = _duToan.DanhSachHangMuc?.Sum(hm => hm.TongVL) ?? 0m;
            _tongNC = _duToan.DanhSachHangMuc?.Sum(hm => hm.TongNC) ?? 0m;
            _tongMay = _duToan.DanhSachHangMuc?.Sum(hm => hm.TongMay) ?? 0m;
            _tongT = _tongVL + _tongNC + _tongMay;

            bool hasScanned = false;
            if (_duToan.BangKinhPhi != null && _duToan.BangKinhPhi.Items.Count > 0)
            {
                _model = _duToan.BangKinhPhi;
            }
            else
            {
                if (!_isFromThamDinh)
                {
                    hasScanned = QuetDuLieuTuSheetTHChiPhiXD();
                }
                if (!hasScanned)
                {
                    string loaiCT = !string.IsNullOrEmpty(_duToan.LoaiCongTrinh) ? _duToan.LoaiCongTrinh : "Dân dụng";
                    string capCT = !string.IsNullOrEmpty(_duToan.CapCongTrinh) ? _duToan.CapCongTrinh : "Cấp III";

                    if (_duToan.ChiPhiXD == null)
                    {
                        decimal cpc = 0m;
                        var listCPC = _cpcRepo.GetByLoaiCongTrinh(loaiCT, null);
                        if (listCPC != null && listCPC.Count > 0)
                            cpc = InterpolationHelper.NoiSuyTiLeCPC(listCPC, _tongT);
                        else
                            cpc = 7.3m;

                        decimal tt = 0m;
                        var objTT = _ttRepo.GetByLoaiCongTrinh(loaiCT, null);
                        if (objTT != null)
                            tt = objTT.TiLe;
                        else
                            tt = 2.5m;

                        decimal tncttt = (loaiCT == "Công nghiệp" || loaiCT == "Giao thông") ? 6.0m : 5.5m;
                        decimal gtgt = 10.0m;
                        decimal nhatam = 1.2m;

                        _duToan.ChiPhiXD = _calcService.Tinh(_tongVL, _tongNC, _tongMay, cpc, tt, tncttt, gtgt, nhatam, "T");
                    }

                    decimal gXD = _duToan.ChiPhiXD?.G ?? 0m;
                    if (gXD <= 0)
                    {
                        decimal tongT = _tongT;
                        gXD = tongT > 0 ? tongT : 10_000_000_000m;
                    }
                    decimal ntTruocThue = Math.Round(gXD * (_duToan.ChiPhiXD?.TiLeNhaTam > 0 ? _duToan.ChiPhiXD.TiLeNhaTam : 1.2m) / 100m, 0);

                    _model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                        loaiCT: !string.IsNullOrEmpty(loaiCT) ? loaiCT : "Dân dụng",
                        capCT: !string.IsNullOrEmpty(capCT) ? capCT : "Cấp III",
                        soBuocTK: _duToan.SoBuocThietKe > 0 ? _duToan.SoBuocThietKe : 2,
                        chiPhiXD: gXD,
                        chiPhiTB: _duToan.ChiPhiThietBi,
                        chiPhiBT: 0m,
                        chiPhiNhaTam: ntTruocThue
                    );
                    _model.LoaiCongTrinh = loaiCT;
                    _model.CapCongTrinh = capCT;
                    if (_duToan.SoBuocThietKe == 0) _model.SoBuocThietKe = 0;
                }
            }
        }

        /// <summary>
        /// Yêu cầu 5: Tự động quét và liên kết Chi phí xây dựng từ sheet TH_ChiPhiXD trong workbook hiện tại
        /// Quét toàn bộ: VL, NC, M, CPC, TT, TL, G (trước thuế), GTGT, Gxd (sau thuế), LT (nhà tạm)
        /// </summary>
        private bool QuetDuLieuTuSheetTHChiPhiXD()
        {
            try
            {
                var app = (ExcelApp)ExcelDnaUtil.Application;
                var wb = app?.ActiveWorkbook;
                if (wb == null) return false;

                Microsoft.Office.Interop.Excel.Worksheet wsTH = null;
                foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in wb.Sheets)
                {
                    if (sheet.Name == "TH_ChiPhiXD")
                    {
                        wsTH = sheet;
                        break;
                    }
                }

                if (wsTH == null)
                {
                    // Thử quét từ sheet DuToan nếu có sẵn trong workbook
                    try
                    {
                        foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in wb.Sheets)
                        {
                            if (sheet.Name.StartsWith("DuToan", StringComparison.OrdinalIgnoreCase))
                            {
                                for (int r = 6; r <= 1000; r++)
                                {
                                    string c = sheet.Cells[r, 3]?.Value2?.ToString()?.Trim() ?? "";
                                    if (c == "TỔNG CỘNG" || c.Contains("TỔNG CỘNG") || c == "CỘNG")
                                    {
                                        object valI = sheet.Cells[r, 9]?.Value2;
                                        object valJ = sheet.Cells[r, 10]?.Value2;
                                        object valK = sheet.Cells[r, 11]?.Value2;
                                        decimal vI = valI != null ? UIHelper.ParseTien(valI.ToString()) : 0m;
                                        decimal vJ = valJ != null ? UIHelper.ParseTien(valJ.ToString()) : 0m;
                                        decimal vK = valK != null ? UIHelper.ParseTien(valK.ToString()) : 0m;
                                        if (vI > 0) _tongVL = vI;
                                        if (vJ > 0) _tongNC = vJ;
                                        if (vK > 0) _tongMay = vK;
                                        if (_tongVL > 0 || _tongNC > 0 || _tongMay > 0)
                                            _tongT = _tongVL + _tongNC + _tongMay;
                                        break;
                                    }
                                }
                                break;
                            }
                        }
                    }
                    catch { }
                    return false;
                }

                decimal scannedVL = 0m, scannedNC = 0m, scannedM = 0m;
                decimal gXD = 0m;
                decimal tienThueXD = 0m;
                decimal ltNhaTam = 0m;
                decimal rateCPC = 0m, rateTT = 0m, rateTL = 0m, rateGTGT = 0m, rateLT = 0m;

                // Quét thông minh theo Ký hiệu ở cột C hoặc Tên ở cột B
                for (int r = 4; r <= 25; r++)
                {
                    string kyHieu = wsTH.Cells[r, 3]?.Value2?.ToString()?.Trim() ?? "";
                    string noiDung = wsTH.Cells[r, 2]?.Value2?.ToString()?.Trim() ?? "";
                    string cachTinh = wsTH.Cells[r, 4]?.Value2?.ToString()?.Trim() ?? "";
                    object valE = wsTH.Cells[r, 5]?.Value2;
                    decimal num = 0m;
                    if (valE != null) decimal.TryParse(valE.ToString(), out num);

                    if (kyHieu == "VL" || noiDung.Contains("Chi phí vật liệu"))
                    {
                        if (num > 0) scannedVL = num;
                    }
                    else if (kyHieu == "NC" || noiDung.Contains("Chi phí nhân công"))
                    {
                        if (num > 0) scannedNC = num;
                    }
                    else if (kyHieu == "M" || noiDung.Contains("Chi phí máy"))
                    {
                        if (num > 0) scannedM = num;
                    }
                    else if (kyHieu == "CPC" || noiDung.Contains("Chi phí chung"))
                    {
                        var m = System.Text.RegularExpressions.Regex.Match(cachTinh, @"([\d,\.]+)%");
                        if (m.Success) decimal.TryParse(m.Groups[1].Value.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out rateCPC);
                    }
                    else if (kyHieu == "TT" || noiDung.Contains("không xác định"))
                    {
                        var m = System.Text.RegularExpressions.Regex.Match(cachTinh, @"([\d,\.]+)%");
                        if (m.Success) decimal.TryParse(m.Groups[1].Value.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out rateTT);
                    }
                    else if (kyHieu == "TL" || noiDung.Contains("chịu thuế tính trước"))
                    {
                        var m = System.Text.RegularExpressions.Regex.Match(cachTinh, @"([\d,\.]+)%");
                        if (m.Success) decimal.TryParse(m.Groups[1].Value.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out rateTL);
                    }
                    else if (kyHieu == "G" || noiDung.Contains("Chi phí xây dựng trước thuế"))
                    {
                        if (num > 0) gXD = num;
                    }
                    else if (kyHieu == "GTGT" || noiDung.Contains("Thuế giá trị gia tăng"))
                    {
                        if (num > 0) tienThueXD = num;
                        var m = System.Text.RegularExpressions.Regex.Match(cachTinh, @"([\d,\.]+)%");
                        if (m.Success) decimal.TryParse(m.Groups[1].Value.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out rateGTGT);
                    }
                    else if (kyHieu == "LT" || noiDung.Contains("nhà tạm"))
                    {
                        if (num > 0) ltNhaTam = num;
                        var m = System.Text.RegularExpressions.Regex.Match(cachTinh, @"([\d,\.]+)%");
                        if (m.Success) decimal.TryParse(m.Groups[1].Value.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out rateLT);
                    }
                }

                // Fallback cố định nếu không tìm thấy theo text
                if (scannedVL == 0m)
                {
                    object valE7 = wsTH.Range["E7"].Value2;
                    if (valE7 != null) decimal.TryParse(valE7.ToString(), out scannedVL);
                }
                if (scannedNC == 0m)
                {
                    object valE8 = wsTH.Range["E8"].Value2;
                    if (valE8 != null) decimal.TryParse(valE8.ToString(), out scannedNC);
                }
                if (scannedM == 0m)
                {
                    object valE9 = wsTH.Range["E9"].Value2;
                    if (valE9 != null) decimal.TryParse(valE9.ToString(), out scannedM);
                }
                if (gXD == 0m)
                {
                    object valE14 = wsTH.Range["E14"].Value2;
                    object valE15 = wsTH.Range["E15"].Value2;
                    object valE17 = wsTH.Range["E17"].Value2;
                    if (valE14 != null) decimal.TryParse(valE14.ToString(), out gXD);
                    if (valE15 != null) decimal.TryParse(valE15.ToString(), out tienThueXD);
                    if (valE17 != null) decimal.TryParse(valE17.ToString(), out ltNhaTam);

                    if (gXD == 0m)
                    {
                        object valE12 = wsTH.Range["E12"].Value2;
                        object valE13 = wsTH.Range["E13"].Value2;
                        object valOld15 = wsTH.Range["E15"].Value2;
                        if (valE12 != null) decimal.TryParse(valE12.ToString(), out gXD);
                        if (valE13 != null) decimal.TryParse(valE13.ToString(), out tienThueXD);
                        if (valOld15 != null) decimal.TryParse(valOld15.ToString(), out ltNhaTam);
                    }
                }

                if (_tongT == 0)
                {
                    if (scannedVL > 0) _tongVL = scannedVL;
                    if (scannedNC > 0) _tongNC = scannedNC;
                    if (scannedM > 0) _tongMay = scannedM;
                    if (_tongVL > 0 || _tongNC > 0 || _tongMay > 0)
                    {
                        _tongT = _tongVL + _tongNC + _tongMay;
                    }
                }

                if (_duToan.ChiPhiXD == null) _duToan.ChiPhiXD = new ChiPhiXayDung();
                if (rateCPC > 0) _duToan.ChiPhiXD.TiLeCPC = rateCPC;
                if (rateTT > 0) _duToan.ChiPhiXD.TiLeTT = rateTT;
                if (rateTL > 0) _duToan.ChiPhiXD.TiLeTNCTTT = rateTL;
                if (rateGTGT > 0) _duToan.ChiPhiXD.TiLeGTGT = rateGTGT;
                if (rateLT > 0) _duToan.ChiPhiXD.TiLeNhaTam = rateLT;

                if (gXD > 0)
                {
                    decimal vatRate = 0.10m;
                    if (tienThueXD > 0)
                    {
                        vatRate = Math.Round(tienThueXD / gXD, 2);
                    }

                    decimal ntTruocThue = 0m;
                    if (ltNhaTam > 0)
                    {
                        ntTruocThue = Math.Round(ltNhaTam / (1m + vatRate), 0);
                    }

                    if (_model == null)
                    {
                        string loaiCT = !string.IsNullOrEmpty(_duToan.LoaiCongTrinh) ? _duToan.LoaiCongTrinh : "";
                        string capCT = !string.IsNullOrEmpty(_duToan.CapCongTrinh) ? _duToan.CapCongTrinh : "";
                        int soBuoc = _duToan.SoBuocThietKe > 0 ? _duToan.SoBuocThietKe : 2;

                        _model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                            loaiCT: !string.IsNullOrEmpty(loaiCT) ? loaiCT : "Dân dụng",
                            capCT: !string.IsNullOrEmpty(capCT) ? capCT : "Cấp III",
                            soBuocTK: soBuoc,
                            chiPhiXD: gXD,
                            chiPhiTB: _duToan.ChiPhiThietBi,
                            chiPhiBT: 0m,
                            chiPhiNhaTam: ntTruocThue
                        );
                        _model.LoaiCongTrinh = loaiCT;
                        _model.CapCongTrinh = capCT;
                        if (_duToan.SoBuocThietKe == 0) _model.SoBuocThietKe = 0;
                    }
                    else
                    {
                        _model.ChiPhiXDTruocThue = gXD;
                        _model.ChiPhiNhaTamTruocThue = ntTruocThue;
                    }

                    var itemXD = _model.Items.FirstOrDefault(x => x.MaChiPhi == "G_XD");
                    if (itemXD != null)
                    {
                        itemXD.GiaTriTruocThue = gXD;
                        itemXD.ThueSuatGTGT = vatRate;
                    }

                    var itemNT = _model.Items.FirstOrDefault(x => x.MaChiPhi == "G_NHA_TAM");
                    if (itemNT != null)
                    {
                        itemNT.GiaTriTruocThue = ntTruocThue;
                        itemNT.ThueSuatGTGT = vatRate;
                    }

                    return true;
                }
            }
            catch { }
            return false;
        }

        private void InitializeComponent()
        {
            this.Text = "Hệ thống Quản lý Tổng hợp Dự toán & Tổng mức đầu tư xây dựng (Thông tư 36/2026/TT-BXD, TT 38/2026 & Bộ Tài chính)";
            this.Size = new Size(1300, 850);
            this.MinimumSize = new Size(1100, 700);
            this.WindowState = FormWindowState.Maximized; // Yêu cầu 1: Mở full 100% màn hình
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimizeBox = true;
            this.MaximizeBox = true;
            this.ShowInTaskbar = true;
            this.Font = UIHelper.GetFont(10.5f); // Yêu cầu 4: Font Be Vietnam Pro chuẩn

            // =========================================================================
            // 1. TOP CONTAINER: THIẾT KẾ BỐ CỤC CÂN ĐỐI, KHÔNG CHE KHUẤT CHỮ (Yêu cầu 3)
            // =========================================================================
            var topPanel = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(12, 10, 12, 8),
                BackColor = Color.FromArgb(248, 250, 253)
            };

            // Khung chế độ bảng tính (Yêu cầu 2: Cho người dùng chọn Bảng 2.1 THDT hay Bảng 1.2 TMĐT)
            var grpCheDo = new GroupBox
            {
                Text = "  1. CHỌN CHẾ ĐỘ BẢNG TÍNH THEO THÔNG TƯ 36/2026/TT-BXD  ",
                Dock = DockStyle.Top,
                Height = 68,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                Padding = new Padding(10, 6, 10, 6)
            };

            var pnlRadioCheDo = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };

            radBangTHDT = new RadioButton
            {
                Text = "Bảng 2.1: Tổng hợp Dự toán công trình (Không gồm bồi thường GPMB, Dự phòng chi phí 5%)",
                AutoSize = true,
                Checked = true,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204),
                Margin = new Padding(10, 6, 30, 0),
                Cursor = Cursors.Hand
            };
            radBangTHDT.CheckedChanged += (s, e) => { if (radBangTHDT.Checked) ChuyenCheDoBangTinh(); };

            radBangTMDT = new RadioButton
            {
                Text = "Bảng 1.2: Tổng mức đầu tư xây dựng (Bao gồm bồi thường GPMB, Dự phòng chi phí 10%)",
                AutoSize = true,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(153, 51, 0),
                Margin = new Padding(10, 6, 10, 0),
                Cursor = Cursors.Hand
            };
            radBangTMDT.CheckedChanged += (s, e) => { if (radBangTMDT.Checked) ChuyenCheDoBangTinh(); };

            pnlRadioCheDo.Controls.Add(radBangTHDT);
            pnlRadioCheDo.Controls.Add(radBangTMDT);
            grpCheDo.Controls.Add(pnlRadioCheDo);

            // Khung thông số công trình & giá trị đầu vào (Bố cục 2 dòng chuẩn, cân đối, không lệch dòng)
            var grpThongSo = new GroupBox
            {
                Text = "  2. THÔNG SỐ CÔNG TRÌNH, GIÁ TRỊ ĐẦU VÀO & THUẾ SUẤT VAT TOÀN BẢNG  ",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                Padding = new Padding(12, 10, 12, 12),
                Margin = new Padding(0, 6, 0, 4)
            };

            var tblInputs = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 8,
                RowCount = 2,
                Padding = new Padding(4, 6, 4, 6)
            };

            // Thiết lập tỷ lệ cột thông thoáng
            tblInputs.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Nhãn 1
            tblInputs.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25)); // Ô nhập 1
            tblInputs.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Nhãn 2
            tblInputs.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25)); // Ô nhập 2
            tblInputs.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Nhãn 3
            tblInputs.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25)); // Ô nhập 3
            tblInputs.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Nhãn 4
            tblInputs.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25)); // Ô nhập 4

            tblInputs.RowStyles.Add(new RowStyle(SizeType.Absolute, 42f));
            tblInputs.RowStyles.Add(new RowStyle(SizeType.Absolute, 42f));

            // Hàng 1: Loại CT, Cấp CT, Bước TK, Thuế VAT chung
            var lblLoai = new Label { Text = "Loại công trình:", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = UIHelper.GetFont(10.5f, FontStyle.Regular), ForeColor = Color.Black, Margin = new Padding(2, 0, 6, 0) };
            cboLoaiCongTrinh = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = UIHelper.GetFont(10.5f), Margin = new Padding(0, 6, 12, 6) };
            cboLoaiCongTrinh.Items.AddRange(new object[] { "-- Chọn loại công trình --", "Dân dụng", "Công nghiệp", "Giao thông", DinhMucTT38Database.LoaiNongNghiepMoiTruong, "Hạ tầng kỹ thuật" });
            cboLoaiCongTrinh.SelectedIndex = 0;
            cboLoaiCongTrinh.SelectedIndexChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);
            tblInputs.Controls.Add(lblLoai, 0, 0);
            tblInputs.Controls.Add(cboLoaiCongTrinh, 1, 0);

            var lblCap = new Label { Text = "Cấp công trình:", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = UIHelper.GetFont(10.5f, FontStyle.Regular), ForeColor = Color.Black, Margin = new Padding(2, 0, 6, 0) };
            cboCapCongTrinh = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = UIHelper.GetFont(10.5f), Margin = new Padding(0, 6, 12, 6) };
            cboCapCongTrinh.Items.AddRange(new object[] { "-- Chọn cấp công trình --", "Cấp đặc biệt", "Cấp I", "Cấp II", "Cấp III", "Cấp IV" });
            cboCapCongTrinh.SelectedIndex = 0;
            cboCapCongTrinh.SelectedIndexChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);
            tblInputs.Controls.Add(lblCap, 2, 0);
            tblInputs.Controls.Add(cboCapCongTrinh, 3, 0);

            var lblBuoc = new Label { Text = "Bước thiết kế:", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = UIHelper.GetFont(10.5f, FontStyle.Regular), ForeColor = Color.Black, Margin = new Padding(2, 0, 6, 0) };
            cboSoBuocThietKe = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = UIHelper.GetFont(10.5f), Margin = new Padding(0, 6, 12, 6) };
            cboSoBuocThietKe.Items.AddRange(new object[] { "-- Chọn bước thiết kế --", "1 bước (Báo cáo KT-KT)", "2 bước (TKBVTC)", "3 bước (TKKT & TKBVTC)" });
            cboSoBuocThietKe.SelectedIndex = 0;
            cboSoBuocThietKe.SelectedIndexChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);
            tblInputs.Controls.Add(lblBuoc, 4, 0);
            tblInputs.Controls.Add(cboSoBuocThietKe, 5, 0);

            var lblVAT = new Label { Text = "Thuế VAT chung:", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = UIHelper.GetFont(10.5f, FontStyle.Bold), ForeColor = Color.FromArgb(0, 51, 102), Margin = new Padding(2, 0, 6, 0) };
            var pnlVAT = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Margin = new Padding(0, 4, 0, 4) };
            cboVATChung = new ComboBox { Width = 95, DropDownStyle = ComboBoxStyle.DropDown, Font = UIHelper.GetFont(10.5f, FontStyle.Bold), Margin = new Padding(0, 2, 0, 0) };
            cboVATChung.Items.AddRange(new object[] { "10%", "8%", "5%", "0%" });
            cboVATChung.SelectedIndex = 0;
            cboVATChung.SelectedIndexChanged += (s, e) => XuLyDoiVATChung();
            cboVATChung.Leave += (s, e) => XuLyDoiVATChung();
            cboVATChung.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    XuLyDoiVATChung();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            };

            pnlVAT.Controls.Add(cboVATChung);
            tblInputs.Controls.Add(lblVAT, 6, 0);
            tblInputs.Controls.Add(pnlVAT, 7, 0);

            // Hàng 2: Chi phí Xây dựng, Chi phí Nhà tạm, Thiết bị, Bồi thường (Live formatting)
            var lblXD = new Label { Text = "Chi phí XD Gxd (đ):", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = UIHelper.GetFont(10.5f, FontStyle.Regular), ForeColor = Color.Black, Margin = new Padding(2, 0, 6, 0) };
            txtChiPhiXD = new TextBox { Dock = DockStyle.Fill, Font = UIHelper.GetFont(10.5f), TextAlign = HorizontalAlignment.Right, Margin = new Padding(0, 6, 12, 6) };
            txtChiPhiXD.LostFocus += (s, e) => CapNhatGiaTriDauVao();
            txtChiPhiXD.TextChanged += (s, e) => FormatLiveCurrency(txtChiPhiXD);
            txtChiPhiXD.KeyPress += OnMoneyKeyPress;
            txtChiPhiXD.KeyDown += OnMoneyKeyDown;
            tblInputs.Controls.Add(lblXD, 0, 1);
            tblInputs.Controls.Add(txtChiPhiXD, 1, 1);

            var lblNT = new Label { Text = "Chi phí nhà tạm Gnt (đ):", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = UIHelper.GetFont(10.5f, FontStyle.Regular), ForeColor = Color.Black, Margin = new Padding(2, 0, 6, 0) };
            txtChiPhiNT = new TextBox { Dock = DockStyle.Fill, Font = UIHelper.GetFont(10.5f), TextAlign = HorizontalAlignment.Right, Margin = new Padding(0, 6, 12, 6) };
            txtChiPhiNT.LostFocus += (s, e) => CapNhatGiaTriDauVao();
            txtChiPhiNT.TextChanged += (s, e) => FormatLiveCurrency(txtChiPhiNT);
            txtChiPhiNT.KeyPress += OnMoneyKeyPress;
            txtChiPhiNT.KeyDown += OnMoneyKeyDown;
            tblInputs.Controls.Add(lblNT, 2, 1);
            tblInputs.Controls.Add(txtChiPhiNT, 3, 1);

            var lblTB = new Label { Text = "Chi phí TB Gtb (đ):", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = UIHelper.GetFont(10.5f, FontStyle.Regular), ForeColor = Color.Black, Margin = new Padding(2, 0, 6, 0) };
            txtChiPhiTB = new TextBox { Dock = DockStyle.Fill, Font = UIHelper.GetFont(10.5f), TextAlign = HorizontalAlignment.Right, Margin = new Padding(0, 6, 12, 6) };
            txtChiPhiTB.LostFocus += (s, e) => CapNhatGiaTriDauVao();
            txtChiPhiTB.TextChanged += (s, e) => FormatLiveCurrency(txtChiPhiTB);
            txtChiPhiTB.KeyPress += OnMoneyKeyPress;
            txtChiPhiTB.KeyDown += OnMoneyKeyDown;
            tblInputs.Controls.Add(lblTB, 4, 1);
            tblInputs.Controls.Add(txtChiPhiTB, 5, 1);

            var lblBT = new Label { Text = "Bồi thường Gbt (đ):", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = UIHelper.GetFont(10.5f, FontStyle.Regular), ForeColor = Color.Black, Margin = new Padding(2, 0, 6, 0) };
            txtChiPhiBT = new TextBox { Dock = DockStyle.Fill, Font = UIHelper.GetFont(10.5f), TextAlign = HorizontalAlignment.Right, Margin = new Padding(0, 6, 0, 6) };
            txtChiPhiBT.LostFocus += (s, e) => CapNhatGiaTriDauVao();
            txtChiPhiBT.TextChanged += (s, e) => FormatLiveCurrency(txtChiPhiBT);
            txtChiPhiBT.KeyPress += OnMoneyKeyPress;
            txtChiPhiBT.KeyDown += OnMoneyKeyDown;
            tblInputs.Controls.Add(lblBT, 6, 1);
            tblInputs.Controls.Add(txtChiPhiBT, 7, 1);

            grpThongSo.Controls.Add(tblInputs);

            // Khung điều kiện áp dụng hệ số (Yêu cầu 3: Lưới 4 cột x 2 hàng, không che chữ, không chèn ép)
            var grpDieuKien = new GroupBox
            {
                Text = "  3. CÁC ĐIỀU KIỆN ĐẶC THÙ ÁP DỤNG HỆ SỐ ĐIỀU CHỈNH (THÔNG TƯ 38/2026/TT-BXD & BỘ TÀI CHÍNH)  ",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                Padding = new Padding(10, 4, 10, 6),
                Margin = new Padding(0, 4, 0, 0)
            };

            var tblCheckBoxes = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 4,
                RowCount = 2,
                Padding = new Padding(2)
            };
            tblCheckBoxes.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 28));
            tblCheckBoxes.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 29));
            tblCheckBoxes.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 21));
            tblCheckBoxes.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 22));

            tblCheckBoxes.RowStyles.Add(new RowStyle(SizeType.Absolute, 30f));
            tblCheckBoxes.RowStyles.Add(new RowStyle(SizeType.Absolute, 30f));

            chkThietBi50 = new CheckBox { Text = "Thiết bị ≥ 50% (k=0,7 KT, QT)", Dock = DockStyle.Fill, AutoSize = true, Font = UIHelper.GetFont(10f), Margin = new Padding(2) };
            chkThietBi50.CheckedChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            chkDaKiemToan = new CheckBox { Text = "Đã kiểm toán độc lập/KTNN (k=0,5 QT)", Dock = DockStyle.Fill, AutoSize = true, Font = UIHelper.GetFont(10f), Margin = new Padding(2) };
            chkDaKiemToan.CheckedChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            chkThueThamTra = new CheckBox { Text = "Yêu cầu thuê thẩm tra (k=0,5 TĐ)", Dock = DockStyle.Fill, AutoSize = true, Font = UIHelper.GetFont(10f), Margin = new Padding(2) };
            chkThueThamTra.CheckedChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            chkCdtTuQL = new CheckBox { Text = "Chủ đầu tư tự QLDA (k=0,8)", Dock = DockStyle.Fill, AutoSize = true, Font = UIHelper.GetFont(10f), Margin = new Padding(2) };
            chkCdtTuQL.CheckedChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            chkVungKhoKhan = new CheckBox { Text = "Vùng sâu xa, hải đảo (k=1,35)", Dock = DockStyle.Fill, AutoSize = true, Font = UIHelper.GetFont(10f), Margin = new Padding(2) };
            chkVungKhoKhan.CheckedChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            chkTuyenNhieuTinh = new CheckBox { Text = "Tuyến qua nhiều tỉnh (k=1,1)", Dock = DockStyle.Fill, AutoSize = true, Font = UIHelper.GetFont(10f), Margin = new Padding(2) };
            chkTuyenNhieuTinh.CheckedChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            chkCaiTao = new CheckBox { Text = "Cải tạo, sửa chữa (k=1,15 TK)", Dock = DockStyle.Fill, AutoSize = true, Font = UIHelper.GetFont(10f), Margin = new Padding(2) };
            chkCaiTao.CheckedChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            chkLapLai = new CheckBox { Text = "Thiết kế lặp lại (k=0,36)", Dock = DockStyle.Fill, AutoSize = true, Font = UIHelper.GetFont(10f), Margin = new Padding(2) };
            chkLapLai.CheckedChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            tblCheckBoxes.Controls.Add(chkThietBi50, 0, 0);
            tblCheckBoxes.Controls.Add(chkDaKiemToan, 1, 0);
            tblCheckBoxes.Controls.Add(chkThueThamTra, 2, 0);
            tblCheckBoxes.Controls.Add(chkCdtTuQL, 3, 0);

            tblCheckBoxes.Controls.Add(chkVungKhoKhan, 0, 1);
            tblCheckBoxes.Controls.Add(chkTuyenNhieuTinh, 1, 1);
            tblCheckBoxes.Controls.Add(chkCaiTao, 2, 1);
            tblCheckBoxes.Controls.Add(chkLapLai, 3, 1);

            grpDieuKien.Controls.Add(tblCheckBoxes);

            topPanel.Controls.Add(grpDieuKien);
            topPanel.Controls.Add(grpThongSo);
            topPanel.Controls.Add(grpCheDo);

            // =========================================================================
            // 2. TOOLBAR: NÚT THAO TÁC ĐỒNG BỘ CHIỀU CAO 36px, TỰ CO GIÃN ĐỘ RỘNG
            // =========================================================================
            var toolPanel = new Panel { Dock = DockStyle.Top, Height = 48, Padding = new Padding(12, 6, 12, 6), BackColor = Color.FromArgb(238, 242, 248) };
            var pnlToolButtons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };

            btnThemThuVien = new Button
            {
                Text = "➕ Thêm từ Thư viện...",
                AutoSize = true,
                MinimumSize = new Size(190, 36),
                Height = 36,
                Padding = new Padding(12, 0, 12, 0),
                BackColor = Color.FromArgb(0, 102, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f, FontStyle.Bold),
                Margin = new Padding(0, 0, 10, 0),
                Cursor = Cursors.Hand
            };
            btnThemThuVien.Click += BtnThemThuVien_Click;

            btnThemTuyBien = new Button
            {
                Text = "➕ Thêm dòng mới",
                AutoSize = true,
                MinimumSize = new Size(160, 36),
                Height = 36,
                Padding = new Padding(12, 0, 12, 0),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f),
                Margin = new Padding(0, 0, 10, 0),
                Cursor = Cursors.Hand
            };
            btnThemTuyBien.Click += BtnThemTuyBien_Click;

            btnXoaChiPhi = new Button
            {
                Text = "❌ Xóa dòng chọn",
                AutoSize = true,
                MinimumSize = new Size(150, 36),
                Height = 36,
                Padding = new Padding(12, 0, 12, 0),
                BackColor = Color.White,
                ForeColor = Color.DarkRed,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f),
                Margin = new Padding(0, 0, 10, 0),
                Cursor = Cursors.Hand
            };
            btnXoaChiPhi.Click += BtnXoaChiPhi_Click;

            btnKhoiPhuc = new Button
            {
                Text = "↶ Khôi phục chuẩn",
                AutoSize = true,
                MinimumSize = new Size(160, 36),
                Height = 36,
                Padding = new Padding(12, 0, 12, 0),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f),
                Margin = new Padding(0, 0, 10, 0),
                Cursor = Cursors.Hand
            };
            btnKhoiPhuc.Click += (s, e) =>
            {
                if (MessageBox.Show("Khôi phục danh mục chi phí về cấu hình chuẩn ban đầu?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    _model.Items = DinhMucTT38Engine.KhoiTaoDanhSachKhoanMucChuan();
                    DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(_model);
                    ChuyenCheDoBangTinh();
                }
            };

            btnTraLaiDinhMuc = new Button
            {
                Text = "🔄 Tra lại định mức",
                AutoSize = true,
                MinimumSize = new Size(160, 36),
                Height = 36,
                Padding = new Padding(12, 0, 12, 0),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f),
                Margin = new Padding(0, 0, 10, 0),
                Cursor = Cursors.Hand
            };
            btnTraLaiDinhMuc.Click += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            pnlToolButtons.Controls.Add(btnThemThuVien);
            pnlToolButtons.Controls.Add(btnThemTuyBien);
            pnlToolButtons.Controls.Add(btnXoaChiPhi);
            pnlToolButtons.Controls.Add(btnKhoiPhuc);
            pnlToolButtons.Controls.Add(btnTraLaiDinhMuc);
            toolPanel.Controls.Add(pnlToolButtons);

            // =========================================================================
            // 3. BOTTOM PANEL: TỔNG KẾT & CÁC NÚT XUẤT EXCEL ĐỒNG BỘ 42px (Yêu cầu 2, 3)
            // =========================================================================
            // =========================================================================
            // 3. KHỐI TỔNG KẾT (ĐƯA VÀO TAB 2 - DOCK BOTTOM DƯỚI LƯỚI CHI PHÍ)
            // =========================================================================
            pnlSummary = new TableLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 62,
                ColumnCount = 3,
                RowCount = 2,
                BackColor = Color.FromArgb(240, 245, 252),
                Padding = new Padding(12, 4, 12, 4)
            };
            pnlSummary.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            pnlSummary.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            pnlSummary.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            pnlSummary.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            pnlSummary.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));

            lblTongTruocThue = new Label { Text = "Trước thuế: 0 đ", Dock = DockStyle.Fill, Font = UIHelper.GetFont(11f, FontStyle.Bold), ForeColor = Color.FromArgb(0, 51, 102), TextAlign = ContentAlignment.MiddleLeft };
            lblTongGTGT = new Label { Text = "Thuế GTGT: 0 đ", Dock = DockStyle.Fill, Font = UIHelper.GetFont(11f, FontStyle.Bold), ForeColor = Color.FromArgb(0, 51, 102), TextAlign = ContentAlignment.MiddleLeft };

            var pnlGrandTotal = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = false };
            lblTongSauThue = new Label { Text = "0 đ", AutoSize = true, Font = UIHelper.GetFont(13.5f, FontStyle.Bold), ForeColor = Color.DarkRed, Margin = new Padding(0, 2, 0, 0) };
            lblTieuDeTongSauThue = new Label { Text = "TỔNG DỰ TOÁN CT:", AutoSize = true, Font = UIHelper.GetFont(12.5f, FontStyle.Bold), ForeColor = Color.DarkRed, Margin = new Padding(0, 4, 8, 0) };
            pnlGrandTotal.Controls.Add(lblTongSauThue);
            pnlGrandTotal.Controls.Add(lblTieuDeTongSauThue);

            lblBangChu = new Label
            {
                Text = "Bằng chữ: Không đồng.",
                Dock = DockStyle.Fill,
                Font = UIHelper.GetFont(9.5f, FontStyle.Italic),
                ForeColor = Color.FromArgb(70, 80, 95),
                TextAlign = ContentAlignment.MiddleRight
            };

            pnlSummary.Controls.Add(lblTongTruocThue, 0, 0);
            pnlSummary.Controls.Add(lblTongGTGT, 1, 0);
            pnlSummary.Controls.Add(pnlGrandTotal, 2, 0);
            pnlSummary.Controls.Add(lblBangChu, 0, 1);
            pnlSummary.SetColumnSpan(lblBangChu, 3);

            // =========================================================================
            // 4. BOTTOM PANEL: THANH TÁC VỤ CHUNG Ở ĐÁY MODAL (GỌN GÀNG, KHÔNG CHECKBOX)
            // =========================================================================
            var bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 62,
                Padding = new Padding(16, 11, 16, 11),
                BackColor = Color.FromArgb(245, 248, 252)
            };

            var lblBottomTip = new Label
            {
                Text = "💡 Bấm 'Xuất Excel theo lựa chọn' để chọn các sheet cần xuất sang Workbook hiện hành.",
                AutoSize = true,
                Dock = DockStyle.Left,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = UIHelper.GetFont(9.5f, FontStyle.Italic),
                ForeColor = Color.FromArgb(100, 116, 139)
            };

            var pnlActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                WrapContents = false
            };

            btnDong = new Button
            {
                Text = "Đóng",
                Size = new Size(100, 40),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f),
                Margin = new Padding(0, 0, 10, 0),
                Cursor = Cursors.Hand,
                UseCompatibleTextRendering = true
            };
            btnDong.FlatAppearance.BorderColor = Color.FromArgb(209, 213, 219);
            btnDong.Click += (s, e) => this.Close();

            btnLuu = new Button
            {
                Text = "💾 Lưu cấu hình",
                Size = new Size(155, 40),
                BackColor = Color.FromArgb(43, 87, 154),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f, FontStyle.Bold),
                Margin = new Padding(0, 0, 10, 0),
                Cursor = Cursors.Hand,
                UseCompatibleTextRendering = true
            };
            btnLuu.FlatAppearance.BorderSize = 0;
            btnLuu.Click += BtnLuu_Click;

            btnXuatExcelChinh = new Button
            {
                Text = "📥 Xuất Excel theo lựa chọn",
                Size = new Size(260, 40),
                BackColor = Color.FromArgb(33, 115, 70),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                Margin = new Padding(0),
                Cursor = Cursors.Hand,
                UseCompatibleTextRendering = true
            };
            btnXuatExcelChinh.FlatAppearance.BorderSize = 0;
            btnXuatExcelChinh.Click += (s, e) => XuatExcelTongHop();

            pnlActions.Controls.Add(btnDong);
            pnlActions.Controls.Add(btnLuu);
            pnlActions.Controls.Add(btnXuatExcelChinh);

            bottomPanel.Controls.Add(pnlActions);
            bottomPanel.Controls.Add(lblBottomTip);

            // =========================================================================
            // 4. CENTER: DATAGRIDVIEW CO GIÃN TỰ ĐỘNG 100% (Yêu cầu 1, 3, 4, 6, 7)
            // =========================================================================
            dgvChiPhi = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, // Tự động giãn cột 100%
                Font = UIHelper.GetFont(10.5f)
            };
            UIHelper.ApplyStyle(dgvChiPhi);
            dgvChiPhi.RowTemplate.Height = 36; // Rộng rãi, chữ to rõ ràng
            dgvChiPhi.CellValueChanged += DgvChiPhi_CellValueChanged;
            dgvChiPhi.CellDoubleClick += DgvChiPhi_CellDoubleClick;
            dgvChiPhi.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (dgvChiPhi.IsCurrentCellDirty)
                    dgvChiPhi.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            dgvChiPhi.DataError += (s, e) => { e.Cancel = true; }; // Chống crash ComboBox

            var colActive = new DataGridViewCheckBoxColumn { Name = "colActive", HeaderText = "Dùng", Width = 55, FillWeight = 4 };
            var colSTT = new DataGridViewTextBoxColumn { Name = "colSTT", HeaderText = "STT", Width = 65, FillWeight = 5, ReadOnly = true, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter } };
            var colTen = new DataGridViewTextBoxColumn { Name = "colTen", HeaderText = "Nội dung khoản mục chi phí", Width = 380, FillWeight = 36 };

            var colCachTinh = new DataGridViewComboBoxColumn
            {
                Name = "colCachTinh",
                HeaderText = "Cách tính",
                Width = 125,
                FillWeight = 11,
                DataSource = new string[] { "Theo tỷ lệ %", "Tự nhập tiền" }
            };

            var colCoSo = new DataGridViewComboBoxColumn
            {
                Name = "colCoSo",
                HeaderText = "Cơ sở tính",
                Width = 130,
                FillWeight = 11,
                DataSource = new string[] { "-", "G_XD", "G_TB", "G_XD + G_TB", "Tổng trước DP", "Toàn bộ TMĐT" }
            };

            var colTyLe = new DataGridViewTextBoxColumn { Name = "colTyLe", HeaderText = "Tỷ lệ %", Width = 80, FillWeight = 7, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight } };
            var colHeSo = new DataGridViewTextBoxColumn { Name = "colHeSo", HeaderText = "Hệ số k", Width = 70, FillWeight = 6, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight } };
            var colTruocThue = new DataGridViewTextBoxColumn { Name = "colTruocThue", HeaderText = "Trước thuế (đ)", Width = 145, FillWeight = 14, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight } };

            // Yêu cầu 7: Cột VAT chọn hoặc sửa từng dòng
            var colVAT = new DataGridViewComboBoxColumn
            {
                Name = "colVAT",
                HeaderText = "VAT",
                Width = 70,
                FillWeight = 6,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            };
            colVAT.Items.AddRange(new object[] { "10%", "8%", "5%", "0%" });

            var colSauThue = new DataGridViewTextBoxColumn { Name = "colSauThue", HeaderText = "Sau thuế (đ)", Width = 150, FillWeight = 14, ReadOnly = true, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Font = UIHelper.GetFont(10.5f, FontStyle.Bold) } };
            var colKyHieu = new DataGridViewTextBoxColumn { Name = "colKyHieu", HeaderText = "Ký hiệu", Width = 70, FillWeight = 6, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter } };

            dgvChiPhi.Columns.AddRange(colActive, colSTT, colTen, colCachTinh, colCoSo, colTyLe, colHeSo, colTruocThue, colVAT, colSauThue, colKyHieu);

            // =========================================================================
            // 5. THIẾT LẬP TAB 1: CHI PHÍ XÂY DỰNG (BẢNG 3.8 TT 36)
            // =========================================================================
            tabChiPhiXD = new TabPage
            {
                Text = "1. Chi phí Xây dựng (Bảng 3.8 TT 36/2026)",
                Padding = new Padding(10),
                BackColor = Color.FromArgb(248, 250, 253),
                Font = UIHelper.GetFont(10.5f)
            };

            var pnlLeftXD = new Panel
            {
                Dock = DockStyle.Left,
                Width = 530,
                Padding = new Padding(8, 8, 12, 8),
                AutoScroll = true
            };

            var grpThongSoXD = new GroupBox
            {
                Text = "  1. Thông số phân loại công trình (TT 36/2026)  ",
                Dock = DockStyle.Top,
                Height = 290,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                Padding = new Padding(10, 12, 10, 10)
            };

            int ty = 24;
            var lblHangMuc = new Label { Text = "Hạng mục tính toán:", Location = new Point(14, ty), Width = 175, Font = UIHelper.GetFont(10f, FontStyle.Bold), ForeColor = Color.FromArgb(0, 51, 102) };
            cboHangMucSelector = new ComboBox { Location = new Point(190, ty - 4), Width = 310, DropDownStyle = ComboBoxStyle.DropDownList, Font = UIHelper.GetFont(10f, FontStyle.Bold) };
            cboHangMucSelector.SelectedIndexChanged += CboHangMucSelector_SelectedIndexChanged;

            ty += 34;
            var lblGiaiDoanXD = new Label { Text = "Giai đoạn thực hiện:", Location = new Point(14, ty), Width = 175, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            cboGiaiDoanXD = new ComboBox { Location = new Point(190, ty - 4), Width = 310, DropDownStyle = ComboBoxStyle.DropDownList, Font = UIHelper.GetFont(10f) };
            cboGiaiDoanXD.SelectedIndexChanged += (s, e) =>
            {
                CapNhatTrangThaiQuyMoXD();
                TuDongTraTiLeXD();
            };

            ty += 35;
            var lblLoaiCTXD = new Label { Text = "Loại công trình:", Location = new Point(14, ty), Width = 175, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            cboLoaiCongTrinhXD = new ComboBox { Location = new Point(190, ty - 4), Width = 310, DropDownStyle = ComboBoxStyle.DropDownList, Font = UIHelper.GetFont(10f) };
            cboLoaiCongTrinhXD.SelectedIndexChanged += CboLoaiCongTrinhXD_SelectedIndexChanged;

            ty += 35;
            var lblPhanLoaiXD = new Label { Text = "Phân loại chi tiết:", Location = new Point(14, ty), Width = 175, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            cboPhanLoaiPhuXD = new ComboBox { Location = new Point(190, ty - 4), Width = 310, DropDownStyle = ComboBoxStyle.DropDownList, Font = UIHelper.GetFont(10f) };
            cboPhanLoaiPhuXD.SelectedIndexChanged += (s, e) => TuDongTraTiLeXD();

            ty += 35;
            var lblQuyMoXD = new Label { Text = "CP XD trong TMĐT:", Location = new Point(14, ty), Width = 175, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            txtQuyMoXD = new TextBox { Location = new Point(190, ty - 4), Width = 110, Font = UIHelper.GetFont(10f), Text = "15", TextAlign = HorizontalAlignment.Right };
            var lblDonViQuyMo = new Label { Text = "(tỷ đồng)", Location = new Point(308, ty), Width = 80, Font = UIHelper.GetFont(10f), ForeColor = Color.FromArgb(74, 85, 104) };
            txtQuyMoXD.TextChanged += (s, e) =>
            {
                if (_isUpdatingQuyMoXD) return;
                if (txtQuyMoXD.Enabled && decimal.TryParse(txtQuyMoXD.Text.Replace(',', '.'), out _))
                {
                    _lastCustomQuyMo = txtQuyMoXD.Text;
                }
                TuDongTraTiLeXD();
            };

            ty += 35;
            var lblLoaiNhaTamXD = new Label { Text = "Loại CT tính Nhà tạm:", Location = new Point(14, ty), Width = 175, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            cboLoaiNhaTamXD = new ComboBox { Location = new Point(190, ty - 4), Width = 310, DropDownStyle = ComboBoxStyle.DropDownList, Font = UIHelper.GetFont(10f) };
            cboLoaiNhaTamXD.SelectedIndexChanged += (s, e) => TuDongTraTiLeXD();

            ty += 35;
            chkVungSauXaXD = new CheckBox
            {
                Text = "Công trình tại vùng núi, biên giới, hải đảo (hệ số CPC x 1,1)",
                Location = new Point(14, ty),
                Width = 490,
                Font = UIHelper.GetFont(9.5f),
                ForeColor = Color.FromArgb(180, 50, 0)
            };
            chkVungSauXaXD.CheckedChanged += (s, e) => TuDongTraTiLeXD();

            grpThongSoXD.Controls.AddRange(new Control[] {
                lblHangMuc, cboHangMucSelector,
                lblGiaiDoanXD, cboGiaiDoanXD,
                lblLoaiCTXD, cboLoaiCongTrinhXD,
                lblPhanLoaiXD, cboPhanLoaiPhuXD,
                lblQuyMoXD, txtQuyMoXD, lblDonViQuyMo,
                lblLoaiNhaTamXD, cboLoaiNhaTamXD,
                chkVungSauXaXD
            });

            var grpTyLeXD = new GroupBox
            {
                Text = "  2. Tỷ lệ % định mức áp dụng (TT 36/2026)  ",
                Dock = DockStyle.Top,
                Height = 245,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                Padding = new Padding(10, 12, 10, 10),
                Margin = new Padding(0, 10, 0, 10)
            };

            int gy = 28;
            var lblCPC = new Label { Text = "Chi phí chung (CPC):", Location = new Point(14, gy), Width = 180, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            txtCPCXD = new TextBox { Location = new Point(200, gy - 2), Width = 85, Font = UIHelper.GetFont(10f), TextAlign = HorizontalAlignment.Right };
            var lblDonViCPC = new Label { Text = "%", Location = new Point(295, gy), Width = 35, Font = UIHelper.GetFont(10f) };

            gy += 38;
            var lblTT = new Label { Text = "Chi phí ko XĐ KL (TT):", Location = new Point(14, gy), Width = 180, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            txtTTXD = new TextBox { Location = new Point(200, gy - 2), Width = 85, Font = UIHelper.GetFont(10f), TextAlign = HorizontalAlignment.Right };
            var lblDonViTT = new Label { Text = "%", Location = new Point(295, gy), Width = 35, Font = UIHelper.GetFont(10f) };

            gy += 38;
            var lblTL = new Label { Text = "Lợi nhuận (TNCTTT):", Location = new Point(14, gy), Width = 180, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            txtTNCTTTXD = new TextBox { Location = new Point(200, gy - 2), Width = 85, Font = UIHelper.GetFont(10f), Text = "5,5", TextAlign = HorizontalAlignment.Right };
            var lblDonViTL = new Label { Text = "%", Location = new Point(295, gy), Width = 35, Font = UIHelper.GetFont(10f) };

            gy += 38;
            var lblGTGT = new Label { Text = "Thuế GTGT:", Location = new Point(14, gy), Width = 180, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            txtGTGTXD = new TextBox { Location = new Point(200, gy - 2), Width = 85, Font = UIHelper.GetFont(10f), Text = "8", TextAlign = HorizontalAlignment.Right };
            var lblDonViGTGT = new Label { Text = "%", Location = new Point(295, gy), Width = 35, Font = UIHelper.GetFont(10f) };

            gy += 38;
            var lblNhaTamXDLabel = new Label { Text = "Chi phí nhà tạm (LT):", Location = new Point(14, gy), Width = 180, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            txtNhaTamXD = new TextBox { Location = new Point(200, gy - 2), Width = 85, Font = UIHelper.GetFont(10f), Text = "1,1", TextAlign = HorizontalAlignment.Right };
            var lblDonViNT = new Label { Text = "%", Location = new Point(295, gy), Width = 35, Font = UIHelper.GetFont(10f) };

            grpTyLeXD.Controls.AddRange(new Control[] {
                lblCPC, txtCPCXD, lblDonViCPC,
                lblTT, txtTTXD, lblDonViTT,
                lblTL, txtTNCTTTXD, lblDonViTL,
                lblGTGT, txtGTGTXD, lblDonViGTGT,
                lblNhaTamXDLabel, txtNhaTamXD, lblDonViNT
            });

            _tyLeDebounce = new System.Windows.Forms.Timer { Interval = 400 };
            _tyLeDebounce.Tick += (s, e) =>
            {
                _tyLeDebounce.Stop();
                TinhToanChiPhiXD();
            };

            txtCPCXD.TextChanged += (s, e) => { _tyLeDebounce.Stop(); _tyLeDebounce.Start(); };
            txtTTXD.TextChanged += (s, e) => { _tyLeDebounce.Stop(); _tyLeDebounce.Start(); };
            txtTNCTTTXD.TextChanged += (s, e) => { _tyLeDebounce.Stop(); _tyLeDebounce.Start(); };
            txtGTGTXD.TextChanged += (s, e) => { _tyLeDebounce.Stop(); _tyLeDebounce.Start(); };
            txtNhaTamXD.TextChanged += (s, e) => { _tyLeDebounce.Stop(); _tyLeDebounce.Start(); };

            btnChuyenSangTab2 = new Button
            {
                Text = "Tiếp tục sang Tab 2: Lập TMĐT / TH Dự toán  ➜",
                Dock = DockStyle.Top,
                Height = 48,
                BackColor = Color.FromArgb(0, 102, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 14, 0, 0)
            };
            btnChuyenSangTab2.Click += (s, e) =>
            {
                TinhToanChiPhiXD();
                CapNhatGiaTriDauVao();
                tabMain.SelectedTab = tabTMDT;
            };

            pnlLeftXD.Controls.Add(btnChuyenSangTab2);
            pnlLeftXD.Controls.Add(grpTyLeXD);
            pnlLeftXD.Controls.Add(grpThongSoXD);

            var pnlRightXD = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(8, 8, 8, 8)
            };

            var lblTieuDeGridXD = new Label
            {
                Text = "BẢNG TỔNG HỢP CHI PHÍ XÂY DỰNG (BẢNG 3.8 TT 36/2026/TT-BXD)",
                Dock = DockStyle.Top,
                Height = 36,
                Font = UIHelper.GetFont(12f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                TextAlign = ContentAlignment.MiddleCenter
            };

            dgvPreviewChiPhiXD = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                Font = UIHelper.GetFont(10.5f)
            };
            UIHelper.ApplyStyle(dgvPreviewChiPhiXD);
            dgvPreviewChiPhiXD.RowTemplate.Height = 36;
            dgvPreviewChiPhiXD.DefaultCellStyle.Padding = new Padding(4, 3, 4, 3);

            dgvPreviewChiPhiXD.Columns.Add("DienGiai", "Nội dung chi phí");
            dgvPreviewChiPhiXD.Columns.Add("KyHieu", "Ký hiệu");
            dgvPreviewChiPhiXD.Columns.Add("CachTinh", "Cách tính");
            dgvPreviewChiPhiXD.Columns.Add("GiaTri", "Giá trị (đồng)");
            dgvPreviewChiPhiXD.Columns[0].FillWeight = 42;
            dgvPreviewChiPhiXD.Columns[1].FillWeight = 14;
            dgvPreviewChiPhiXD.Columns[2].FillWeight = 22;
            dgvPreviewChiPhiXD.Columns[3].FillWeight = 22;
            dgvPreviewChiPhiXD.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvPreviewChiPhiXD.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvPreviewChiPhiXD.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvPreviewChiPhiXD.Columns[3].DefaultCellStyle.Format = "N0";

            pnlRightXD.Controls.Add(dgvPreviewChiPhiXD);
            pnlRightXD.Controls.Add(lblTieuDeGridXD);

            tabChiPhiXD.Controls.Add(pnlRightXD);
            tabChiPhiXD.Controls.Add(pnlLeftXD);

            // =========================================================================
            // 6. THIẾT LẬP TAB 2: TỔNG MỨC ĐẦU TƯ & TH DỰ TOÁN (BẢNG 1.2 & 2.1)
            // =========================================================================
            tabTMDT = new TabPage
            {
                Text = "2. Tổng mức đầu tư & TH Dự toán (Bảng 1.2 & 2.1)",
                Padding = new Padding(0),
                BackColor = Color.FromArgb(248, 250, 253),
                Font = UIHelper.GetFont(10.5f)
            };

            tabTMDT.Controls.Add(dgvChiPhi);
            tabTMDT.Controls.Add(pnlSummary);
            tabTMDT.Controls.Add(toolPanel);
            tabTMDT.Controls.Add(topPanel);

            // =========================================================================
            // 7. GỘP CÁC TAB VÀO TABCONTROL CHÍNH & BOTTOM BAR
            // =========================================================================
            tabMain = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = UIHelper.GetFont(11f, FontStyle.Bold),
                ItemSize = new Size(440, 42),
                SizeMode = TabSizeMode.Fixed,
                DrawMode = TabDrawMode.OwnerDrawFixed
            };
            tabMain.DrawItem += TabMain_DrawItem;
            tabMain.SelectedIndexChanged += (s, e) =>
            {
                tabMain.Invalidate();
                if (tabMain.SelectedTab == tabTMDT)
                {
                    TinhToanChiPhiXD();
                    CapNhatGiaTriDauVao();
                }
            };
            tabMain.TabPages.Add(tabChiPhiXD);
            tabMain.TabPages.Add(tabTMDT);

            this.Controls.Add(tabMain);
            this.Controls.Add(bottomPanel);
        }

        private void NapDuLieuLenGiaoDien()
        {
            _isUpdating = true;
            try
            {
                // Nạp dữ liệu Tab 1: Chi phí Xây dựng
                NapDuLieuTabChiPhiXD();

                // Yêu cầu 2: Để trống các ô thông số đầu vào nếu chưa chọn
                if (!string.IsNullOrEmpty(_model.LoaiCongTrinh))
                {
                    cboLoaiCongTrinh.SelectedItem = _model.LoaiCongTrinh;
                }
                if (cboLoaiCongTrinh.SelectedIndex < 0) cboLoaiCongTrinh.SelectedIndex = 0;

                if (!string.IsNullOrEmpty(_model.CapCongTrinh))
                {
                    cboCapCongTrinh.SelectedItem = _model.CapCongTrinh;
                }
                if (cboCapCongTrinh.SelectedIndex < 0) cboCapCongTrinh.SelectedIndex = 0;

                if (_model.SoBuocThietKe > 0 && _model.SoBuocThietKe <= 3)
                {
                    cboSoBuocThietKe.SelectedIndex = _model.SoBuocThietKe; // Index 0 là placeholder
                }
                else
                {
                    cboSoBuocThietKe.SelectedIndex = 0;
                }

                // Format số chuẩn Việt Nam (dấu chấm hàng nghìn)
                txtChiPhiXD.Text = UIHelper.FormatTien(_model.ChiPhiXDTruocThue);
                txtChiPhiNT.Text = UIHelper.FormatTien(_model.ChiPhiNhaTamTruocThue);
                txtChiPhiTB.Text = UIHelper.FormatTien(_model.ChiPhiTBTruocThue);
                txtChiPhiBT.Text = UIHelper.FormatTien(_model.ChiPhiBTTruocThue);

                if (_model.ChiPhiTBTruocThue > 0)
                {
                    var itemGSTB = _model.Items.FirstOrDefault(x => x.MaChiPhi == "TV_GS_TB");
                    if (itemGSTB != null) itemGSTB.IsActive = true;
                    var itemTB = _model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiThietBi);
                    if (itemTB != null) itemTB.IsActive = true;
                }

                chkThietBi50.Checked = _model.ThietBiTren50Pct;
                chkDaKiemToan.Checked = _model.DaKiemToanDocLap;
                chkThueThamTra.Checked = _model.YeuCauThueThamTra;
                chkCdtTuQL.Checked = _model.CdtTuQuanLy;
                chkVungKhoKhan.Checked = _model.VungKhoKhan;
                chkTuyenNhieuTinh.Checked = _model.TuyenQuaNhieuTinh;
                chkCaiTao.Checked = _model.CaiTaoSuaChua;
                chkLapLai.Checked = _model.ThietKeLapLai;

                // Mặc định thuế VAT của Chi phí quản lý dự án là 0%
                var itemQLDA = _model.Items.FirstOrDefault(x => x.MaChiPhi == "G_QLDA" || x.Nhom == NhomChiPhi.QuanLyDuAn);
                if (itemQLDA != null && itemQLDA.ThueSuatGTGT == 0.10m)
                {
                    itemQLDA.ThueSuatGTGT = 0m;
                }

                // Mặc định thuế VAT của Chi phí bồi thường, hỗ trợ, TĐC là 0%
                var itemBT = _model.Items.FirstOrDefault(x => x.MaChiPhi == "G_BT" || x.Nhom == NhomChiPhi.BoiThuong_TDC);
                if (itemBT != null)
                {
                    itemBT.ThueSuatGTGT = 0m;
                }

                // Đồng bộ cboVATChung với thuế suất của G_XD nếu có
                var itemXD = _model.Items.FirstOrDefault(x => x.MaChiPhi == "G_XD");
                if (itemXD != null)
                {
                    string vatXDStr = $"{(itemXD.ThueSuatGTGT * 100m):G29}%";
                    if (!cboVATChung.Items.Contains(vatXDStr))
                    {
                        cboVATChung.Items.Add(vatXDStr);
                    }
                    cboVATChung.SelectedItem = vatXDStr;
                    cboVATChung.Text = vatXDStr;
                }

                ChuyenCheDoBangTinh();
            }
            finally
            {
                _isUpdating = false;
            }

            // Đảm bảo tính toán Chi phí Xây dựng và đồng bộ đầy đủ sang Tab 2 ngay từ đầu
            TuDongTraTiLeXD();
        }

        private void NapDuLieuTabChiPhiXD()
        {
            LoadGiaiDoanXD();
            LoadLoaiCongTrinhXD();
            LoadLoaiNhaTamXD();
            LoadDanhSachHangMucSelector();
        }

        private void LoadDanhSachHangMucSelector()
        {
            if (cboHangMucSelector == null) return;
            cboHangMucSelector.Items.Clear();

            if (_duToan?.DanhSachHangMuc != null && _duToan.DanhSachHangMuc.Count > 0)
            {
                if (_duToan.DanhSachHangMuc.Count > 1)
                {
                    cboHangMucSelector.Items.Add($"--- TỔNG HỢP TOÀN DỰ ÁN ({_duToan.DanhSachHangMuc.Count} hạng mục) ---");
                }

                foreach (var hm in _duToan.DanhSachHangMuc)
                {
                    string roman = LapDuToanExcelService.ToRomanNumeral(hm.STT);
                    string tenHM = hm.TenHangMuc?.Trim() ?? "";
                    string label = string.IsNullOrEmpty(roman) || tenHM.StartsWith(roman + ".", StringComparison.OrdinalIgnoreCase)
                        ? tenHM
                        : $"{roman}. {tenHM}";
                    cboHangMucSelector.Items.Add(label);
                }
            }
            else
            {
                cboHangMucSelector.Items.Add("Hạng mục mặc định");
            }

            _isChangingHangMuc = true;
            try
            {
                cboHangMucSelector.SelectedIndex = 0;
            }
            finally
            {
                _isChangingHangMuc = false;
            }

            int idx = cboHangMucSelector.SelectedIndex;
            if (_duToan.DanhSachHangMuc != null && _duToan.DanhSachHangMuc.Count > 1 && idx == 0)
            {
                _currentHangMuc = null;
                HienThiTongHopToanDuAn();
            }
            else
            {
                int hmIdx = (_duToan.DanhSachHangMuc != null && _duToan.DanhSachHangMuc.Count > 1) ? idx - 1 : idx;
                if (_duToan.DanhSachHangMuc != null && hmIdx >= 0 && hmIdx < _duToan.DanhSachHangMuc.Count)
                {
                    _currentHangMuc = _duToan.DanhSachHangMuc[hmIdx];
                    NapDuLieuChoHangMucHienHanh(_currentHangMuc);
                }
                else
                {
                    CapNhatTrangThaiQuyMoXD();
                    TuDongTraTiLeXD();
                }
            }
        }

        private void CboHangMucSelector_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_isChangingHangMuc || cboHangMucSelector.SelectedItem == null) return;
            int idx = cboHangMucSelector.SelectedIndex;

            if (_duToan.DanhSachHangMuc != null && _duToan.DanhSachHangMuc.Count > 1 && idx == 0)
            {
                _currentHangMuc = null;
                HienThiTongHopToanDuAn();
            }
            else
            {
                int hmIdx = (_duToan.DanhSachHangMuc != null && _duToan.DanhSachHangMuc.Count > 1) ? idx - 1 : idx;
                if (_duToan.DanhSachHangMuc != null && hmIdx >= 0 && hmIdx < _duToan.DanhSachHangMuc.Count)
                {
                    _currentHangMuc = _duToan.DanhSachHangMuc[hmIdx];
                    NapDuLieuChoHangMucHienHanh(_currentHangMuc);
                }
            }
        }

        private void NapDuLieuChoHangMucHienHanh(HangMuc hm)
        {
            if (hm == null) return;

            _tongVL = hm.TongVL;
            _tongNC = hm.TongNC;
            _tongMay = hm.TongMay;

            // Đồng bộ trực tiếp từ ô công thức tổng của hạng mục trên sheet DuToan nếu đang mở
            if (hm.RowIndex > 0)
            {
                try
                {
                    var app = (ExcelApp)ExcelDnaUtil.Application;
                    var wb = app?.ActiveWorkbook;
                    if (wb != null)
                    {
                        Microsoft.Office.Interop.Excel.Worksheet wsDuToan = null;
                        foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in wb.Sheets)
                        {
                            if (sheet.Name.StartsWith("DuToan", StringComparison.OrdinalIgnoreCase))
                            {
                                wsDuToan = sheet;
                                break;
                            }
                        }
                        if (wsDuToan != null)
                        {
                            object valI = wsDuToan.Cells[hm.RowIndex, 9]?.Value2;
                            object valJ = wsDuToan.Cells[hm.RowIndex, 10]?.Value2;
                            object valK = wsDuToan.Cells[hm.RowIndex, 11]?.Value2;
                            decimal vI = valI != null ? UIHelper.ParseTien(valI.ToString()) : 0m;
                            decimal vJ = valJ != null ? UIHelper.ParseTien(valJ.ToString()) : 0m;
                            decimal vK = valK != null ? UIHelper.ParseTien(valK.ToString()) : 0m;
                            if (vI > 0) _tongVL = vI;
                            if (vJ > 0) _tongNC = vJ;
                            if (vK > 0) _tongMay = vK;
                        }
                    }
                }
                catch { }
            }

            _tongT = _tongVL + _tongNC + _tongMay;

            if (!string.IsNullOrEmpty(hm.LoaiCongTrinh) && cboLoaiCongTrinhXD.Items.Contains(hm.LoaiCongTrinh))
            {
                cboLoaiCongTrinhXD.SelectedItem = hm.LoaiCongTrinh;
            }

            CapNhatTrangThaiQuyMoXD();

            if (hm.ChiPhiXD != null)
            {
                txtCPCXD.Text = hm.ChiPhiXD.TiLeCPC.ToString("0.000").Replace('.', ',');
                txtTTXD.Text = hm.ChiPhiXD.TiLeTT.ToString("0.000").Replace('.', ',');
                txtTNCTTTXD.Text = hm.ChiPhiXD.TiLeTNCTTT.ToString("0.0").Replace('.', ',');
                txtGTGTXD.Text = hm.ChiPhiXD.TiLeGTGT.ToString("0").Replace('.', ',');
                txtNhaTamXD.Text = hm.ChiPhiXD.TiLeNhaTam.ToString("0.00").Replace('.', ',');
                TinhToanChiPhiXD();
            }
            else
            {
                TuDongTraTiLeXD();
            }
        }

        private void CapNhatTongChiPhiXDToanDuAn()
        {
            if (_duToan.DanhSachHangMuc == null || _duToan.DanhSachHangMuc.Count == 0) return;

            foreach (var hm in _duToan.DanhSachHangMuc)
            {
                if (hm.ChiPhiXD == null)
                {
                    decimal defaultCPC = InterpolationHelper.LayTiLeCPCTMDT(hm.LoaiCongTrinh ?? "Dân dụng", hm.PhanLoaiPhu);
                    var ttModel = _ttRepo.GetByLoaiCongTrinh(hm.LoaiCongTrinh ?? "Dân dụng", hm.PhanLoaiPhu);
                    decimal defaultTT = ttModel?.TiLe ?? 2.0m;
                    decimal defaultTL = InterpolationHelper.LayTiLeTNCTTT(hm.LoaiCongTrinh ?? "Dân dụng");
                    decimal defaultGTGT = 10m;
                    decimal defaultNT = 1.0m;
                    hm.ChiPhiXD = _calcService.Tinh(hm.TongVL, hm.TongNC, hm.TongMay, defaultCPC, defaultTT, defaultTL, defaultGTGT, defaultNT, "T");
                }
            }

            decimal totalVL = _duToan.DanhSachHangMuc.Sum(h => h.ChiPhiXD?.VL ?? h.TongVL);
            decimal totalNC = _duToan.DanhSachHangMuc.Sum(h => h.ChiPhiXD?.NC ?? h.TongNC);
            decimal totalM = _duToan.DanhSachHangMuc.Sum(h => h.ChiPhiXD?.M ?? h.TongMay);
            decimal totalCPC = _duToan.DanhSachHangMuc.Sum(h => h.ChiPhiXD?.CPC ?? 0);
            decimal totalTT = _duToan.DanhSachHangMuc.Sum(h => h.ChiPhiXD?.TT ?? 0);
            decimal totalTL = _duToan.DanhSachHangMuc.Sum(h => h.ChiPhiXD?.TL ?? 0);
            decimal totalGTGT = _duToan.DanhSachHangMuc.Sum(h => h.ChiPhiXD?.GTGT ?? 0);
            decimal totalLT = _duToan.DanhSachHangMuc.Sum(h => h.ChiPhiXD?.LT ?? 0);

            _duToan.ChiPhiXD = new ChiPhiXayDung
            {
                VL = totalVL,
                NC = totalNC,
                M = totalM,
                CPC = totalCPC,
                TT = totalTT,
                TL = totalTL,
                GTGT = totalGTGT,
                LT = totalLT
            };

            if (_model != null)
            {
                _model.ChiPhiXDTruocThue = _duToan.ChiPhiXD.GXDTT;
                _model.ChiPhiNhaTamTruocThue = totalLT;
            }
        }

        private void HienThiTongHopToanDuAn()
        {
            CapNhatTongChiPhiXDToanDuAn();
            var kq = _duToan.ChiPhiXD;
            if (kq == null || dgvPreviewChiPhiXD == null) return;

            dgvPreviewChiPhiXD.Rows.Clear();
            
            var rTong = new DataGridViewRow();
            rTong.CreateCells(dgvPreviewChiPhiXD, "TỔNG CHI PHÍ XÂY DỰNG TOÀN DỰ ÁN", "G_XD", "Σ G_XD các hạng mục", kq.GXD);
            rTong.DefaultCellStyle.Font = new Font(dgvPreviewChiPhiXD.Font, FontStyle.Bold);
            rTong.DefaultCellStyle.BackColor = Color.FromArgb(230, 244, 234);
            dgvPreviewChiPhiXD.Rows.Add(rTong);

            dgvPreviewChiPhiXD.Rows.Add("Chi phí xây dựng trước thuế (GXDTT)", "GXDTT", "Σ GXDTT các hạng mục", kq.GXDTT);
            dgvPreviewChiPhiXD.Rows.Add("Thuế GTGT", "GTGT", "Σ GTGT các hạng mục", kq.GTGT);
            dgvPreviewChiPhiXD.Rows.Add("Chi phí nhà tạm (LT)", "LT", "Σ LT các hạng mục", kq.LT);
            dgvPreviewChiPhiXD.Rows.Add("Chi phí trực tiếp (T)", "T", "VL + NC + M", kq.T);
            dgvPreviewChiPhiXD.Rows.Add("1. Chi phí vật liệu (VL)", "VL", "", kq.VL);
            dgvPreviewChiPhiXD.Rows.Add("2. Chi phí nhân công (NC)", "NC", "", kq.NC);
            dgvPreviewChiPhiXD.Rows.Add("3. Chi phí máy thi công (M)", "M", "", kq.M);
            dgvPreviewChiPhiXD.Rows.Add("Chi phí gián tiếp (GT)", "GT", "C + TT", kq.GT);
            dgvPreviewChiPhiXD.Rows.Add("Thu nhập chịu thuế tính trước (TL)", "TL", "", kq.TL);

            var rSep = new DataGridViewRow();
            rSep.CreateCells(dgvPreviewChiPhiXD, "--- TỔNG HỢP THEO TỪNG HẠNG MỤC ---", "", "", 0);
            rSep.DefaultCellStyle.Font = new Font(dgvPreviewChiPhiXD.Font, FontStyle.Bold);
            rSep.DefaultCellStyle.ForeColor = Color.FromArgb(0, 102, 204);
            dgvPreviewChiPhiXD.Rows.Add(rSep);

            foreach (var hm in _duToan.DanhSachHangMuc)
            {
                decimal gxd = hm.ChiPhiXD?.GXD ?? 0;
                decimal t = hm.ChiPhiXD?.T ?? hm.TongChiPhi;
                string roman = LapDuToanExcelService.ToRomanNumeral(hm.STT);
                string tenHM = hm.TenHangMuc?.Trim() ?? "";
                string label = string.IsNullOrEmpty(roman) || tenHM.StartsWith(roman + ".", StringComparison.OrdinalIgnoreCase)
                    ? tenHM
                    : $"{roman}. {tenHM}";
                dgvPreviewChiPhiXD.Rows.Add(label, $"T: {t:N0}", "G_XD:", gxd);
            }

            foreach (DataGridViewRow r in dgvPreviewChiPhiXD.Rows)
            {
                r.Height = 36;
            }
            dgvPreviewChiPhiXD?.AutoFit();

            decimal.TryParse(txtGTGTXD.Text.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal gtgtTongHop);
            decimal.TryParse(txtNhaTamXD.Text.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal nhatamTongHop);
            if (gtgtTongHop <= 0) gtgtTongHop = 10m;
            if (nhatamTongHop <= 0) nhatamTongHop = 1.0m;

            DongBoSangTab2(kq, gtgtTongHop, nhatamTongHop);
        }

        private void LoadGiaiDoanXD()
        {
            cboGiaiDoanXD.DataSource = null;
            cboGiaiDoanXD.Items.Clear();
            cboGiaiDoanXD.Items.AddRange(new object[] {
                "Lập dự toán xây dựng",
                "Lập Báo cáo kinh tế - kỹ thuật",
                "Lập Báo cáo NCKT (Tổng mức đầu tư)"
            });
            cboGiaiDoanXD.SelectedIndex = 0;
        }

        private void LoadLoaiNhaTamXD()
        {
            cboLoaiNhaTamXD.DataSource = null;
            cboLoaiNhaTamXD.Items.Clear();
            cboLoaiNhaTamXD.Items.AddRange(new object[] {
                "Công trình xây dựng còn lại",
                "Công trình xây dựng theo tuyến"
            });
            cboLoaiNhaTamXD.SelectedIndex = 0;
        }


        private void TabMain_DrawItem(object sender, DrawItemEventArgs e)
        {
            var tabCtrl = (TabControl)sender;
            var page = tabCtrl.TabPages[e.Index];
            var rect = tabCtrl.GetTabRect(e.Index);
            bool isSelected = (tabCtrl.SelectedIndex == e.Index);

            using var sf = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center,
                FormatFlags = StringFormatFlags.NoWrap,
                Trimming = StringTrimming.EllipsisCharacter
            };

            if (isSelected)
            {
                // Tab active: Nền xanh dương thương hiệu nổi bật
                using var brushBg = new SolidBrush(Color.FromArgb(0, 102, 204));
                e.Graphics.FillRectangle(brushBg, rect);

                // Đường viền nhấn màu vàng cam ở chân tab
                using var brushIndicator = new SolidBrush(Color.FromArgb(255, 152, 0));
                e.Graphics.FillRectangle(brushIndicator, rect.X, rect.Bottom - 4, rect.Width, 4);

                // Chữ trắng in đậm
                using var fontBold = UIHelper.GetFont(11f, FontStyle.Bold);
                using var brushText = new SolidBrush(Color.White);
                e.Graphics.DrawString(page.Text, fontBold, brushText, rect, sf);
            }
            else
            {
                // Tab inactive: Nền xám nhạt, chữ xám đen
                using var brushBg = new SolidBrush(Color.FromArgb(238, 242, 246));
                e.Graphics.FillRectangle(brushBg, rect);

                using var penBorder = new Pen(Color.FromArgb(209, 213, 219));
                e.Graphics.DrawRectangle(penBorder, rect.X, rect.Y, rect.Width - 1, rect.Height - 1);

                using var fontNormal = UIHelper.GetFont(10.5f, FontStyle.Regular);
                using var brushText = new SolidBrush(Color.FromArgb(74, 85, 104));
                e.Graphics.DrawString(page.Text, fontNormal, brushText, rect, sf);
            }
        }

        private void LoadLoaiCongTrinhXD()
        {
            cboLoaiCongTrinhXD.DataSource = null;
            cboLoaiCongTrinhXD.Items.Clear();
            cboLoaiCongTrinhXD.Items.AddRange(new object[] { "-- Chọn loại công trình --", "Dân dụng", "Công nghiệp", "Giao thông", DinhMucTT38Database.LoaiNongNghiepMoiTruong, "Hạ tầng kỹ thuật" });

            string target = !string.IsNullOrEmpty(_duToan.LoaiCongTrinh) ? _duToan.LoaiCongTrinh : (!string.IsNullOrEmpty(_model?.LoaiCongTrinh) ? _model.LoaiCongTrinh : "Dân dụng");
            if (target == "Nông nghiệp và môi trường" || target == "Nông nghiệp & PTNT") target = DinhMucTT38Database.LoaiNongNghiepMoiTruong;

            if (!string.IsNullOrEmpty(target) && cboLoaiCongTrinhXD.Items.Contains(target))
            {
                cboLoaiCongTrinhXD.SelectedItem = target;
            }
            else
            {
                cboLoaiCongTrinhXD.SelectedIndex = 1;
            }
        }

        private void CboLoaiCongTrinhXD_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboLoaiCongTrinhXD.SelectedItem == null) return;
            string loaiCT = cboLoaiCongTrinhXD.SelectedItem.ToString();

            if (cboLoaiCongTrinhXD.SelectedIndex <= 0 || loaiCT.StartsWith("--"))
            {
                if (!_isSyncingLoaiCT && cboLoaiCongTrinh != null)
                {
                    try
                    {
                        _isSyncingLoaiCT = true;
                        if (cboLoaiCongTrinh.SelectedIndex != 0)
                            cboLoaiCongTrinh.SelectedIndex = 0;
                    }
                    finally { _isSyncingLoaiCT = false; }
                }
                _duToan.LoaiCongTrinh = "";
                if (_model != null) _model.LoaiCongTrinh = "";

                cboPhanLoaiPhuXD.Items.Clear();
                cboPhanLoaiPhuXD.Items.Add("--- Mặc định ---");
                cboPhanLoaiPhuXD.SelectedIndex = 0;
                cboPhanLoaiPhuXD.Enabled = false;
                return;
            }

            // Đồng bộ sang cboLoaiCongTrinh của Tab 2
            if (!_isSyncingLoaiCT && cboLoaiCongTrinh != null && cboLoaiCongTrinh.Items.Contains(loaiCT))
            {
                try
                {
                    _isSyncingLoaiCT = true;
                    if (cboLoaiCongTrinh.SelectedItem?.ToString() != loaiCT)
                    {
                        cboLoaiCongTrinh.SelectedItem = loaiCT;
                    }
                }
                finally { _isSyncingLoaiCT = false; }
            }
            _duToan.LoaiCongTrinh = loaiCT;
            if (_model != null) _model.LoaiCongTrinh = loaiCT;

            var allData = _cpcRepo.GetAll().Where(x => 
                x.LoaiCongTrinh == loaiCT || 
                (loaiCT == DinhMucTT38Database.LoaiNongNghiepMoiTruong && (x.LoaiCongTrinh == "Nông nghiệp và môi trường" || x.LoaiCongTrinh == "Nông nghiệp & PTNT"))
            ).ToList();

            var phanLoai = allData.Where(x => !string.IsNullOrEmpty(x.PhanLoaiPhu))
                                  .Select(x => x.PhanLoaiPhu)
                                  .Distinct()
                                  .ToList();

            cboPhanLoaiPhuXD.Items.Clear();
            if (phanLoai.Count == 0)
            {
                cboPhanLoaiPhuXD.Items.Add("--- Không có ---");
                cboPhanLoaiPhuXD.Enabled = false;
            }
            else
            {
                cboPhanLoaiPhuXD.Enabled = true;
                cboPhanLoaiPhuXD.Items.Add("--- Mặc định ---");
                foreach (var item in phanLoai)
                    cboPhanLoaiPhuXD.Items.Add(item);
            }
            cboPhanLoaiPhuXD.SelectedIndex = 0;

            // Tra Thu nhập chịu thuế tính trước theo Bảng 3.6
            decimal defTNCTTT = InterpolationHelper.LayTiLeTNCTTT(loaiCT);
            txtTNCTTTXD.Text = defTNCTTT.ToString("0.0").Replace('.', ',');

            TuDongTraTiLeXD();
        }

        private void CapNhatTrangThaiQuyMoXD()
        {
            if (cboGiaiDoanXD == null || txtQuyMoXD == null) return;
            string giaiDoan = cboGiaiDoanXD.SelectedItem?.ToString() ?? "Lập dự toán xây dựng";

            if (giaiDoan.IndexOf("kinh tế - kỹ thuật", StringComparison.OrdinalIgnoreCase) >= 0 ||
                giaiDoan.IndexOf("KT-KT", StringComparison.OrdinalIgnoreCase) >= 0 ||
                giaiDoan.IndexOf("KTKT", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                if (decimal.TryParse(txtQuyMoXD.Text.Replace(',', '.'), out _))
                {
                    _lastCustomQuyMo = txtQuyMoXD.Text;
                }
                txtQuyMoXD.Text = "≤ 40";
                txtQuyMoXD.Enabled = false;
                txtQuyMoXD.BackColor = Color.FromArgb(240, 243, 246);
                txtQuyMoXD.ForeColor = Color.FromArgb(0, 102, 204);
            }
            else if (giaiDoan.IndexOf("Tổng mức đầu tư", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     giaiDoan.IndexOf("NCKT", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                if (decimal.TryParse(txtQuyMoXD.Text.Replace(',', '.'), out _))
                {
                    _lastCustomQuyMo = txtQuyMoXD.Text;
                }
                txtQuyMoXD.Text = "Bảng 3.2";
                txtQuyMoXD.Enabled = false;
                txtQuyMoXD.BackColor = Color.FromArgb(240, 243, 246);
                txtQuyMoXD.ForeColor = Color.FromArgb(120, 120, 120);
            }
            else
            {
                txtQuyMoXD.Enabled = true;
                txtQuyMoXD.BackColor = Color.White;
                txtQuyMoXD.ForeColor = Color.Black;
                if (txtQuyMoXD.Text.Contains("≤") || txtQuyMoXD.Text.Contains("Bảng") || !decimal.TryParse(txtQuyMoXD.Text.Replace(',', '.'), out _))
                {
                    // Giai đoạn lập dự toán: ưu tiên quy mô thực tế từ dự toán
                    decimal quyMoThucTe = TinhQuyMoTuChiPhiXD();
                    if (quyMoThucTe <= 0 &&
                        decimal.TryParse((_lastCustomQuyMo ?? "").Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal lastQuyMo) &&
                        lastQuyMo > 0)
                    {
                        quyMoThucTe = lastQuyMo;
                    }
                    if (quyMoThucTe <= 0) quyMoThucTe = 15m;
                    HienThiQuyMoXD(quyMoThucTe);
                }
            }
        }

        /// <summary>
        /// Xác định quy mô chi phí xây dựng trước thuế trong TMĐT (tỷ đồng) từ dự toán thực tế.
        /// Dùng làm cơ sở tra Bảng 3.3 (CPC) và Bảng 3.7 (nhà tạm) theo TT 36/2026.
        /// </summary>
        private decimal TinhQuyMoTuChiPhiXD()
        {
            const decimal TY = 1_000_000_000m;

            decimal gxdtt = _duToan?.ChiPhiXD?.GXDTT ?? 0m;
            if (gxdtt <= 0)
                gxdtt = _currentHangMuc?.ChiPhiXD?.GXDTT ?? 0m;

            if (gxdtt <= 0 && _tongT > 0)
            {
                // Ước lượng GXDTT từ chi phí trực tiếp với tỷ lệ đại diện
                // GXDTT = T × (1 + (CPC + TT)/100) × (1 + TNCTTT/100)
                decimal cpc = 7.3m, tt = 2.5m, tncttt = 5.5m;
                gxdtt = _tongT * (1m + (cpc + tt) / 100m) * (1m + tncttt / 100m);
            }

            if (gxdtt <= 0) return 0m;
            return Math.Round(gxdtt / TY, 4, MidpointRounding.AwayFromZero);
        }

        private void HienThiQuyMoXD(decimal quyMo)
        {
            if (txtQuyMoXD == null || quyMo <= 0) return;
            string text = quyMo.ToString("0.####").Replace('.', ',');
            if (txtQuyMoXD.Text == text) return;
            _isUpdatingQuyMoXD = true;
            try
            {
                txtQuyMoXD.Text = text;
            }
            finally
            {
                _isUpdatingQuyMoXD = false;
            }
        }

        private void TuDongTraTiLeXD()
        {
            if (cboLoaiCongTrinhXD == null || cboLoaiCongTrinhXD.SelectedItem == null) return;
            string loaiCT = cboLoaiCongTrinhXD.SelectedItem.ToString();
            if (cboLoaiCongTrinhXD.SelectedIndex <= 0 || loaiCT.StartsWith("--")) return;
            string phanLoai = cboPhanLoaiPhuXD.SelectedIndex > 0 ? cboPhanLoaiPhuXD.SelectedItem.ToString() : null;

            string giaiDoan = cboGiaiDoanXD?.SelectedItem?.ToString() ?? "Lập dự toán xây dựng";
            string loaiNhaTam = cboLoaiNhaTamXD?.SelectedItem?.ToString() ?? "Công trình xây dựng còn lại";
            bool laVungSauXa = chkVungSauXaXD?.Checked ?? false;

            decimal quyMo = 15m;
            bool laLapDuToan = !(giaiDoan.IndexOf("Tổng mức đầu tư", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 giaiDoan.IndexOf("NCKT", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 giaiDoan.IndexOf("kinh tế - kỹ thuật", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 giaiDoan.IndexOf("KT-KT", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 giaiDoan.IndexOf("KTKT", StringComparison.OrdinalIgnoreCase) >= 0);

            if (laLapDuToan)
            {
                // Giai đoạn lập dự toán xây dựng: quy mô lấy trực tiếp từ
                // chi phí xây dựng trước thuế trong TMĐT của dự toán (tỷ đồng)
                quyMo = TinhQuyMoTuChiPhiXD();
                if (quyMo <= 0) quyMo = 15m;
                HienThiQuyMoXD(quyMo);
            }
            else
            {
                string rawQuyMo = (txtQuyMoXD?.Text ?? "").Replace("≤", "").Replace("<=", "").Replace("tỷ", "").Trim();
                if (decimal.TryParse(rawQuyMo.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal parsedQuyMo) && parsedQuyMo > 0)
                {
                    quyMo = parsedQuyMo;
                }
                else if (giaiDoan.IndexOf("kinh tế - kỹ thuật", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         giaiDoan.IndexOf("KT-KT", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         giaiDoan.IndexOf("KTKT", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    quyMo = 40m;
                }
            }

            decimal cpc = 0m;

            if (giaiDoan.IndexOf("Tổng mức đầu tư", StringComparison.OrdinalIgnoreCase) >= 0 ||
                giaiDoan.IndexOf("NCKT", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Bảng 3.2 TT 36: Giai đoạn lập TMĐT (tỷ lệ cố định theo loại CT)
                cpc = InterpolationHelper.LayTiLeCPCTMDT(loaiCT, phanLoai);
            }
            else if (giaiDoan.IndexOf("kinh tế - kỹ thuật", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     giaiDoan.IndexOf("KT-KT", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     giaiDoan.IndexOf("KTKT", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Điểm đ khoản 3.1.2 TT 36: Dự án chỉ lập BC KT-KT thì lấy theo cột [3] Bảng 3.3 (mốc <= 40 tỷ)
                var listCPC = _cpcRepo.GetByLoaiCongTrinh(loaiCT, phanLoai);
                if (listCPC.Count == 0 && phanLoai != null)
                    listCPC = _cpcRepo.GetByLoaiCongTrinh(loaiCT, null);

                if (listCPC.Count > 0)
                {
                    cpc = listCPC.OrderBy(x => x.QuyMoMin).First().TiLe;
                }
                else
                {
                    cpc = InterpolationHelper.LayTiLeCPCTMDT(loaiCT, phanLoai);
                }
            }
            else
            {
                // Khoản 3.1.2.a TT 36: Giai đoạn lập dự toán xây dựng (nội suy Bảng 3.3)
                var listCPC = _cpcRepo.GetByLoaiCongTrinh(loaiCT, phanLoai);
                if (listCPC.Count == 0 && phanLoai != null)
                    listCPC = _cpcRepo.GetByLoaiCongTrinh(loaiCT, null);

                if (listCPC.Count > 0)
                {
                    cpc = InterpolationHelper.NoiSuyTiLeCPC(listCPC, quyMo);
                }
                else
                {
                    cpc = InterpolationHelper.LayTiLeCPCTMDT(loaiCT, phanLoai);
                }
            }

            // Khoản 3.1.4: Công trình tại vùng núi, biên giới, hải đảo x 1.1
            if (laVungSauXa && cpc > 0)
            {
                cpc = Math.Round(cpc * 1.1m, 3, MidpointRounding.AwayFromZero);
            }

            if (cpc > 0)
            {
                txtCPCXD.Text = cpc.ToString("0.000").Replace('.', ',');
            }

            // Tra TT (Bảng 3.5)
            var tt = _ttRepo.GetByLoaiCongTrinh(loaiCT, phanLoai);
            if (tt == null && phanLoai != null)
                tt = _ttRepo.GetByLoaiCongTrinh(loaiCT, null);

            if (tt != null)
            {
                txtTTXD.Text = tt.TiLe.ToString("0.000").Replace('.', ',');
            }

            // Tra Thu nhập chịu thuế tính trước (Bảng 3.6)
            decimal tncttt = InterpolationHelper.LayTiLeTNCTTT(loaiCT);
            txtTNCTTTXD.Text = tncttt.ToString("0.0").Replace('.', ',');

            // Tra Chi phí nhà tạm (Bảng 3.7)
            // Giai đoạn KT-KT khóa ô quy mô hiển thị "≤ 40" nên không dùng quyMo này tra Bảng 3.7
            // (40 tỷ sẽ nội suy ra 1,07%/2,14%). Phải dùng quy mô thực tế của dự toán.
            decimal quyMoNhaTam = TinhQuyMoTuChiPhiXD();
            if (quyMoNhaTam <= 0) quyMoNhaTam = quyMo;
            decimal tlNhaTam = InterpolationHelper.NoiSuyTiLeNhaTam(loaiNhaTam, quyMoNhaTam);
            txtNhaTamXD.Text = tlNhaTam.ToString("0.00").Replace('.', ',');

            // Đồng bộ giai đoạn với Tab 2 (Chế độ bảng tính và số bước thiết kế)
            if (giaiDoan.IndexOf("kinh tế - kỹ thuật", StringComparison.OrdinalIgnoreCase) >= 0 ||
                giaiDoan.IndexOf("KT-KT", StringComparison.OrdinalIgnoreCase) >= 0 ||
                giaiDoan.IndexOf("KTKT", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                if (radBangTMDT != null && !radBangTMDT.Checked) radBangTMDT.Checked = true;
                if (cboSoBuocThietKe != null && cboSoBuocThietKe.Items.Count > 1 && cboSoBuocThietKe.SelectedIndex != 1)
                {
                    cboSoBuocThietKe.SelectedIndex = 1; // 1 bước (Báo cáo KT-KT)
                }
            }
            else if (giaiDoan.IndexOf("Tổng mức đầu tư", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     giaiDoan.IndexOf("NCKT", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                if (radBangTMDT != null && !radBangTMDT.Checked) radBangTMDT.Checked = true;
            }
            else
            {
                if (radBangTHDT != null && !radBangTHDT.Checked) radBangTHDT.Checked = true;
            }

            TinhToanChiPhiXD();
        }

        private void TinhToanChiPhiXD()
        {
            if (_isUpdating) return;

            decimal.TryParse(txtCPCXD.Text.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal cpc);
            decimal.TryParse(txtTTXD.Text.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal tt);
            decimal.TryParse(txtTNCTTTXD.Text.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal tncttt);
            decimal.TryParse(txtGTGTXD.Text.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal gtgt);
            decimal.TryParse(txtNhaTamXD.Text.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal nhatam);

            decimal tongVL = _tongVL;
            decimal tongNC = _tongNC;
            decimal tongMay = _tongMay;

            if (tongVL == 0 && tongNC == 0 && tongMay == 0)
            {
                tongMay = _tongT - _tongNC - (_duToan.DanhSachHangMuc?.Sum(hm => hm.DanhSachCongTac.Sum(c => c.ThanhTienVL)) ?? 0);
                tongVL = _tongT - _tongNC - tongMay;
            }

            string giaiDoan = cboGiaiDoanXD?.SelectedItem?.ToString() ?? "Lập dự toán xây dựng";
            string loaiNhaTam = cboLoaiNhaTamXD?.SelectedItem?.ToString() ?? "Công trình xây dựng còn lại";
            bool laVungSauXa = chkVungSauXaXD?.Checked ?? false;

            if (_currentHangMuc == null && _duToan.DanhSachHangMuc != null && _duToan.DanhSachHangMuc.Count > 1)
            {
                HienThiTongHopToanDuAn();
                return;
            }

            var kq = _calcService.Tinh(tongVL, tongNC, tongMay, cpc, tt, tncttt, gtgt, nhatam, "T", giaiDoan, loaiNhaTam, laVungSauXa);
            if (_currentHangMuc != null)
            {
                _currentHangMuc.ChiPhiXD = kq;
                _currentHangMuc.LoaiCongTrinh = cboLoaiCongTrinhXD?.SelectedItem?.ToString() ?? "";
            }
            else
            {
                _duToan.ChiPhiXD = kq;
            }

            CapNhatTongChiPhiXDToanDuAn();

            if (dgvPreviewChiPhiXD != null)
            {
                dgvPreviewChiPhiXD.Rows.Clear();
                dgvPreviewChiPhiXD.Rows.Add("I. Chi phí trực tiếp", "T", "VL + NC + M", kq.T);
                dgvPreviewChiPhiXD.Rows.Add("1. Chi phí vật liệu", "VL", "Σ(KL × ĐG_VL)", kq.VL);
                dgvPreviewChiPhiXD.Rows.Add("2. Chi phí nhân công", "NC", "Σ(KL × ĐG_NC × Knc)", kq.NC);
                dgvPreviewChiPhiXD.Rows.Add("3. Chi phí máy và thiết bị thi công", "M", "Σ(KL × ĐG_M × Km)", kq.M);

                // Nếu hạng mục hiện tại có các hạng mục con, hiển thị phân rã chi tiết
                if (_currentHangMuc?.DanhSachHangMucCon != null && _currentHangMuc.DanhSachHangMucCon.Count > 0)
                {
                    foreach (var hmc in _currentHangMuc.DanhSachHangMucCon)
                    {
                        var rHmc = new DataGridViewRow();
                        rHmc.CreateCells(dgvPreviewChiPhiXD, $"   ▸ {hmc.TenHangMucCon}", "T_con", $"VL: {hmc.TongVL:N0} | NC: {hmc.TongNC:N0} | M: {hmc.TongMay:N0}", hmc.TongChiPhi);
                        rHmc.DefaultCellStyle.Font = new Font(dgvPreviewChiPhiXD.Font, FontStyle.Italic);
                        rHmc.DefaultCellStyle.ForeColor = Color.FromArgb(70, 80, 95);
                        dgvPreviewChiPhiXD.Rows.Add(rHmc);
                    }
                }

                dgvPreviewChiPhiXD.Rows.Add("II. Chi phí gián tiếp", "GT", "C + TT", kq.GT);
                dgvPreviewChiPhiXD.Rows.Add($"1. Chi phí chung ({cpc}%)", "C", "T × tỷ lệ", kq.CPC);
                dgvPreviewChiPhiXD.Rows.Add($"2. CP một số CV không XĐ được KL ({tt}%)", "TT", "T × tỷ lệ", kq.TT);
                dgvPreviewChiPhiXD.Rows.Add($"III. Thu nhập chịu thuế tính trước ({tncttt}%)", "TL", "(T + GT) × tỷ lệ", kq.TL);
                dgvPreviewChiPhiXD.Rows.Add("Chi phí xây dựng trước thuế", "GXDTT", "T + GT + TL", kq.GXDTT);
                dgvPreviewChiPhiXD.Rows.Add($"IV. Thuế GTGT ({gtgt}%)", "GTGT", "GXDTT × thuế suất", kq.GTGT);

                var rSauThue = new DataGridViewRow();
                rSauThue.CreateCells(dgvPreviewChiPhiXD, "CHI PHÍ XÂY DỰNG SAU THUẾ", "GXD", "GXDTT + GTGT", kq.GXD);
                rSauThue.DefaultCellStyle.Font = new Font(dgvPreviewChiPhiXD.Font, FontStyle.Bold);
                rSauThue.DefaultCellStyle.BackColor = Color.FromArgb(230, 244, 234);
                dgvPreviewChiPhiXD.Rows.Add(rSauThue);

                var rNhaTam = new DataGridViewRow();
                rNhaTam.CreateCells(dgvPreviewChiPhiXD, $"V. Chi phí nhà tạm ({nhatam}%)", "LT", "GXDTT × tỷ lệ × (1 + TGTGT)", kq.LT);
                rNhaTam.DefaultCellStyle.Font = new Font(dgvPreviewChiPhiXD.Font, FontStyle.Bold);
                dgvPreviewChiPhiXD.Rows.Add(rNhaTam);

                foreach (DataGridViewRow r in dgvPreviewChiPhiXD.Rows)
                {
                    r.Height = 36;
                }
                dgvPreviewChiPhiXD?.AutoFit();
            }

            DongBoSangTab2(_duToan.ChiPhiXD ?? kq, gtgt, nhatam);
        }

        private void DongBoSangTab2(ChiPhiXayDung kq, decimal gtgt, decimal nhatam)
        {
            if (_model == null) return;

            if (_isSyncingVAT)
            {
                CapNhatThanhTongCong();
                return;
            }

            decimal vatRate = gtgt / 100m;
            decimal ntTruocThue = Math.Round(kq.GXDTT * nhatam / 100m, 0);

            _model.ChiPhiXDTruocThue = kq.GXDTT;
            _model.ChiPhiNhaTamTruocThue = ntTruocThue;

            if (txtChiPhiXD != null) txtChiPhiXD.Text = UIHelper.FormatTien(kq.GXDTT);
            if (txtChiPhiNT != null) txtChiPhiNT.Text = UIHelper.FormatTien(ntTruocThue);


            var colVATSync = dgvChiPhi?.Columns["colVAT"] as DataGridViewComboBoxColumn;
            if (colVATSync != null)
            {
                string vatTextCol = $"{gtgt:G29}%";
                if (!colVATSync.Items.Contains(vatTextCol))
                {
                    colVATSync.Items.Add(vatTextCol);
                }
            }

            var itemXD = _model.Items.FirstOrDefault(x => x.MaChiPhi == "G_XD");
            if (itemXD != null)
            {
                itemXD.GiaTriTruocThue = kq.G;
                itemXD.ThueSuatGTGT = vatRate;
            }

            var itemNT = _model.Items.FirstOrDefault(x => x.MaChiPhi == "G_NHA_TAM");
            if (itemNT != null)
            {
                itemNT.GiaTriTruocThue = ntTruocThue;
                itemNT.ThueSuatGTGT = vatRate;
            }

            var itemTB = _model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiThietBi || x.MaChiPhi == "G_TB");
            if (itemTB != null)
            {
                itemTB.ThueSuatGTGT = vatRate;
            }

            DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(_model);
            _model.TinhToanLai();

            if (dgvChiPhi != null)
            {
                DanhLaiSoThuTu();
                HienThiDuLieuLenGrid();
            }
            CapNhatThanhTongCong();
        }

        /// <summary>
        /// Xử lý chuyển đổi giữa Bảng Tổng hợp dự toán (Bảng 2.1) và Bảng Tổng mức đầu tư (Bảng 1.2) - Yêu cầu 2
        /// </summary>
        private void ChuyenCheDoBangTinh()
        {
            if (LaCheDoTongMucDauTu)
            {
                // CHẾ ĐỘ TỔNG MỨC ĐẦU TƯ (BẢNG 1.2)
                txtChiPhiBT.Enabled = true;
                var itemBT = _model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.BoiThuong_TDC);
                if (itemBT != null)
                {
                    itemBT.IsActive = true;
                    itemBT.ThueSuatGTGT = 0m;
                }

                // Dự phòng TMĐT chuẩn 10%
                var itemDP = _model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiDuPhong);
                if (itemDP != null) itemDP.TyLePhanTram = 10.0m;

                lblTieuDeTongSauThue.Text = "TỔNG MỨC ĐẦU TƯ:";
            }
            else
            {
                // CHẾ ĐỘ TỔNG HỢP DỰ TOÁN CÔNG TRÌNH (BẢNG 2.1)
                txtChiPhiBT.Enabled = false;
                var itemBT = _model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.BoiThuong_TDC);
                if (itemBT != null) itemBT.IsActive = false; // THDT không có bồi thường GPMB

                // Dự phòng THDT chuẩn 5%
                var itemDP = _model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiDuPhong);
                if (itemDP != null) itemDP.TyLePhanTram = 5.0m;

                lblTieuDeTongSauThue.Text = "TỔNG DỰ TOÁN CT:";
            }

            DanhLaiSoThuTu();
            _model.TinhToanLai();
            HienThiDuLieuLenGrid();
        }

        /// <summary>
        /// Đánh lại số thứ tự chuẩn cho toàn bộ danh mục chi phí theo chế độ TMĐT (Bảng 1.2) hoặc THDT (Bảng 2.1)
        /// </summary>
        private void DanhLaiSoThuTu()
        {
            bool laTMDT = LaCheDoTongMucDauTu;
            bool hasTB = _model.ChiPhiTBTruocThue > 0;

            int sttBT = 1;
            int sttXD = 1;
            int sttTB = 1;
            int sttQLDA = 1;
            int sttTV = 1;
            int sttKhac = 1;
            int sttDP = 1;

            int countBT = _model.Items.Count(x => x.Nhom == NhomChiPhi.BoiThuong_TDC);
            int countTB = _model.Items.Count(x => x.Nhom == NhomChiPhi.ChiPhiThietBi);
            int countQLDA = _model.Items.Count(x => x.Nhom == NhomChiPhi.QuanLyDuAn);

            int prefixXD = laTMDT ? 2 : 1;
            int prefixTB = laTMDT ? 3 : 2;
            int prefixQLDA = hasTB ? (laTMDT ? 4 : 3) : (laTMDT ? 3 : 2);
            int prefixTV = prefixQLDA + 1;
            int prefixKhac = prefixTV + 1;
            int prefixDP = prefixKhac + 1;

            foreach (var item in _model.Items)
            {
                switch (item.Nhom)
                {
                    case NhomChiPhi.BoiThuong_TDC:
                        item.STT = countBT > 1 ? $"1.{sttBT++}" : "1";
                        break;

                    case NhomChiPhi.ChiPhiXayDung:
                        if (item.MaChiPhi == "G_XD") item.STT = $"{prefixXD}.1";
                        else if (item.MaChiPhi == "G_NHA_TAM") item.STT = $"{prefixXD}.2";
                        else item.STT = $"{prefixXD}.{++sttXD}";
                        break;

                    case NhomChiPhi.ChiPhiThietBi:
                        item.STT = countTB > 1 ? $"{prefixTB}.{sttTB++}" : $"{prefixTB}";
                        break;

                    case NhomChiPhi.QuanLyDuAn:
                        item.STT = countQLDA > 1 ? $"{prefixQLDA}.{sttQLDA++}" : $"{prefixQLDA}";
                        break;

                    case NhomChiPhi.TuVanDauTuXD:
                        if (!hasTB && item.CoSoTinh == CoSoTinhChiPhi.ChiPhiThietBi) break;
                        item.STT = $"{prefixTV}.{sttTV++}";
                        break;

                    case NhomChiPhi.ChiPhiKhac:
                        if (!hasTB && item.CoSoTinh == CoSoTinhChiPhi.ChiPhiThietBi) break;
                        item.STT = $"{prefixKhac}.{sttKhac++}";
                        break;

                    case NhomChiPhi.ChiPhiDuPhong:
                        item.STT = $"{prefixDP}.{sttDP++}";
                        break;
                }
            }
        }

        private void HienThiDuLieuLenGrid()
        {
            _isUpdating = true;
            dgvChiPhi.Rows.Clear();

            // Lọc các khoản mục hiển thị:
            // 1. Nếu là THDT thì ẩn nhóm Bồi thường GPMB
            // 2. Nếu không có chi phí Thiết bị (G_TB == 0), ẩn hoàn toàn dòng Chi phí thiết bị và các khoản mục phụ thuộc thiết bị
            bool laTMDT = LaCheDoTongMucDauTu;
            bool hasTB = _model.ChiPhiTBTruocThue > 0;
            int prefixXD = laTMDT ? 2 : 1;
            int prefixQLDA = hasTB ? (laTMDT ? 4 : 3) : (laTMDT ? 3 : 2);
            int prefixTV = prefixQLDA + 1;
            int prefixKhac = prefixTV + 1;
            int prefixDP = prefixKhac + 1;

            var itemsToShow = _model.Items.Where(x =>
            {
                if (!laTMDT && x.Nhom == NhomChiPhi.BoiThuong_TDC) return false;
                if (!hasTB && (x.Nhom == NhomChiPhi.ChiPhiThietBi || x.CoSoTinh == CoSoTinhChiPhi.ChiPhiThietBi)) return false;
                return true;
            }).ToList();

            bool addedXDHeader = false;
            bool addedTVHeader = false;
            bool addedKhacHeader = false;
            bool addedDPHeader = false;

            foreach (var item in itemsToShow)
            {
                // Thêm dòng tổng nhóm 2 (hoặc 1 nếu THDT): Chi phí xây dựng
                if (item.Nhom == NhomChiPhi.ChiPhiXayDung && !addedXDHeader)
                {
                    addedXDHeader = true;
                    decimal ttXD = itemsToShow.Where(x => x.IsActive && x.Nhom == NhomChiPhi.ChiPhiXayDung).Sum(x => x.GiaTriTruocThue);
                    decimal stXD = _model.GetTongSauThueNhom(NhomChiPhi.ChiPhiXayDung);

                    int gRowIdx = dgvChiPhi.Rows.Add();
                    var gRow = dgvChiPhi.Rows[gRowIdx];
                    gRow.Tag = "GROUP_XD";
                    gRow.Cells["colActive"] = new DataGridViewTextBoxCell { Value = "" };
                    gRow.Cells["colSTT"].Value = prefixXD.ToString();
                    gRow.Cells["colTen"].Value = "Chi phí xây dựng";
                    gRow.Cells["colCachTinh"] = new DataGridViewTextBoxCell { Value = "" };
                    gRow.Cells["colCoSo"] = new DataGridViewTextBoxCell { Value = "" };
                    gRow.Cells["colTyLe"].Value = "";
                    gRow.Cells["colHeSo"].Value = "";
                    gRow.Cells["colTruocThue"].Value = UIHelper.FormatTien(ttXD);
                    gRow.Cells["colVAT"] = new DataGridViewTextBoxCell { Value = "" };
                    gRow.Cells["colSauThue"].Value = UIHelper.FormatTien(stXD);
                    gRow.Cells["colKyHieu"].Value = "GXD";
                    gRow.ReadOnly = true;
                    gRow.DefaultCellStyle.Font = UIHelper.GetFont(10.5f, FontStyle.Bold);
                    gRow.DefaultCellStyle.BackColor = Color.FromArgb(220, 235, 252);
                }

                // Thêm dòng tổng nhóm 4 (hoặc 5 nếu có TB): Chi phí tư vấn đầu tư xây dựng
                if (item.Nhom == NhomChiPhi.TuVanDauTuXD && !addedTVHeader)
                {
                    addedTVHeader = true;
                    decimal ttTV = itemsToShow.Where(x => x.IsActive && x.Nhom == NhomChiPhi.TuVanDauTuXD).Sum(x => x.GiaTriTruocThue);
                    decimal stTV = _model.GetTongSauThueNhom(NhomChiPhi.TuVanDauTuXD);

                    int gRowIdx = dgvChiPhi.Rows.Add();
                    var gRow = dgvChiPhi.Rows[gRowIdx];
                    gRow.Tag = "GROUP_TV";
                    gRow.Cells["colActive"] = new DataGridViewTextBoxCell { Value = "" };
                    gRow.Cells["colSTT"].Value = prefixTV.ToString();
                    gRow.Cells["colTen"].Value = "Chi phí tư vấn đầu tư xây dựng";
                    gRow.Cells["colCachTinh"] = new DataGridViewTextBoxCell { Value = "" };
                    gRow.Cells["colCoSo"] = new DataGridViewTextBoxCell { Value = "" };
                    gRow.Cells["colTyLe"].Value = "";
                    gRow.Cells["colHeSo"].Value = "";
                    gRow.Cells["colTruocThue"].Value = UIHelper.FormatTien(ttTV);
                    gRow.Cells["colVAT"] = new DataGridViewTextBoxCell { Value = "" };
                    gRow.Cells["colSauThue"].Value = UIHelper.FormatTien(stTV);
                    gRow.Cells["colKyHieu"].Value = "Gtv";
                    gRow.ReadOnly = true;
                    gRow.DefaultCellStyle.Font = UIHelper.GetFont(10.5f, FontStyle.Bold);
                    gRow.DefaultCellStyle.BackColor = Color.FromArgb(220, 235, 252);
                }

                // Thêm dòng tổng nhóm 5 (hoặc 6 nếu có TB): Chi phí khác
                if (item.Nhom == NhomChiPhi.ChiPhiKhac && !addedKhacHeader)
                {
                    addedKhacHeader = true;
                    decimal ttK = itemsToShow.Where(x => x.IsActive && x.Nhom == NhomChiPhi.ChiPhiKhac).Sum(x => x.GiaTriTruocThue);
                    decimal stK = _model.GetTongSauThueNhom(NhomChiPhi.ChiPhiKhac);

                    int gRowIdx = dgvChiPhi.Rows.Add();
                    var gRow = dgvChiPhi.Rows[gRowIdx];
                    gRow.Tag = "GROUP_KHAC";
                    gRow.Cells["colActive"] = new DataGridViewTextBoxCell { Value = "" };
                    gRow.Cells["colSTT"].Value = prefixKhac.ToString();
                    gRow.Cells["colTen"].Value = "Chi phí khác";
                    gRow.Cells["colCachTinh"] = new DataGridViewTextBoxCell { Value = "" };
                    gRow.Cells["colCoSo"] = new DataGridViewTextBoxCell { Value = "" };
                    gRow.Cells["colTyLe"].Value = "";
                    gRow.Cells["colHeSo"].Value = "";
                    gRow.Cells["colTruocThue"].Value = UIHelper.FormatTien(ttK);
                    gRow.Cells["colVAT"] = new DataGridViewTextBoxCell { Value = "" };
                    gRow.Cells["colSauThue"].Value = UIHelper.FormatTien(stK);
                    gRow.Cells["colKyHieu"].Value = "Gk";
                    gRow.ReadOnly = true;
                    gRow.DefaultCellStyle.Font = UIHelper.GetFont(10.5f, FontStyle.Bold);
                    gRow.DefaultCellStyle.BackColor = Color.FromArgb(220, 235, 252);
                }

                // Thêm dòng tổng nhóm 6 (hoặc 7 nếu có TB): Chi phí dự phòng
                if (item.Nhom == NhomChiPhi.ChiPhiDuPhong && !addedDPHeader)
                {
                    addedDPHeader = true;
                    decimal ttDP = itemsToShow.Where(x => x.IsActive && x.Nhom == NhomChiPhi.ChiPhiDuPhong).Sum(x => x.GiaTriTruocThue);
                    decimal stDP = _model.GetTongSauThueNhom(NhomChiPhi.ChiPhiDuPhong);

                    int gRowIdx = dgvChiPhi.Rows.Add();
                    var gRow = dgvChiPhi.Rows[gRowIdx];
                    gRow.Tag = "GROUP_DP";
                    gRow.Cells["colActive"] = new DataGridViewTextBoxCell { Value = "" };
                    gRow.Cells["colSTT"].Value = prefixDP.ToString();
                    gRow.Cells["colTen"].Value = "Chi phí dự phòng";
                    gRow.Cells["colCachTinh"] = new DataGridViewTextBoxCell { Value = "" };
                    gRow.Cells["colCoSo"] = new DataGridViewTextBoxCell { Value = "" };
                    gRow.Cells["colTyLe"].Value = "";
                    gRow.Cells["colHeSo"].Value = "";
                    gRow.Cells["colTruocThue"].Value = UIHelper.FormatTien(ttDP);
                    gRow.Cells["colVAT"] = new DataGridViewTextBoxCell { Value = "" };
                    gRow.Cells["colSauThue"].Value = UIHelper.FormatTien(stDP);
                    gRow.Cells["colKyHieu"].Value = "GDP";
                    gRow.ReadOnly = true;
                    gRow.DefaultCellStyle.Font = UIHelper.GetFont(10.5f, FontStyle.Bold);
                    gRow.DefaultCellStyle.BackColor = Color.FromArgb(220, 235, 252);
                }

                string cachTinhStr = item.CachTinh == CachTinhChiPhi.TheoTyLeDinhMuc ? "Theo tỷ lệ %" : "Tự nhập tiền";
                string coSoStr = "-";
                if (item.CachTinh == CachTinhChiPhi.TheoTyLeDinhMuc)
                {
                    switch (item.CoSoTinh)
                    {
                        case CoSoTinhChiPhi.ChiPhiXayDung: coSoStr = "G_XD"; break;
                        case CoSoTinhChiPhi.ChiPhiThietBi: coSoStr = "G_TB"; break;
                        case CoSoTinhChiPhi.TongChiPhiTruocDuPhong: coSoStr = "Tổng trước DP"; break;
                        case CoSoTinhChiPhi.TongMucDauTu: coSoStr = "Toàn bộ TMĐT"; break;
                        default: coSoStr = "G_XD + G_TB"; break;
                    }
                }

                string vatStr = $"{(item.ThueSuatGTGT * 100m):G29}%";
                var colVATCol = dgvChiPhi.Columns["colVAT"] as DataGridViewComboBoxColumn;
                if (colVATCol != null && !colVATCol.Items.Contains(vatStr))
                {
                    colVATCol.Items.Add(vatStr);
                }

                string tenHienThi = item.TenChiPhi;
                if ((item.Nhom == NhomChiPhi.ChiPhiXayDung || item.Nhom == NhomChiPhi.ChiPhiDuPhong) && !tenHienThi.StartsWith("-"))
                {
                    tenHienThi = "- " + tenHienThi;
                }

                int rowIdx = dgvChiPhi.Rows.Add(
                    item.IsActive,
                    item.STT,
                    tenHienThi,
                    cachTinhStr,
                    coSoStr,
                    item.TyLePhanTram > 0 ? UIHelper.FormatTyLe(item.TyLePhanTram) : "",
                    item.HeSoDieuChinh.ToString("0.00", UIHelper.ViCulture),
                    UIHelper.FormatTien(item.GiaTriTruocThue),
                    vatStr,
                    UIHelper.FormatTien(item.GiaTriSauThue),
                    item.KyHieu
                );

                var row = dgvChiPhi.Rows[rowIdx];
                row.Tag = item;

                if (item.CachTinh != CachTinhChiPhi.TheoTyLeDinhMuc)
                {
                    row.Cells["colCoSo"].ReadOnly = true;
                }

                // Định dạng nổi bật các dòng chi phí đơn cấp 1 (Bồi thường, Thiết bị, QLDA)
                if (item.Nhom == NhomChiPhi.BoiThuong_TDC || 
                    item.Nhom == NhomChiPhi.ChiPhiThietBi || 
                    item.Nhom == NhomChiPhi.QuanLyDuAn)
                {
                    row.DefaultCellStyle.Font = UIHelper.GetFont(10.5f, FontStyle.Bold);
                    row.DefaultCellStyle.BackColor = Color.FromArgb(220, 235, 252);
                }
                else
                {
                    // Tất cả các dòng chi tiết con (2.1, 2.2, 4.1.., 5.1.., 6.1..) hiển thị chữ thường
                    row.DefaultCellStyle.Font = UIHelper.GetFont(10f, FontStyle.Regular);
                }

                if (!item.IsActive)
                {
                    row.DefaultCellStyle.ForeColor = Color.Gray;
                }
            }

            CapNhatThanhTongCong();
            dgvChiPhi?.AutoFit();
            _isUpdating = false;
        }

        private void CapNhatTongNhomTrenGrid()
        {
            if (dgvChiPhi == null || dgvChiPhi.Rows.Count == 0) return;

            foreach (DataGridViewRow r in dgvChiPhi.Rows)
            {
                if (r.Tag is string tag)
                {
                    if (tag == "GROUP_XD")
                    {
                        decimal ttXD = _model.Items.Where(x => x.IsActive && x.Nhom == NhomChiPhi.ChiPhiXayDung).Sum(x => x.GiaTriTruocThue);
                        decimal stXD = _model.GetTongSauThueNhom(NhomChiPhi.ChiPhiXayDung);
                        r.Cells["colTruocThue"].Value = UIHelper.FormatTien(ttXD);
                        r.Cells["colSauThue"].Value = UIHelper.FormatTien(stXD);
                    }
                    else if (tag == "GROUP_TV")
                    {
                        decimal ttTV = _model.Items.Where(x => x.IsActive && x.Nhom == NhomChiPhi.TuVanDauTuXD).Sum(x => x.GiaTriTruocThue);
                        decimal stTV = _model.GetTongSauThueNhom(NhomChiPhi.TuVanDauTuXD);
                        r.Cells["colTruocThue"].Value = UIHelper.FormatTien(ttTV);
                        r.Cells["colSauThue"].Value = UIHelper.FormatTien(stTV);
                    }
                    else if (tag == "GROUP_KHAC")
                    {
                        decimal ttK = _model.Items.Where(x => x.IsActive && x.Nhom == NhomChiPhi.ChiPhiKhac).Sum(x => x.GiaTriTruocThue);
                        decimal stK = _model.GetTongSauThueNhom(NhomChiPhi.ChiPhiKhac);
                        r.Cells["colTruocThue"].Value = UIHelper.FormatTien(ttK);
                        r.Cells["colSauThue"].Value = UIHelper.FormatTien(stK);
                    }
                    else if (tag == "GROUP_DP")
                    {
                        decimal ttDP = _model.Items.Where(x => x.IsActive && x.Nhom == NhomChiPhi.ChiPhiDuPhong).Sum(x => x.GiaTriTruocThue);
                        decimal stDP = _model.GetTongSauThueNhom(NhomChiPhi.ChiPhiDuPhong);
                        r.Cells["colTruocThue"].Value = UIHelper.FormatTien(ttDP);
                        r.Cells["colSauThue"].Value = UIHelper.FormatTien(stDP);
                    }
                }
            }
        }

        private void CapNhatThanhTongCong()
        {
            CapNhatTongNhomTrenGrid();
            lblTongTruocThue.Text = $"Trước thuế: {UIHelper.FormatTien(_model.TongTruocThue)} đ";
            lblTongGTGT.Text = $"Thuế GTGT: {UIHelper.FormatTien(_model.TongTienThueGTGT)} đ";
            lblTongSauThue.Text = $"{UIHelper.FormatTien(_model.TongSauThue)} đ";

            string docChu = UIHelper.DocSoThanhChu(_model.TongSauThue);
            if (lblBangChu != null)
            {
                lblBangChu.Text = $"Bằng chữ: {docChu}.";
            }
            if (_toolTip != null)
            {
                _toolTip.SetToolTip(lblTongSauThue, $"Bằng chữ: {docChu}");
                _toolTip.SetToolTip(lblTieuDeTongSauThue, $"Bằng chữ: {docChu}");
            }
        }

        private void CapNhatKhiDoiThongSo(bool traLaiDinhMuc)
        {
            if (_isUpdating) return;

            string loaiSel = cboLoaiCongTrinh.SelectedIndex > 0 ? cboLoaiCongTrinh.SelectedItem?.ToString() : "";
            string capSel = cboCapCongTrinh.SelectedIndex > 0 ? cboCapCongTrinh.SelectedItem?.ToString() : "";
            int buocSel = cboSoBuocThietKe.SelectedIndex > 0 ? cboSoBuocThietKe.SelectedIndex : 0;

            _model.LoaiCongTrinh = loaiSel;
            _model.CapCongTrinh = capSel;
            _model.SoBuocThietKe = buocSel;

            // Đồng bộ Loại công trình ngược về Tab 1
            if (!_isSyncingLoaiCT && cboLoaiCongTrinhXD != null)
            {
                try
                {
                    _isSyncingLoaiCT = true;
                    if (cboLoaiCongTrinh.SelectedIndex <= 0)
                    {
                        if (cboLoaiCongTrinhXD.SelectedIndex != 0)
                            cboLoaiCongTrinhXD.SelectedIndex = 0;
                    }
                    else if (!string.IsNullOrEmpty(loaiSel))
                    {
                        string target = loaiSel;
                        if (target == "Nông nghiệp và môi trường" || target == "Nông nghiệp & PTNT") target = DinhMucTT38Database.LoaiNongNghiepMoiTruong;
                        if (cboLoaiCongTrinhXD.Items.Contains(target) && cboLoaiCongTrinhXD.SelectedItem?.ToString() != target)
                        {
                            cboLoaiCongTrinhXD.SelectedItem = target;
                        }
                    }
                }
                finally { _isSyncingLoaiCT = false; }
            }

            _model.ThietBiTren50Pct = chkThietBi50.Checked;
            _model.DaKiemToanDocLap = chkDaKiemToan.Checked;
            _model.YeuCauThueThamTra = chkThueThamTra.Checked;
            _model.CdtTuQuanLy = chkCdtTuQL.Checked;
            _model.VungKhoKhan = chkVungKhoKhan.Checked;
            _model.TuyenQuaNhieuTinh = chkTuyenNhieuTinh.Checked;
            _model.CaiTaoSuaChua = chkCaiTao.Checked;
            _model.ThietKeLapLai = chkLapLai.Checked;

            if (traLaiDinhMuc && !string.IsNullOrEmpty(loaiSel) && !string.IsNullOrEmpty(capSel))
            {
                DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(_model);
                var itemDP = _model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiDuPhong);
                if (itemDP != null)
                {
                    itemDP.TyLePhanTram = LaCheDoTongMucDauTu ? 10.0m : 5.0m;
                }
            }

            _model.TinhToanLai();
            HienThiDuLieuLenGrid();
        }

        private void CapNhatGiaTriDauVao()
        {
            if (_isUpdating) return;

            _model.ChiPhiXDTruocThue = UIHelper.ParseTien(txtChiPhiXD.Text);
            _model.ChiPhiNhaTamTruocThue = UIHelper.ParseTien(txtChiPhiNT.Text);
            _model.ChiPhiTBTruocThue = UIHelper.ParseTien(txtChiPhiTB.Text);
            _model.ChiPhiBTTruocThue = UIHelper.ParseTien(txtChiPhiBT.Text);

            // Format lại đẹp chuẩn dấu chấm
            txtChiPhiXD.Text = UIHelper.FormatTien(_model.ChiPhiXDTruocThue);
            txtChiPhiNT.Text = UIHelper.FormatTien(_model.ChiPhiNhaTamTruocThue);
            txtChiPhiTB.Text = UIHelper.FormatTien(_model.ChiPhiTBTruocThue);
            txtChiPhiBT.Text = UIHelper.FormatTien(_model.ChiPhiBTTruocThue);

            // Tự động kích hoạt chi phí thiết bị và chi phí giám sát lắp đặt thiết bị (TV_GS_TB) khi chi phí thiết bị > 0
            var itemTB = _model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiThietBi);
            if (itemTB != null)
            {
                itemTB.GiaTriTruocThue = _model.ChiPhiTBTruocThue;
                itemTB.IsActive = _model.ChiPhiTBTruocThue > 0;
            }

            var itemGSTB = _model.Items.FirstOrDefault(x => x.MaChiPhi == "TV_GS_TB");
            if (itemGSTB != null)
            {
                itemGSTB.IsActive = _model.ChiPhiTBTruocThue > 0;
            }

            // Cập nhật lại định mức cho các khoản mục phụ thuộc G_TB và G_XD + G_TB
            DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(_model);

            DanhLaiSoThuTu();
            _model.TinhToanLai();
            HienThiDuLieuLenGrid();
        }

        private static decimal ParseVATRate(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0.10m;
            string cleaned = text.Replace("%", "").Trim().Replace(',', '.');
            if (decimal.TryParse(cleaned, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal val))
            {
                if (text.Contains("%"))
                    return val / 100m;
                if (val > 1m)
                    return val / 100m;
                if (val > 0m && val < 0.5m)
                    return val;
                return val / 100m;
            }
            return 0.10m;
        }

        private void XuLyDoiVATChung()
        {
            if (_isUpdating) return;
            string text = cboVATChung?.Text?.Trim() ?? "";
            if (string.IsNullOrEmpty(text)) return;
            decimal rate = ParseVATRate(text);
            string formatted = $"{(rate * 100m):G29}%";

            if (!cboVATChung.Items.Contains(formatted))
            {
                cboVATChung.Items.Add(formatted);
            }
            if (cboVATChung.Text != formatted)
            {
                cboVATChung.Text = formatted;
            }

            ApDungVATToanBang(rate);
        }

        private void ApDungVATToanBang()
        {
            decimal rate = ParseVATRate(cboVATChung?.Text ?? "10%");
            ApDungVATToanBang(rate);
        }

        private void ApDungVATToanBang(decimal newVAT)
        {
            if (_model == null) return;

            string formatted = $"{(newVAT * 100m):G29}%";

            // Đảm bảo colVAT trong DataGridView có mục này
            var colVAT = dgvChiPhi.Columns["colVAT"] as DataGridViewComboBoxColumn;
            if (colVAT != null && !colVAT.Items.Contains(formatted))
            {
                colVAT.Items.Add(formatted);
            }

            // cboVATChung chỉ áp dụng trên tab TMĐT/Tổng hợp dự toán, không đồng bộ sang Tab 1

            foreach (var item in _model.Items)
            {
                // Mặc định Chi phí QLDA không chịu thuế GTGT (0%) và giữ 0% cho các khoản phí ngân sách nhà nước
                // G_XD và G_NHA_TAM được đồng bộ riêng từ Tab Chi phí XD, không đổi qua cboVATChung
                // G_BT luôn 0% (chi phí bồi thường không chịu thuế GTGT)
                if (item.MaChiPhi == "G_QLDA" || item.Nhom == NhomChiPhi.QuanLyDuAn ||
                    item.MaChiPhi == "K_TD_DA" || item.MaChiPhi == "K_TD_TK" || 
                    item.MaChiPhi == "K_TD_DT" || item.MaChiPhi == "K_TT_QUYETTOAN" ||
                    item.MaChiPhi == "G_XD" || item.MaChiPhi == "G_NHA_TAM" ||
                    item.MaChiPhi == "G_BT")
                {
                    continue;
                }
                item.ThueSuatGTGT = newVAT;
            }

            _model.TinhToanLai();
            HienThiDuLieuLenGrid();
            CapNhatThanhTongCong();
        }

        /// <summary>
        /// Phần D - Popup Kiểm tra định mức: click đúp vào dòng khoản mục chi phí để tra cứu
        /// bảng định mức TT 38/2026 đang áp dụng (bảng tra, tỷ lệ nội suy, trang PDF...).
        /// </summary>
        private void DgvChiPhi_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (_dangKiemTraDinhMuc) return;
            var row = dgvChiPhi.Rows[e.RowIndex];
            if (row.Tag is not ChiPhiKinhPhiItem item) return;

            _dangKiemTraDinhMuc = true;
            try
            {
                // 1.3: tính toán lại model trước khi tra cứu để số liệu nội suy phản ánh đúng quy mô hiện tại
                _model.TinhToanLai();
                var tt = DinhMucTT38Engine.TraCuuThongTinChiPhiItem(_model, item.MaChiPhi);
                if (tt == null)
                {
                    MessageBox.Show(
                        $"Khoản mục [{item.TenChiPhi}] ({item.MaChiPhi}) không áp dụng định mức tỷ lệ theo TT 38/2026 —\nchỉ nhập trực tiếp theo dự toán riêng (hoặc theo văn bản Bộ Tài chính).",
                        "Kiểm tra định mức", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using var frm = new ThongTinDinhMucForm(tt, item, _model, () =>
                {
                    HienThiDuLieuLenGrid();
                    CapNhatThanhTongCong();
                });
                frm.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tra cứu định mức: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _dangKiemTraDinhMuc = false;
            }
        }

        private void DgvChiPhi_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (_isUpdating || e.RowIndex < 0 || e.ColumnIndex < 0) return;

            var row = dgvChiPhi.Rows[e.RowIndex];
            if (row.Tag is not ChiPhiKinhPhiItem item) return;

            string colName = dgvChiPhi.Columns[e.ColumnIndex].Name;

            if (colName == "colActive")
            {
                item.IsActive = Convert.ToBoolean(row.Cells["colActive"].Value ?? true);
                row.DefaultCellStyle.ForeColor = item.IsActive ? Color.Black : Color.Gray;
                DinhMucTT38Engine.CapNhatTiLeThamDinhDuAn(_model, null, _model.YeuCauThueThamTra);
                CapNhatThanhTongCong();
                return;
            }

            if (colName == "colTen")
            {
                item.TenChiPhi = row.Cells["colTen"].Value?.ToString() ?? "";
            }
            else if (colName == "colCachTinh")
            {
                string ct = row.Cells["colCachTinh"].Value?.ToString() ?? "";
                item.CachTinh = ct.Contains("tỷ lệ") ? CachTinhChiPhi.TheoTyLeDinhMuc : CachTinhChiPhi.NhapTruocThue;
            }
            else if (colName == "colCoSo")
            {
                string cs = row.Cells["colCoSo"].Value?.ToString() ?? "";
                if (cs == "G_XD") item.CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung;
                else if (cs == "G_TB") item.CoSoTinh = CoSoTinhChiPhi.ChiPhiThietBi;
                else if (cs == "Tổng trước DP") item.CoSoTinh = CoSoTinhChiPhi.TongChiPhiTruocDuPhong;
                else if (cs == "Toàn bộ TMĐT") item.CoSoTinh = CoSoTinhChiPhi.TongMucDauTu;
                else item.CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi;
            }
            else if (colName == "colTyLe")
            {
                item.TyLePhanTram = UIHelper.ParseTyLe(row.Cells["colTyLe"].Value?.ToString() ?? "0");
            }
            else if (colName == "colHeSo")
            {
                item.HeSoDieuChinh = UIHelper.ParseTyLe(row.Cells["colHeSo"].Value?.ToString() ?? "1");
            }
            else if (colName == "colTruocThue")
            {
                decimal tt = UIHelper.ParseTien(row.Cells["colTruocThue"].Value?.ToString() ?? "0");
                item.GiaTriTruocThue = tt;
                item.CachTinh = CachTinhChiPhi.NhapTruocThue;
                if (item.MaChiPhi == "G_XD")
                {
                    _model.ChiPhiXDTruocThue = tt;
                    txtChiPhiXD.Text = UIHelper.FormatTien(tt);
                }
                else if (item.MaChiPhi == "G_NHA_TAM")
                {
                    _model.ChiPhiNhaTamTruocThue = tt;
                    txtChiPhiNT.Text = UIHelper.FormatTien(tt);
                }
                else if (item.Nhom == NhomChiPhi.ChiPhiThietBi)
                {
                    _model.ChiPhiTBTruocThue = tt;
                    txtChiPhiTB.Text = UIHelper.FormatTien(tt);
                    var itemGSTB = _model.Items.FirstOrDefault(x => x.MaChiPhi == "TV_GS_TB");
                    if (itemGSTB != null)
                    {
                        itemGSTB.IsActive = tt > 0;
                    }
                    DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(_model);
                }
                else if (item.Nhom == NhomChiPhi.BoiThuong_TDC)
                {
                    _model.ChiPhiBTTruocThue = tt;
                    txtChiPhiBT.Text = UIHelper.FormatTien(tt);
                }
            }
            else if (colName == "colVAT")
            {
                // Yêu cầu 7: Cho phép sửa thuế VAT theo từng dòng riêng lẻ
                string vatStr = row.Cells["colVAT"].Value?.ToString() ?? "10%";
                decimal vatRate = ParseVATRate(vatStr);
                item.ThueSuatGTGT = vatRate;
                string formatted = $"{(vatRate * 100m):G29}%";

                var colVAT = dgvChiPhi.Columns["colVAT"] as DataGridViewComboBoxColumn;
                if (colVAT != null && !colVAT.Items.Contains(formatted))
                {
                    colVAT.Items.Add(formatted);
                }

                // Nếu sửa VAT của Chi phí xây dựng, Lán trại tạm hoặc Chi phí thiết bị:
                // 1. Đồng bộ cả 3 khoản mục G_XD, G_NHA_TAM, G_TB dùng chung thuế suất VAT
                // 2. Đồng bộ hai chiều sang Tab 1 (txtGTGTXD) để tự động tính toán lại bảng chi phí xây dựng
                if (item.MaChiPhi == "G_XD" || item.MaChiPhi == "G_NHA_TAM" || item.MaChiPhi == "G_TB" || item.Nhom == NhomChiPhi.ChiPhiThietBi)
                {
                    var itemXD = _model.Items.FirstOrDefault(x => x.MaChiPhi == "G_XD");
                    var itemNT = _model.Items.FirstOrDefault(x => x.MaChiPhi == "G_NHA_TAM");
                    var itemTB = _model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiThietBi || x.MaChiPhi == "G_TB");
                    if (itemXD != null) itemXD.ThueSuatGTGT = vatRate;
                    if (itemNT != null) itemNT.ThueSuatGTGT = vatRate;
                    if (itemTB != null) itemTB.ThueSuatGTGT = vatRate;

                    _model.TinhToanLai();
                    _isUpdating = true;
                    foreach (DataGridViewRow r in dgvChiPhi.Rows)
                    {
                        var it = r.Tag as ChiPhiKinhPhiItem;
                        if (it != null && (it.MaChiPhi == "G_XD" || it.MaChiPhi == "G_NHA_TAM" || it.MaChiPhi == "G_TB" || it.Nhom == NhomChiPhi.ChiPhiThietBi))
                        {
                            r.Cells["colVAT"].Value = formatted;
                            r.Cells["colTruocThue"].Value = UIHelper.FormatTien(it.GiaTriTruocThue);
                            r.Cells["colSauThue"].Value = UIHelper.FormatTien(it.GiaTriSauThue);
                        }
                    }
                    _isUpdating = false;
                    CapNhatThanhTongCong();

                    this.BeginInvoke(new Action(() =>
                    {
                        if (txtGTGTXD != null)
                        {
                            string vatNum = (vatRate * 100m).ToString("G29", UIHelper.ViCulture);
                            if (txtGTGTXD.Text.Trim() != vatNum)
                            {
                                try
                                {
                                    _isSyncingVAT = true;
                                    txtGTGTXD.Text = vatNum; // Kích hoạt txtGTGTXD.TextChanged -> TinhToanChiPhiXD() ở Tab 1
                                }
                                finally
                                {
                                    _isSyncingVAT = false;
                                }
                            }
                        }
                    }));
                    return;
                }
            }
            else if (colName == "colKyHieu")
            {
                item.KyHieu = row.Cells["colKyHieu"].Value?.ToString() ?? "";
            }

            _model.TinhToanLai();
            if (item.MaChiPhi != "K_TD_DA")
            {
                DinhMucTT38Engine.CapNhatTiLeThamDinhDuAn(_model, null, _model.YeuCauThueThamTra);
            }
            _isUpdating = true;
            row.Cells["colTyLe"].Value = item.TyLePhanTram > 0 ? UIHelper.FormatTyLe(item.TyLePhanTram) : "";
            row.Cells["colTruocThue"].Value = UIHelper.FormatTien(item.GiaTriTruocThue);
            row.Cells["colSauThue"].Value = UIHelper.FormatTien(item.GiaTriSauThue);
            _isUpdating = false;
            CapNhatThanhTongCong();
        }

        private void BtnThemThuVien_Click(object sender, EventArgs e)
        {
            using var dlg = new ChonChiPhiThuVienForm();
            if (dlg.ShowDialog() == DialogResult.OK && dlg.SelectedItems.Count > 0)
            {
                int insertIdx = _model.Items.FindIndex(x => x.Nhom == NhomChiPhi.ChiPhiDuPhong);
                if (insertIdx < 0) insertIdx = _model.Items.Count;

                foreach (var item in dlg.SelectedItems)
                {
                    if (_model.Items.Any(x => x.MaChiPhi == item.MaChiPhi))
                    {
                        item.MaChiPhi += "_" + Guid.NewGuid().ToString("N").Substring(0, 4);
                    }
                    _model.Items.Insert(insertIdx++, item);
                }

                DanhLaiSoThuTu();
                _model.TinhToanLai();
                HienThiDuLieuLenGrid();
            }
        }

        private void BtnThemTuyBien_Click(object sender, EventArgs e)
        {
            var newItem = new ChiPhiKinhPhiItem
            {
                MaChiPhi = "K_CUSTOM_" + (_model.Items.Count + 1),
                TenChiPhi = "Chi phí mới (Nhấp đôi để sửa tên)",
                Nhom = NhomChiPhi.ChiPhiKhac,
                CachTinh = CachTinhChiPhi.NhapTruocThue,
                CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                GiaTriTruocThue = 0m,
                ThueSuatGTGT = 0.10m,
                KyHieu = "Gk" + (_model.Items.Count(x => x.Nhom == NhomChiPhi.ChiPhiKhac) + 1),
                IsActive = true,
                IsUserAdded = true
            };

            int insertIdx = _model.Items.FindIndex(x => x.Nhom == NhomChiPhi.ChiPhiDuPhong);
            if (insertIdx < 0) insertIdx = _model.Items.Count;

            _model.Items.Insert(insertIdx, newItem);
            DanhLaiSoThuTu();
            HienThiDuLieuLenGrid();
        }

        private void BtnXoaChiPhi_Click(object sender, EventArgs e)
        {
            if (dgvChiPhi.CurrentRow?.Tag is not ChiPhiKinhPhiItem item) return;

            if (item.IsReadOnly)
            {
                MessageBox.Show($"Khoản mục [{item.TenChiPhi}] là chi phí cốt lõi của bảng tính, không được xóa!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show($"Bạn có chắc muốn xóa khoản mục [{item.TenChiPhi}]?", "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                _model.Items.Remove(item);
                DanhLaiSoThuTu();
                _model.TinhToanLai();
                HienThiDuLieuLenGrid();
            }
        }

        private void BtnLuu_Click(object sender, EventArgs e)
        {
            _duToan.BangKinhPhi = _model;
            _duToan.LoaiCongTrinh = _model.LoaiCongTrinh;
            _duToan.CapCongTrinh = _model.CapCongTrinh;

            MessageBox.Show("Đã lưu thiết lập Bảng tổng hợp kinh phí thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Kiểm tra tính hợp lệ của dự toán trước khi xuất: cảnh báo khi thiếu dữ liệu cơ bản
        /// hoặc các tỷ lệ định mức phải dùng giá trị mặc định. Trả về danh sách cảnh báo (rỗng = OK).
        /// </summary>
        private List<string> KiemTraHopLe()
        {
            var list = new List<string>();
            if (_duToan == null)
            {
                list.Add("Không có dữ liệu dự toán.");
                return list;
            }

            int tongCongTac = _duToan.DanhSachHangMuc?.Sum(h => h.DanhSachCongTac.Count) ?? 0;
            if (tongCongTac == 0)
                list.Add("Dự toán không có công tác nào → các giá trị chi phí sẽ ra 0 đồng.");

            if (string.IsNullOrEmpty(_duToan.LoaiCongTrinh))
                list.Add("Loại công trình chưa chọn → các tỷ lệ định mức tính theo mặc định \"Dân dụng\".");

            var cpxd = _duToan.ChiPhiXD;
            if (cpxd == null)
            {
                list.Add("Chưa có dữ liệu Chi phí xây dựng → sẽ tính lại trong quá trình xuất.");
                return list;
            }

            // Nạp định mức Bảng 3.3/3.5 từ CSDL (tương tự BoSungTyLeVaTinhLai) để KHÔNG báo nhầm
            // "thiếu bảng định mức" khi tỷ lệ thực tế đã khớp Thông tư nhưng model chưa lưu giá trị %.
            string loaiCT = !string.IsNullOrEmpty(_duToan.LoaiCongTrinh) ? _duToan.LoaiCongTrinh : "Dân dụng";
            List<DinhMucCPC> cpcApp = null;
            DinhMucTT ttApp = null;
            try
            {
                if (cpxd.TiLeCPC <= 0m)
                    cpcApp = _cpcRepo.GetByLoaiCongTrinh(loaiCT, null);
                if (cpxd.TiLeTT <= 0m)
                    ttApp = _ttRepo.GetByLoaiCongTrinh(loaiCT, null);
            }
            catch { }

            var rate = ChiPhiXayDungRateCalc.TinhTyLe(
                loaiCT,
                null,
                cpxd.GXDTT,
                _duToan.DanhSachHangMuc?.Sum(h => h.TongVL + h.TongNC + h.TongMay) ?? 0m,
                _duToan.SoBuocThietKe > 0 ? _duToan.SoBuocThietKe : 2,
                cpxd.LaVungSauXa,
                cpxd.LoaiCongTrinhNhaTam,
                cpxd.TiLeCPC, cpxd.TiLeTT, cpxd.TiLeTNCTTT, cpxd.TiLeGTGT, cpxd.TiLeNhaTam,
                cpcApp, ttApp);

            foreach (var w in rate.CanhBao.Distinct())
                list.Add(w);

            return list;
        }

        /// <summary>
        /// Đối chiếu sau xuất: so sánh Chi phí xây dựng (GXD) trên sheet
        /// TH_ChiPhiXD / TongMucDauTu với GXD tính trong model; chênh lệch &gt; 1 đồng → cảnh báo.
        /// Phương thức "mềm": lỗi đọc dữ liệu sẽ bỏ qua, không làm gián đoạn luồng xuất.
        /// </summary>
        private List<string> KiemTraDoiChieuBieuMau(ExcelWb wb)
        {
            var ds = new List<string>();
            try
            {
                if (wb == null || _duToan == null) return ds;

                Worksheet th = null, tmdt = null, duToan = null;
                foreach (Worksheet s in wb.Sheets)
                {
                    string n = s.Name;
                    if (th == null && n.StartsWith("TH_ChiPhiXD", StringComparison.OrdinalIgnoreCase)) th = s;
                    else if (tmdt == null && n.StartsWith("TongMucDauTu", StringComparison.OrdinalIgnoreCase)) tmdt = s;
                    else if (duToan == null && n.StartsWith("DuToan", StringComparison.OrdinalIgnoreCase)) duToan = s;
                }

                // GXD model: đa hạng mục = tổng GXD từng hạng mục; ngược lại lấy ChiPhiXD toàn dự toán
                decimal gxdModel = _duToan.ChiPhiXD?.GXD ?? 0m;
                if (_duToan.DanhSachHangMuc != null && _duToan.DanhSachHangMuc.Count > 1)
                    gxdModel = _duToan.DanhSachHangMuc.Sum(h => h.ChiPhiXD?.GXD ?? 0m);

                if (th != null && gxdModel > 0m)
                {
                    decimal gxdSheet = 0m;
                    int lastRow = th.UsedRange.Row + th.UsedRange.Rows.Count - 1;
                    for (int r = lastRow; r >= 6; r--)
                    {
                        string b = th.Cells[r, 2]?.Value2?.ToString()?.Trim() ?? "";
                        if (b.ToUpper().Contains("SAU THUẾ"))
                        {
                            gxdSheet = ToDong(th.Cells[r, 5].Value2);
                            break;
                        }
                    }
                    if (gxdSheet > 0m && Math.Abs(gxdSheet - gxdModel) > 1m)
                        ds.Add($"TH_ChiPhiXD: GXD = {gxdSheet:N0} đồng, khác GXD tính trong dự toán ({gxdModel:N0} đồng).");
                }

                if (tmdt != null && gxdModel > 0m)
                {
                    decimal gxdTM = 0m;
                    int lastRow = tmdt.UsedRange.Row + tmdt.UsedRange.Rows.Count - 1;
                    for (int r = lastRow; r >= 5; r--)
                    {
                        string b = tmdt.Cells[r, 2]?.Value2?.ToString()?.Trim() ?? "";
                        if (b.ToUpper().Contains("CHI PHÍ XÂY DỰNG"))
                        {
                            gxdTM = ToDong(tmdt.Cells[r, 5].Value2);
                            if (gxdTM > 0m) break;
                        }
                    }
                    if (gxdTM > 0m && Math.Abs(gxdTM - gxdModel) > 1m)
                        ds.Add($"TongMucDauTu: GXD = {gxdTM:N0} đồng, khác GXD tính trong dự toán ({gxdModel:N0} đồng).");
                }
            }
            catch
            {
                // Không làm gián đoạn luồng xuất nếu không đọc được bảng
            }
            return ds;
        }

        private static decimal ToDong(object val)
        {
            try
            {
                if (val == null) return 0m;
                if (val is double d) return (decimal)d;
                if (val is int i) return i;
                if (val is decimal m) return m;
                if (val is float f) return (decimal)f;
                if (val is long l) return l;
                return decimal.TryParse(val.ToString().Trim(), out var r) ? r : 0m;
            }
            catch
            {
                return 0m;
            }
        }

        private void XuatExcelTongHop()
        {
            try
            {
                var app = (ExcelApp)ExcelDnaUtil.Application;
                var wb = app.ActiveWorkbook;
                if (wb == null)
                {
                    MessageBox.Show("Không tìm thấy Workbook Excel nào đang mở.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Đảm bảo dữ liệu Chi phí XD và Bảng kinh phí đã được tính toán và đồng bộ đồng nhất
                TinhToanChiPhiXD();
                CapNhatGiaTriDauVao();
                _duToan.BangKinhPhi = _model;

                // Kiểm tra tính hợp lệ & cảnh báo trước khi xuất (2.5)
                var canhBao = KiemTraHopLe();
                if (canhBao.Count > 0)
                {
                    string msg = "Kiểm tra dự toán có một số điểm cần lưu ý trước khi xuất:\n\n"
                        + string.Join("\n", canhBao.Select(x => "• " + x))
                        + "\n\nBạn vẫn muốn tiếp tục xuất Excel?";
                    if (MessageBox.Show(msg, "Cảnh báo trước khi xuất", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                    {
                        return;
                    }
                }

                // Mở hộp thoại chọn các bảng biểu cần xuất
                using var dialog = new ChonBangXuatExcelDialog(isCheDoTMDT: LaCheDoTongMucDauTu);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                var opts = dialog.LuaChon;
                if (opts == null || !opts.CoItNhatMotBangDuocChon())
                {
                    MessageBox.Show("Vui lòng chọn ít nhất một bảng biểu để xuất Excel.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var loading = new LoadingForm("Đang xuất các bảng biểu ra Excel...");
                loading.Show();
                System.Windows.Forms.Application.DoEvents();

                // Thực hiện xuất các bảng đã chọn
                _xuatService.XuatCacBangTheoTuyChon(wb, _duToan, opts);

                loading.Close();
                loading.Dispose();

                // Đối chiếu số liệu giữa các sheet sau khi xuất (2.2)
                var lech = KiemTraDoiChieuBieuMau(wb);
                if (lech.Count > 0)
                {
                    MessageBox.Show(
                        "Sau khi xuất, phát hiện chênh lệch giữa các bảng biểu:\n\n" + string.Join("\n", lech.Select(x => "• " + x)) +
                        "\n\nBạn nên kiểm tra lại dữ liệu nguồn trước khi dùng kết quả.",
                        "Đối chiếu bảng biểu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                var dsDaXuat = opts.GetDanhSachDaChon().Select(x => $"- {x}").ToList();
                MessageBox.Show($"Đã xuất thành công các bảng biểu sang Excel:\n{string.Join("\n", dsDaXuat)}", "Xuất Excel Thành Công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xuất Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool _isFormattingText = false;

        /// <summary>
        /// Yêu cầu 6: Định dạng số thời gian thực chuẩn Việt Nam (phân cách hàng nghìn bằng dấu chấm) khi người dùng gõ phím
        /// Giữ nguyên vị trí con trỏ chuột không bị nhảy về cuối ô.
        /// </summary>
        private void FormatLiveCurrency(TextBox tb)
        {
            if (_isUpdating || _isFormattingText || string.IsNullOrWhiteSpace(tb.Text)) return;

            try
            {
                _isFormattingText = true;
                string raw = tb.Text;
                int selStart = tb.SelectionStart;
                int rightOffset = raw.Length - selStart;

                string integerPart = raw;
                string decimalPart = "";
                int commaIndex = raw.IndexOf(',');
                if (commaIndex >= 0)
                {
                    integerPart = raw.Substring(0, commaIndex);
                    decimalPart = raw.Substring(commaIndex);
                    string decDigits = new string(decimalPart.Skip(1).Where(char.IsDigit).ToArray());
                    decimalPart = "," + decDigits;
                }

                string cleanInt = new string(integerPart.Where(char.IsDigit).ToArray());
                if (string.IsNullOrEmpty(cleanInt))
                {
                    if (commaIndex >= 0)
                    {
                        tb.Text = "0" + decimalPart;
                        tb.SelectionStart = Math.Min(tb.Text.Length, 1);
                    }
                    return;
                }

                if (decimal.TryParse(cleanInt, out decimal intVal))
                {
                    string formattedInt = intVal.ToString("#,##0", UIHelper.ViCulture);
                    string formattedTotal = formattedInt + decimalPart;

                    if (tb.Text != formattedTotal)
                    {
                        tb.Text = formattedTotal;
                        int newPos = formattedTotal.Length - rightOffset;
                        if (newPos < 0) newPos = 0;
                        if (newPos > formattedTotal.Length) newPos = formattedTotal.Length;
                        tb.SelectionStart = newPos;
                    }
                }
            }
            finally
            {
                _isFormattingText = false;
            }
        }

        private void OnMoneyKeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsControl(e.KeyChar)) return;

            // Chuyển đổi phím chấm (numpad) thành phẩy theo chuẩn Việt Nam
            if (e.KeyChar == '.')
            {
                e.KeyChar = ',';
            }

            if (e.KeyChar == ',')
            {
                if (sender is TextBox tb && tb.Text.Contains(","))
                {
                    e.Handled = true; // Chỉ cho phép 1 dấu phẩy
                }
                return;
            }

            if (!char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Chỉ cho phép số
            }
        }

        private void OnMoneyKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                CapNhatGiaTriDauVao();
                e.SuppressKeyPress = true;
            }
        }
    }
}
