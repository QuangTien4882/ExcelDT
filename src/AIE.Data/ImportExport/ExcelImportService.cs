using ClosedXML.Excel;
using AIE.Core.Models;
using AIE.Core.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AIE.Data.ImportExport
{
    /// <summary>
    /// Kết quả import từ file Excel.
    /// </summary>
    public class ImportResult
    {
        public int SoLuongThanhCong { get; set; }
        public int SoLuongLoi { get; set; }
        public int SoLuongMoi { get; set; }
        public int SoLuongCapNhat { get; set; }
        public List<string> DanhSachLoi { get; set; } = new List<string>();

        public string TomTat
        {
            get
            {
                var msg = $"Thành công: {SoLuongThanhCong} dòng (Mới: {SoLuongMoi}, Cập nhật: {SoLuongCapNhat}).";
                if (SoLuongLoi > 0)
                    msg += $"\nLỗi: {SoLuongLoi} dòng.";
                return msg;
            }
        }
    }

    /// <summary>
    /// Service đọc dữ liệu từ file Excel (Templates).
    /// </summary>
    public class ExcelImportService
    {
        /// <summary>
        /// Import file 1_DinhMucCongTac.xlsx (2 sheet: CongTac + HaoPhi)
        /// Trả về danh sách CongTacXayDung đã gắn HaoPhi.
        /// </summary>
        public List<CongTacXayDung> ImportDinhMucCongTac(string filePath, out ImportResult result)
        {
            result = new ImportResult();
            var congTacDict = new Dictionary<string, CongTacXayDung>();

            using (var wb = new XLWorkbook(filePath))
            {
                // Sheet 1: CongTac
                var wsCongTac = wb.Worksheet("CongTac");
                if (wsCongTac == null)
                {
                    result.DanhSachLoi.Add("Không tìm thấy Sheet 'CongTac'.");
                    return new List<CongTacXayDung>();
                }

                foreach (var row in wsCongTac.RowsUsed().Skip(1))
                {
                    try
                    {
                        var maHieu = row.Cell(1).GetString().Trim();
                        if (string.IsNullOrWhiteSpace(maHieu)) continue;

                        congTacDict[maHieu] = new CongTacXayDung
                        {
                            MaHieu = maHieu,
                            TenCongTac = row.Cell(2).GetString().Trim(),
                            DonVi = row.Cell(3).GetString().Trim(),
                            DanhSachHaoPhi = new List<HaoPhi>()
                        };
                        result.SoLuongThanhCong++;
                    }
                    catch (Exception ex)
                    {
                        result.SoLuongLoi++;
                        result.DanhSachLoi.Add("CongTac dòng " + row.RowNumber() + ": " + ex.Message);
                    }
                }

                // Sheet 2: HaoPhi
                var wsHaoPhi = wb.Worksheet("HaoPhi");
                if (wsHaoPhi == null)
                {
                    result.DanhSachLoi.Add("Không tìm thấy Sheet 'HaoPhi'.");
                    return congTacDict.Values.ToList();
                }

                foreach (var row in wsHaoPhi.RowsUsed().Skip(1))
                {
                    try
                    {
                        var maHieuCT = row.Cell(1).GetString().Trim();
                        if (string.IsNullOrWhiteSpace(maHieuCT)) continue;

                        var loaiStr = row.Cell(2).GetString().Trim().ToUpper();
                        LoaiHaoPhi loai;
                        switch (loaiStr)
                        {
                            case "VL": loai = LoaiHaoPhi.VL; break;
                            case "NC": loai = LoaiHaoPhi.NC; break;
                            case "MAY": loai = LoaiHaoPhi.MAY; break;
                            default:
                                result.SoLuongLoi++;
                                result.DanhSachLoi.Add("HaoPhi dòng " + row.RowNumber() + ": LoaiHaoPhi '" + loaiStr + "' không hợp lệ.");
                                continue;
                        }

                        var hp = new HaoPhi
                        {
                            LoaiHaoPhi = loai,
                            MaHieuHP = row.Cell(3).GetString().Trim(),
                            TenHaoPhi = row.Cell(4).GetString().Trim(),
                            DonVi = row.Cell(5).GetString().Trim(),
                            DinhMuc = row.Cell(6).GetValue<decimal>()
                        };

                        if (congTacDict.ContainsKey(maHieuCT))
                        {
                            congTacDict[maHieuCT].DanhSachHaoPhi.Add(hp);
                        }
                        else
                        {
                            result.DanhSachLoi.Add("HaoPhi dòng " + row.RowNumber() + ": Mã hiệu CT '" + maHieuCT + "' không tồn tại trong Sheet CongTac.");
                            result.SoLuongLoi++;
                        }
                    }
                    catch (Exception ex)
                    {
                        result.SoLuongLoi++;
                        result.DanhSachLoi.Add("HaoPhi dòng " + row.RowNumber() + ": " + ex.Message);
                    }
                }
            }

            return congTacDict.Values.ToList();
        }

        public List<VatLieu> ImportGiaVatLieu(string filePath, out ImportResult result)
        {
            result = new ImportResult();
            var data = new List<VatLieu>();

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet(1);
                foreach (var row in ws.RowsUsed().Skip(1))
                {
                    try
                    {
                        var maVL = row.Cell(1).GetString().Trim();
                        if (string.IsNullOrWhiteSpace(maVL)) continue;

                        data.Add(new VatLieu
                        {
                            MaVL = maVL,
                            TenVL = row.Cell(2).GetString().Trim(),
                            DonVi = row.Cell(3).GetString().Trim(),
                            DonGia = row.Cell(4).TryGetValue<decimal>(out var d) ? d : 0,
                            GhiChu = row.Cell(5).GetString().Trim(),
                            NgayCapNhat = DateTime.Now
                        });
                        result.SoLuongThanhCong++;
                    }
                    catch (Exception ex)
                    {
                        result.SoLuongLoi++;
                        result.DanhSachLoi.Add("Dòng " + row.RowNumber() + ": " + ex.Message);
                    }
                }
            }
            return data;
        }

        public List<NhanCong> ImportGiaNhanCong(string filePath, out ImportResult result)
        {
            result = new ImportResult();
            var data = new List<NhanCong>();

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet(1);

                // Có 2 kiểu file mẫu nhân công:
                //  - File mẫu "3_GiaNhanCong.xlsx": MaNC, TenNC, Nhom, DonVi, DonGia, GhiChu (6 cột)
                //  - File xuất đơn giá:           MaNC, Tên,  DonVi, DonGia, GhiChu        (5 cột)
                // Nhận diện theo tiêu đề cột 3 để tránh đọc lệch cột.
                bool coCotNhom = LaTieuDe(ws.Cell(1, 3).GetString(), "Nhom", "Nhóm");

                int colTen = 2;
                int colNhom = coCotNhom ? 3 : -1;
                int colDonVi = coCotNhom ? 4 : 3;
                int colDonGia = coCotNhom ? 5 : 4;
                int colGhiChu = coCotNhom ? 6 : 5;

                foreach (var row in ws.RowsUsed().Skip(1))
                {
                    try
                    {
                        var maNC = row.Cell(1).GetString().Trim();
                        if (string.IsNullOrWhiteSpace(maNC)) continue;

                        data.Add(new NhanCong
                        {
                            MaNC = maNC,
                            TenNC = row.Cell(colTen).GetString().Trim(),
                            Nhom = colNhom > 0 && row.Cell(colNhom).TryGetValue<int>(out var nhom) ? nhom : 0,
                            DonVi = row.Cell(colDonVi).GetString().Trim(),
                            DonGia = row.Cell(colDonGia).TryGetValue<decimal>(out var d) ? d : 0,
                            GhiChu = row.Cell(colGhiChu).GetString().Trim(),
                            NgayCapNhat = DateTime.Now
                        });
                        result.SoLuongThanhCong++;
                    }
                    catch (Exception ex)
                    {
                        result.SoLuongLoi++;
                        result.DanhSachLoi.Add("Dòng " + row.RowNumber() + ": " + ex.Message);
                    }
                }
            }
            return data;
        }

        public List<MayThiCong> ImportGiaMayThiCong(string filePath, out ImportResult result)
        {
            result = new ImportResult();
            var data = new List<MayThiCong>();

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet(1);
                foreach (var row in ws.RowsUsed().Skip(1))
                {
                    try
                    {
                        var maMay = row.Cell(1).GetString().Trim();
                        if (string.IsNullOrWhiteSpace(maMay)) continue;

                        data.Add(new MayThiCong
                        {
                            MaMay = maMay,
                            TenMay = row.Cell(2).GetString().Trim(),
                            DonVi = row.Cell(3).GetString().Trim(),
                            DonGia = row.Cell(4).TryGetValue<decimal>(out var d) ? d : 0,
                            GhiChu = row.Cell(5).GetString().Trim(),
                            NgayCapNhat = DateTime.Now
                        });
                        result.SoLuongThanhCong++;
                    }
                    catch (Exception ex)
                    {
                        result.SoLuongLoi++;
                        result.DanhSachLoi.Add("Dòng " + row.RowNumber() + ": " + ex.Message);
                    }
                }
            }
            return data;
        }

        /// <summary>So khớp tiêu đề cột (bỏ khoảng trắng, không phân biệt hoa thường/dấu).</summary>
        private static bool LaTieuDe(string tieuDe, params string[] ungVien)
        {
            var t = (tieuDe ?? string.Empty).Trim();
            foreach (var u in ungVien)
            {
                if (string.Equals(t, u, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }
    }
}
