using AIE.Core.Enums;
using AIE.Core.Models;
using AIE.ExcelAddIn.Helpers;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AIE.ExcelAddIn.Services;

public class DuToanExcelReader
{
    public List<CongTacThamDinh> Read(ThamDinhConfig config)
    {
        var app = (Application)ExcelDnaUtil.Application;
        var wb = app.ActiveWorkbook;
        Worksheet ws = null;
        foreach (Worksheet sheet in wb.Worksheets)
        {
            if (sheet.Name == config.SheetName)
            {
                ws = sheet;
                break;
            }
        }

        if (ws == null) throw new Exception($"Không tìm thấy sheet '{config.SheetName}' trong file đang mở.");

        // Tạo từ điển tra cứu khối lượng tự động từ các sheet khác trong cùng file Excel (VD: Giá tổng hợp, Khối lượng)
        var klMap = BuildKhoiLuongDictionary(wb, config.SheetName);

        var result = new List<CongTacThamDinh>();
        int row = config.DongBatDau;

        Range usedRange = ws.UsedRange;
        int maxRow = usedRange.Rows.Count + usedRange.Row - 1;

        int cMa = ColLetterToNumber(config.ColMaHieu);
        int cTen = ColLetterToNumber(config.ColTenCT);
        int cDonVi = ColLetterToNumber(config.ColDonVi);
        int cDinhMuc = ColLetterToNumber(config.ColDinhMuc);
        int cDonGia = ColLetterToNumber(config.ColDonGia);
        int cThanhTien = ColLetterToNumber(config.ColThanhTien);

        int maxCol = Math.Max(Math.Max(cMa, cTen), Math.Max(cDonVi, cDinhMuc));
        maxCol = Math.Max(maxCol, Math.Max(cDonGia, cThanhTien));
        maxCol = Math.Max(maxCol, usedRange.Columns.Count + usedRange.Column - 1);

        var reader = new FastRangeReader(ws, 1, 1, maxRow, Math.Max(maxCol, 12));

        Func<int, int, string> getFastVal = (r, c) =>
        {
            if (c <= 0) return string.Empty;
            return reader.GetString(r, c);
        };

        Func<int, int, decimal> getFastDec = (r, c) =>
        {
            if (c <= 0) return 0m;
            return reader.GetDecimal(r, c);
        };

        CongTacThamDinh currentCongTac = null;
        LoaiHaoPhi currentLoaiHp = LoaiHaoPhi.VL; // Mặc định là VL

        while (row <= maxRow)
        {
            string maHieu = getFastVal(row, cMa);
            string ten = getFastVal(row, cTen);
            string donVi = getFastVal(row, cDonVi);
            decimal valDinhMuc = getFastDec(row, cDinhMuc);
            decimal valDonGia = cDonGia <= 0 ? 0 : getFastDec(row, cDonGia);
            decimal valThanhTien = cThanhTien <= 0 ? 0 : getFastDec(row, cThanhTien);

            string strDinhMucRaw = getFastVal(row, cDinhMuc);
            string strDonGiaRaw = cDonGia <= 0 ? "" : getFastVal(row, cDonGia);
            string strThanhTienRaw = cThanhTien <= 0 ? "" : getFastVal(row, cThanhTien);

            // Bỏ qua dòng trống hoàn toàn ở các cột quan trọng
            if (string.IsNullOrEmpty(maHieu) && string.IsNullOrEmpty(ten) && string.IsNullOrEmpty(strDinhMucRaw) && string.IsNullOrEmpty(strDonGiaRaw) && string.IsNullOrEmpty(strThanhTienRaw))
            {
                row++;
                continue;
            }

            if (!string.IsNullOrWhiteSpace(maHieu))
            {
                decimal kl = valDinhMuc;
                decimal dg = valDonGia;
                decimal tt = valThanhTien;

                // Tự động liên kết khối lượng từ sheet dự toán / giá tổng hợp nếu sheet đơn giá chi tiết để trống KL
                if (kl <= 0 && klMap != null && klMap.TryGetValue(maHieu.Trim(), out decimal foundKl) && foundKl > 0)
                {
                    kl = foundKl;
                }

                currentCongTac = new CongTacThamDinh
                {
                    SoDongExcel = row,
                    MaHieu = maHieu.Trim(),
                    TenCongTac = ten,
                    DonVi = donVi,
                    KhoiLuong = kl,
                    DonGia = dg,
                    ThanhTien = tt
                };
                result.Add(currentCongTac);
            }
            else if (currentCongTac != null)
            {
                // Kiểm tra xem dòng này có phải là dòng tiêu đề nhóm không (Vật liệu, Nhân công, Máy thi công)
                string tenLower = ten.ToLower().Trim();
                
                // Nếu dòng này không có định mức số (hoặc rỗng), có khả năng cao là dòng tiêu đề nhóm
                if (string.IsNullOrEmpty(strDinhMucRaw) || valDinhMuc == 0)
                {
                    if (tenLower.Contains("vật liệu") || tenLower == "vl" || tenLower.StartsWith("a)") || tenLower.StartsWith("a."))
                        currentLoaiHp = LoaiHaoPhi.VL;
                    else if (tenLower.Contains("nhân công") || tenLower == "nc" || tenLower.StartsWith("b)") || tenLower.StartsWith("b."))
                        currentLoaiHp = LoaiHaoPhi.NC;
                    else if (tenLower.Contains("máy thi công") || tenLower == "m" || tenLower.Contains("máy tc") || tenLower.Contains("máy") || tenLower.StartsWith("c)") || tenLower.StartsWith("c."))
                        currentLoaiHp = LoaiHaoPhi.MAY;
                }
                else
                {
                    // Nếu là dòng dữ liệu hao phí (có định mức > 0)
                    if (tenLower.StartsWith("nhân công") || tenLower.Contains("thợ bậc") || tenLower.Contains("nhóm"))
                        currentLoaiHp = LoaiHaoPhi.NC;
                    else if (tenLower.StartsWith("máy") || tenLower.Contains("ca máy") || tenLower.Contains("công suất"))
                        currentLoaiHp = LoaiHaoPhi.MAY;

                    // Thêm HaoPhi
                    var hp = new HaoPhiThamDinh
                    {
                        SoDongExcel = row,
                        Loai = currentLoaiHp,
                        TenHaoPhi = ten,
                        DonVi = donVi,
                        DinhMuc = valDinhMuc,
                        DonGia = valDonGia,
                        ThanhTien = valThanhTien > 0 ? valThanhTien : (valDinhMuc * valDonGia)
                    };
                    currentCongTac.DanhSachHaoPhi.Add(hp);
                }
            }

            row++;
        }

        // HẬU XỬ LÝ (POST-PROCESSING): Đảm bảo 100% công tác có đầy đủ Khối lượng, Đơn giá và Thành tiền
        foreach (var ct in result)
        {
            // 1. Khối lượng
            if (ct.KhoiLuong <= 0)
            {
                string cleanMa = ct.MaHieu.Trim().ToUpper();
                if (klMap != null)
                {
                    if (klMap.TryGetValue(cleanMa, out decimal foundKl) && foundKl > 0)
                    {
                        ct.KhoiLuong = foundKl;
                    }
                    else if (!string.IsNullOrEmpty(ct.TenCongTac))
                    {
                        string nameKey = "NAME_" + BoDau(ct.TenCongTac);
                        if (klMap.TryGetValue(nameKey, out decimal foundKlName) && foundKlName > 0)
                        {
                            ct.KhoiLuong = foundKlName;
                        }
                    }
                }
                
                if (ct.KhoiLuong <= 0)
                {
                    ct.KhoiLuong = 1.0m;
                }
            }

            // 2. Thành tiền & Đơn giá
            if (ct.ThanhTien <= 0 && ct.DanhSachHaoPhi.Count > 0)
            {
                decimal sumHp = ct.DanhSachHaoPhi.Sum(h => h.ThanhTien > 0 ? h.ThanhTien : h.DinhMuc * h.DonGia);
                ct.ThanhTien = (ct.KhoiLuong > 0) ? (sumHp * ct.KhoiLuong) : sumHp;
            }

            if (ct.DonGia <= 0 && ct.ThanhTien > 0 && ct.KhoiLuong > 0)
            {
                ct.DonGia = Math.Round(ct.ThanhTien / ct.KhoiLuong, 0);
            }
            else if (ct.ThanhTien <= 0 && ct.DonGia > 0 && ct.KhoiLuong > 0)
            {
                ct.ThanhTien = Math.Round(ct.DonGia * ct.KhoiLuong, 0);
            }
        }

        return result;
    }

    private string GetCellValue(Worksheet ws, int row, string col)
    {
        if (string.IsNullOrEmpty(col)) return string.Empty;
        try
        {
            Range range = ws.Range[$"{col}{row}"];
            if (range.Value2 == null) return string.Empty;
            return range.Value2.ToString().Trim();
        }
        catch
        {
            return string.Empty;
        }
    }

    private string GetCellValue(Worksheet ws, int row, int col)
    {
        if (col <= 0) return string.Empty;
        try
        {
            Range range = ws.Cells[row, col];
            return range?.Value2?.ToString()?.Trim() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private decimal GetCellDecimal(Worksheet ws, int row, string col)
    {
        if (string.IsNullOrEmpty(col)) return 0;
        try
        {
            Range range = ws.Range[$"{col}{row}"];
            return ParseCellDecimalValue(range?.Value2);
        }
        catch
        {
            return 0;
        }
    }

    public decimal GetCellDecimal(Worksheet ws, int row, int col)
    {
        if (col <= 0) return 0;
        try
        {
            Range range = ws.Cells[row, col];
            return ParseCellDecimalValue(range?.Value2);
        }
        catch
        {
            return 0;
        }
    }

    private decimal ParseCellDecimalValue(object v)
    {
        if (v == null) return 0;
        if (v is double d) return (decimal)d;
        if (v is float f) return (decimal)f;
        if (v is int i) return (decimal)i;
        if (v is long l) return (decimal)l;
        if (v is decimal m) return m;

        string s = v.ToString().Trim();
        if (string.IsNullOrEmpty(s)) return 0;

        // Nhận dạng giá trị phần trăm "5%", "4,5%" → chia 100
        bool laPhanTram = s.EndsWith("%");
        if (laPhanTram) s = s.TrimEnd('%').Trim();
        if (s == "") return 0;

        if (decimal.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal vInv))
            return laPhanTram ? vInv / 100m : vInv;
        if (decimal.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out decimal vCur))
            return laPhanTram ? vCur / 100m : vCur;

        // Xử lý dấu chấm / phẩy tiếng Việt
        s = s.Replace(" ", "");
        if (s.Contains(",") && s.Contains("."))
        {
            if (s.LastIndexOf(',') > s.LastIndexOf('.'))
                s = s.Replace(".", "").Replace(",", ".");
            else
                s = s.Replace(",", "");
        }
        else if (s.Contains(","))
        {
            s = s.Replace(",", ".");
        }

        if (decimal.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal vClean))
            return laPhanTram ? vClean / 100m : vClean;

        return 0;
    }

    private static string BoDau(string text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        string normalized = text.Normalize(System.Text.NormalizationForm.FormD);
        var sb = new System.Text.StringBuilder();
        foreach (char c in normalized)
        {
            var uc = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c);
            if (uc != System.Globalization.UnicodeCategory.NonSpacingMark)
                sb.Append(c);
        }
        return sb.ToString().Normalize(System.Text.NormalizationForm.FormC).ToLower().Trim();
    }

    private Dictionary<string, decimal> BuildKhoiLuongDictionary(Workbook wb, string currentSheetName)
    {
        var dict = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        if (wb == null) return dict;

        try
        {
            // Duyệt qua TẤT CẢ các sheet trong Workbook (ưu tiên các sheet dự toán, giá tổng hợp, khối lượng trước)
            var sheets = new List<Worksheet>();
            foreach (Worksheet sheet in wb.Worksheets)
            {
                if (sheet.Name == currentSheetName) continue;
                string sClean = BoDau(sheet.Name);
                if (sClean.Contains("du toan") || sClean.Contains("dutoan") || sClean.Contains("dt") ||
                    sClean.Contains("tong hop") || sClean.Contains("gia th") || sClean.Contains("khoi luong") ||
                    sClean.Contains("kl") || sClean.Contains("cpxd") || sClean.Contains("xay lap"))
                {
                    sheets.Insert(0, sheet); // Ưu tiên hàng đầu
                }
                else
                {
                    sheets.Add(sheet);
                }
            }

            foreach (var ws in sheets)
            {
                try
                {
                    Range u = ws.UsedRange;
                    if (u == null) continue;
                    int rCount = u.Rows.Count;
                    int cCount = Math.Min(u.Columns.Count, 30);
                    int startRow = u.Row;
                    int startCol = u.Column;

                    var reader = new FastRangeReader(ws, startRow, startCol, startRow + rCount - 1, startCol + cCount - 1);
                    if (!reader.HasData) continue;

                    Func<int, int, object> getMatVal = (r, c) => reader.GetValue(r, c);

                    int colMa = -1;
                    int colKl = -1;
                    int colTen = -1;
                    int headerRow = -1;

                    // Quét 25 dòng đầu để tìm dòng tiêu đề
                    for (int r = startRow; r < startRow + Math.Min(25, rCount); r++)
                    {
                        for (int c = startCol; c < startCol + cCount; c++)
                        {
                            string t1 = getMatVal(r, c)?.ToString() ?? "";
                            string t2 = (r + 1 <= startRow + rCount) ? (getMatVal(r + 1, c)?.ToString() ?? "") : "";
                            string combined = BoDau(t1 + " " + t2);

                            if (colMa == -1 && (combined.Contains("ma hieu") || combined.Contains("ma cv") || combined.Contains("ma so") || combined.Contains("ma dm") || combined == "ma" || combined.Contains("sh dm") || combined.Contains("ma dinh muc") || combined.Contains("so hieu")))
                            {
                                colMa = c;
                                headerRow = r;
                            }

                            if (colKl == -1 && (combined.Contains("khoi luong") || combined == "kl" || combined.Contains("so luong") || combined.Contains("thiet ke") || combined.Contains("toan bo")))
                            {
                                colKl = c;
                                headerRow = r;
                            }

                            if (colTen == -1 && (combined.Contains("ten cong tac") || combined.Contains("noi dung") || combined.Contains("ten cv") || combined == "ten cong viec" || combined.Contains("danh muc")))
                            {
                                colTen = c;
                            }
                        }

                        if (colMa != -1 && colKl != -1) break;
                    }

                    // Nếu tìm thấy cột Mã hiệu và cột Khối lượng
                    if (colMa != -1 && colKl != -1 && headerRow != -1)
                    {
                        int scanStart = headerRow + ((getMatVal(headerRow + 1, colMa)?.ToString()?.Length ?? 0) < 3 ? 2 : 1);

                        for (int r = scanStart; r <= startRow + rCount - 1; r++)
                        {
                            string ma = getMatVal(r, colMa)?.ToString()?.Trim() ?? "";
                            string ten = colTen != -1 ? (getMatVal(r, colTen)?.ToString()?.Trim() ?? "") : "";

                            if (!string.IsNullOrEmpty(ma) && ma.Length >= 2)
                            {
                                object v = getMatVal(r, colKl);
                                decimal kl = ParseCellDecimalValue(v);

                                if (kl > 0)
                                {
                                    string cleanMa = ma.Trim().ToUpper();
                                    if (!dict.ContainsKey(cleanMa))
                                    {
                                        dict[cleanMa] = kl;
                                    }

                                    if (!string.IsNullOrEmpty(ten))
                                    {
                                        string nameKey = "NAME_" + BoDau(ten);
                                        if (!dict.ContainsKey(nameKey))
                                        {
                                            dict[nameKey] = kl;
                                        }
                                    }
                                }
                            }
                        }

                        // Nếu đã quét được dữ liệu khối lượng từ sheet này thì tiếp tục bổ sung thêm từ các sheet khác nếu còn thiếu
                    }
                }
                catch { }
            }
        }
        catch { }
        return dict;
    }

    /// <summary>
    /// Đọc Bảng Dự toán chi tiết / Bảng Giá tổng hợp từ Workbook
    /// Trả về tên sheet đã đọc và danh sách các Hạng mục cùng các công tác tiên lượng
    /// </summary>
    public (string SheetName, List<HangMuc> DanhSachHangMuc) ReadBangDuToanChiTiet(Workbook? wb = null, string? sheetName = null)
    {
        if (wb == null)
        {
            var app = (Application)ExcelDnaUtil.Application;
            wb = app.ActiveWorkbook;
        }

        if (wb == null) throw new Exception("Không tìm thấy Workbook nào đang mở trong Excel.");

        Worksheet? ws = null;
        if (!string.IsNullOrEmpty(sheetName))
        {
            foreach (Worksheet s in wb.Worksheets)
            {
                if (string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    ws = s;
                    break;
                }
            }
        }

        // Nếu chưa chỉ định hoặc không tìm thấy, tự động phát hiện sheet phù hợp nhất
        if (ws == null)
        {
            var candidates = new List<(Worksheet Sheet, int Score)>();
            foreach (Worksheet s in wb.Worksheets)
            {
                string sClean = BoDau(s.Name);
                if (sClean.Contains("don gia chi tiet") || sClean.Contains("dgct") || sClean.Contains("phan tich") || sClean.Contains("ptdg") || sClean.Contains("camay") || sClean.Contains("luongnc") || sClean.Contains("thuyet minh") || sClean.Contains("sheet1"))
                    continue;

                int score = 0;
                if (sClean == "gia tong hop" || sClean == "giath" || sClean == "gth") score += 100;
                else if (sClean == "du toan" || sClean == "dutoan" || sClean == "dt") score += 90;
                else if (sClean.Contains("tong hop") || sClean.Contains("du toan") || sClean.Contains("chi tiet")) score += 70;
                else if (sClean.Contains("khoi luong") || sClean == "kl") score += 50;

                if (score > 0)
                {
                    candidates.Add((s, score));
                }
            }

            if (candidates.Count > 0)
            {
                ws = candidates.OrderByDescending(x => x.Score).First().Sheet;
            }
            else
            {
                ws = wb.ActiveSheet as Worksheet;
            }
        }

        if (ws == null) throw new Exception("Không tìm thấy Sheet Dự toán chi tiết / Giá tổng hợp phù hợp trong file.");

        var danhSachHangMuc = new List<HangMuc>();
        Range u = ws.UsedRange;
        if (u == null) return (ws.Name, danhSachHangMuc);

        int maxRow = u.Rows.Count + u.Row - 1;
        int maxCol = Math.Min(u.Columns.Count + u.Column - 1, 30);
        int startRow = u.Row;
        int startCol = u.Column;

        var reader = new FastRangeReader(ws, startRow, startCol, maxRow, maxCol);
        if (!reader.HasData) return (ws.Name, danhSachHangMuc);

        Func<int, int, object> getMatVal = (r, c) => reader.GetValue(r, c);
        Func<int, int, decimal> getMatDec = (r, c) => (c <= 0) ? 0m : reader.GetDecimal(r, c);

        // Quét tìm dòng tiêu đề
        int headerRow = -1;
        int colSTT = -1;
        int colMa = -1;
        int colTen = -1;
        int colDonVi = -1;
        int colKL = -1;
        int colDgVL = -1, colDgNC = -1, colDgMay = -1;
        int colTtVL = -1, colTtNC = -1, colTtMay = -1;

        for (int r = startRow; r <= startRow + Math.Min(25, u.Rows.Count); r++)
        {
            for (int c = startCol; c <= maxCol; c++)
            {
                string t1 = getMatVal(r, c)?.ToString() ?? "";
                string t2 = (r + 1 <= maxRow) ? (getMatVal(r + 1, c)?.ToString() ?? "") : "";
                string combined = BoDau(t1 + " " + t2);

                if (colSTT == -1 && (combined == "stt" || combined.StartsWith("stt ") || combined == "tt"))
                {
                    colSTT = c;
                    headerRow = r;
                }
                if (colMa == -1 && (combined.Contains("ma so") || combined.Contains("ma hieu") || combined.Contains("ma cv") || combined == "ma" || combined.Contains("sh dm") || combined.Contains("ma dinh muc")))
                {
                    colMa = c;
                    headerRow = r;
                }
                if (colTen == -1 && (combined.Contains("ten cong tac") || combined.Contains("ten cong viec") || combined.Contains("ten cv") || combined.Contains("noi dung cong viec") || combined.Contains("danh muc cong viec")))
                {
                    colTen = c;
                    headerRow = r;
                }
                if (colDonVi == -1 && (combined.Contains("don vi") || combined == "dvt" || combined.Contains("dvt")))
                {
                    colDonVi = c;
                }
                if (colKL == -1 && (combined.Contains("khoi luong") || combined == "kl" || combined.Contains("so luong")))
                {
                    colKL = c;
                    headerRow = r;
                }
            }

            if (colMa != -1 && colKL != -1) break;
        }

        // Nếu không tự nhận diện được cột chuẩn, dùng mặc định: STT=1, Mã=2, Tên=3, ĐVT=4, KL=5, ĐG VL=6, NC=7, Máy=8, TT VL=9, NC=10, Máy=11
        if (colMa == -1) colMa = 2;
        if (colTen == -1) colTen = 3;
        if (colDonVi == -1) colDonVi = 4;
        if (colKL == -1) colKL = 5;
        if (colSTT == -1) colSTT = 1;

        // Dò các cột Đơn giá và Thành tiền dựa vào dòng header
        if (headerRow != -1)
        {
            for (int c = colKL + 1; c <= maxCol; c++)
            {
                string t1 = getMatVal(headerRow, c)?.ToString() ?? "";
                string t2 = (headerRow + 1 <= maxRow) ? (getMatVal(headerRow + 1, c)?.ToString() ?? "") : "";
                string combined = BoDau(t1 + " " + t2);

                if (colDgVL == -1 && (combined.Contains("don gia vat lieu") || (combined.Contains("vat lieu") && c < colKL + 5))) colDgVL = c;
                else if (colDgNC == -1 && (combined.Contains("don gia nhan cong") || (combined.Contains("nhan cong") && c < colKL + 5))) colDgNC = c;
                else if (colDgMay == -1 && (combined.Contains("don gia may") || (combined.Contains("may") && c < colKL + 5))) colDgMay = c;
                else if (colTtVL == -1 && (combined.Contains("thanh tien vat lieu") || (combined.Contains("vat lieu") && c >= colKL + 4))) colTtVL = c;
                else if (colTtNC == -1 && (combined.Contains("thanh tien nhan cong") || (combined.Contains("nhan cong") && c >= colKL + 4))) colTtNC = c;
                else if (colTtMay == -1 && (combined.Contains("thanh tien may") || (combined.Contains("may") && c >= colKL + 4))) colTtMay = c;
            }
        }

        // Fallback cột đơn giá & thành tiền nếu không phát hiện được
        if (colDgVL == -1) colDgVL = colKL + 1;
        if (colDgNC == -1) colDgNC = colKL + 2;
        if (colDgMay == -1) colDgMay = colKL + 3;
        if (colTtVL == -1) colTtVL = colKL + 4;
        if (colTtNC == -1) colTtNC = colKL + 5;
        if (colTtMay == -1) colTtMay = colKL + 6;

        int rowStart = (headerRow != -1 ? headerRow + 1 : 6);
        // Kiểm tra nếu dòng rowStart vẫn là dòng sub-header
        string checkSub = getMatVal(rowStart, colTen)?.ToString() ?? "";
        if (checkSub.Contains("(1)") || checkSub.Contains("(2)") || checkSub.Contains("(3)") || BoDau(checkSub).Contains("ten cong tac"))
        {
            rowStart++;
        }

        HangMuc? currentHM = null;
        int hmCounter = 0;
        int ctCounter = 0;

        for (int r = rowStart; r <= maxRow; r++)
        {
            string sttRaw = getMatVal(r, colSTT)?.ToString()?.Trim() ?? "";
            string maHieu = getMatVal(r, colMa)?.ToString()?.Trim() ?? "";
            string ten = getMatVal(r, colTen)?.ToString()?.Trim() ?? "";
            string donVi = (colDonVi > 0) ? (getMatVal(r, colDonVi)?.ToString()?.Trim() ?? "") : "";
            decimal kl = (colKL > 0) ? getMatDec(r, colKL) : 0m;

            if (string.IsNullOrWhiteSpace(sttRaw) && string.IsNullOrWhiteSpace(maHieu) && string.IsNullOrWhiteSpace(ten))
                continue;

            string tenClean = BoDau(ten);
            if (tenClean.StartsWith("tong cong") || tenClean.StartsWith("cong toan") || tenClean.StartsWith("tong hop chi phi") || tenClean.StartsWith("bang du toan"))
                continue;

            // Kiểm tra xem dòng này có phải là dòng HẠNG MỤC không
            bool isHM = false;
            if (LapDuToanExcelService.IsRomanNumeral(sttRaw)) isHM = true;
            else if (string.IsNullOrEmpty(maHieu) && (kl == 0) && (string.IsNullOrEmpty(donVi) || donVi.Length > 10))
            {
                if (tenClean.StartsWith("hang muc") || tenClean.StartsWith("tuyen ") || tenClean.StartsWith("he thong") || tenClean.StartsWith("phan ") || tenClean.StartsWith("cong trinh") || tenClean.StartsWith("goi thau"))
                {
                    isHM = true;
                }
                else if (sttRaw.Length <= 3 && !string.IsNullOrEmpty(ten) && !tenClean.Contains("cong tac") && !tenClean.Contains("thi cong"))
                {
                    isHM = true;
                }
            }

            if (isHM)
            {
                hmCounter++;
                currentHM = new HangMuc
                {
                    STT = hmCounter,
                    RowIndex = r,
                    MaHangMuc = $"HM_{hmCounter:D2}",
                    TenHangMuc = !string.IsNullOrEmpty(sttRaw) ? $"{sttRaw}. {ten}" : ten,
                    LoaiCongTrinh = LapDuToanExcelService.NhanDienLoaiCongTrinh(ten)
                };
                danhSachHangMuc.Add(currentHM);
                continue;
            }

            // Nếu là dòng CÔNG TÁC (có mã hiệu hoặc có tên và khối lượng > 0)
            if (!string.IsNullOrEmpty(maHieu) || (!string.IsNullOrEmpty(ten) && kl > 0))
            {
                if (currentHM == null)
                {
                    hmCounter++;
                    currentHM = new HangMuc
                    {
                        STT = hmCounter,
                        RowIndex = r,
                        MaHangMuc = $"HM_{hmCounter:D2}",
                        TenHangMuc = "(Chưa phân loại)",
                        LoaiCongTrinh = ""
                    };
                    danhSachHangMuc.Add(currentHM);
                }

                ctCounter++;
                decimal dgVL = (colDgVL > 0) ? getMatDec(r, colDgVL) : 0m;
                decimal dgNC = (colDgNC > 0) ? getMatDec(r, colDgNC) : 0m;
                decimal dgMay = (colDgMay > 0) ? getMatDec(r, colDgMay) : 0m;

                var dong = new DongDuToan
                {
                    STT = ctCounter,
                    MaHieu = maHieu,
                    TenCongTac = ten,
                    DonVi = donVi,
                    KhoiLuong = kl,
                    DonGiaVL = dgVL,
                    DonGiaNC = dgNC,
                    DonGiaMay = dgMay
                };

                currentHM.DanhSachCongTac.Add(dong);
            }
        }

        return (ws.Name, danhSachHangMuc);
    }

    public static int ColLetterToNumber(string letter)
    {
        if (string.IsNullOrWhiteSpace(letter)) return 0;
        int col = 0;
        letter = letter.Trim().ToUpperInvariant();
        foreach (char c in letter)
        {
            if (c >= 'A' && c <= 'Z')
            {
                col = col * 26 + (c - 'A' + 1);
            }
        }
        return col;
    }
}
