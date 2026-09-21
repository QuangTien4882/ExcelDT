using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using AIE.Core.Models;
using AIE.ExcelAddIn.Helpers;

namespace AIE.ExcelAddIn.Services;

/// <summary>
/// Service chuyên tương tác (đọc/ghi) với file Excel trong chế độ Lập Dự Toán.
/// </summary>
public class LapDuToanExcelService
{
    /// <summary>
    /// Đọc bảng tiên lượng (BOQ) từ Sheet hiện tại đang mở.
    /// Bắt đầu đọc từ dòng số 5 (Dưới header).
    /// </summary>
    public DuToan ReadBOQFromActiveSheet()
    {
        var app = (Application)ExcelDnaUtil.Application;
        var wb = app.ActiveWorkbook;
        Worksheet ws = null;
        if (wb != null)
        {
            foreach (Worksheet sheet in wb.Sheets)
            {
                if (sheet.Name.StartsWith("DuToan_") || sheet.Name.StartsWith("DuToan "))
                {
                    ws = sheet;
                    break;
                }
            }
        }
        if (ws == null) ws = app.ActiveSheet as Worksheet;
        if (ws == null) throw new Exception("Không có Sheet nào đang mở.");

        var duToan = new DuToan();
        
        // Đọc thông tin công trình từ dòng 1, 2, 3 (Merge cell)
        string r1Val = GetCellValue(ws, 1, 1).Trim();
        string r2Val = GetCellValue(ws, 2, 1).Trim();
        string r3Val = GetCellValue(ws, 3, 1).Trim();

        string tenDuAn = "";
        string diaDiem = "";

        if (r1Val.ToUpper().Contains("BẢNG DỰ TOÁN"))
        {
            tenDuAn = r2Val.Replace("TÊN DỰ ÁN:", "").Replace("Dự án:", "").Trim();
            diaDiem = r3Val.Replace("ĐỊA ĐIỂM XÂY DỰNG:", "").Replace("Địa điểm xây dựng:", "").Replace("ĐỊA ĐIỂM:", "").Replace("Địa điểm:", "").Trim();
        }
        else
        {
            tenDuAn = r1Val.Replace("TÊN DỰ ÁN:", "").Replace("Dự án:", "").Trim();
            diaDiem = r2Val.Replace("ĐỊA ĐIỂM XÂY DỰNG:", "").Replace("Địa điểm xây dựng:", "").Replace("ĐỊA ĐIỂM:", "").Replace("Địa điểm:", "").Trim();
        }
        
        duToan.TenCongTrinh = string.IsNullOrEmpty(tenDuAn) ? "Công trình mặc định" : tenDuAn;
        duToan.DiaDiem = string.IsNullOrEmpty(diaDiem) ? "Không xác định" : diaDiem;

        // Đọc Vùng áp dụng từ ô Z1 (cột 26), nếu không có thử đọc từ L1 (cột 12) hoặc J1 (cột 10)
        string vungStr = GetCellValue(ws, 1, 26);
        if (string.IsNullOrEmpty(vungStr)) vungStr = GetCellValue(ws, 1, 12);
        if (string.IsNullOrEmpty(vungStr)) vungStr = GetCellValue(ws, 1, 10);
        if (int.TryParse(vungStr, out int vungInt) && System.Enum.IsDefined(typeof(AIE.Core.Enums.Vung), vungInt))
        {
            duToan.VungApDung = (AIE.Core.Enums.Vung)vungInt;
        }

        // Đảm bảo mở lại hiển thị cột 10 (J - Thành tiền Nhân công) nếu vô tình bị ẩn trước đó
        try
        {
            ((Range)ws.Columns[10]).Hidden = false;
        }
        catch { }

        Range usedRange = ws.UsedRange;
        int maxRow = usedRange.Rows.Count + usedRange.Row - 1;
        
        // Tìm dòng bắt đầu dữ liệu bằng cách tìm ô 'STT' ở các dòng đầu
        int startRow = 6;
        for (int r = 1; r <= 10; r++)
        {
            if (GetCellValue(ws, r, 1) == "STT")
            {
                startRow = r + 2;
                break;
            }
        }

        HangMuc currentHM = null;
        HangMucCon currentHMC = null;
        int currentHMIndex = 0;
        int currentHMCIndex = 0;

        int colCount = Math.Max(12, usedRange.Columns.Count);
        object[,] rawData = null;
        try
        {
            Range readRange = ws.Range[ws.Cells[1, 1], ws.Cells[maxRow, colCount]];
            rawData = readRange.Value2 as object[,];
        }
        catch { }

        string GetFastVal(int row, int col)
        {
            if (rawData != null && row >= 1 && row <= maxRow && col >= 1 && col <= colCount)
            {
                object v = rawData[row, col];
                return v?.ToString()?.Trim() ?? "";
            }
            return GetCellValue(ws, row, col).Trim();
        }

        decimal GetFastDec(int row, int col)
        {
            if (rawData != null && row >= 1 && row <= maxRow && col >= 1 && col <= colCount)
            {
                object val = rawData[row, col];
                if (val == null) return 0m;
                if (val is double d) return (decimal)d;
                if (val is decimal m) return m;
                if (val is int i) return i;
                if (val is long l) return l;
                string s = val.ToString().Trim();
                if (decimal.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal res))
                    return res;
                if (decimal.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out res))
                    return res;
                return 0m;
            }
            return GetCellDecimal(ws, row, col);
        }

        for (int r = startRow; r <= maxRow; r++)
        {
            string sttRaw = GetFastVal(r, 1); // Cột A
            string maHieu = GetFastVal(r, 2); // Cột B
            string ten = GetFastVal(r, 3);    // Cột C
            string donVi = GetFastVal(r, 4);  // Cột D
            decimal khoiLuong = GetFastDec(r, 5);  // Cột E
            decimal donGiaVL = GetFastDec(r, 6);   // Cột F
            decimal donGiaNC = GetFastDec(r, 7);   // Cột G
            decimal donGiaMay = GetFastDec(r, 8);  // Cột H

            // Dòng hoàn toàn rỗng -> bỏ qua
            if (string.IsNullOrWhiteSpace(sttRaw) && 
                string.IsNullOrWhiteSpace(maHieu) && 
                string.IsNullOrWhiteSpace(ten))
            {
                continue;
            }

            // Bỏ qua dòng TỔNG CỘNG hoặc dòng CỘNG cũ
            if (ten.StartsWith("TỔNG CỘNG", StringComparison.OrdinalIgnoreCase) ||
                ten.StartsWith("CỘNG TOÀN", StringComparison.OrdinalIgnoreCase) ||
                ten.Equals("TỔNG HỢP", StringComparison.OrdinalIgnoreCase) ||
                ten.StartsWith("CỘNG HẠNG MỤC", StringComparison.OrdinalIgnoreCase) ||
                ten.StartsWith("CỘNG PHẦN", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            // Phân biệt: Công tác vs Dòng Tiêu đề (Header)
            bool isWorkItem = !string.IsNullOrWhiteSpace(maHieu) || 
                              (!string.IsNullOrWhiteSpace(donVi) && khoiLuong > 0);

            if (!isWorkItem && !string.IsNullOrWhiteSpace(ten))
            {
                // Bỏ qua dòng Diễn giải đo bóc khối lượng hoặc ghi chú (không tạo Hạng mục rác)
                if (IsDienGiaiKhoiLuong(sttRaw, ten))
                {
                    continue;
                }

                // Dòng Tiêu đề
                if (IsLevel1Header(sttRaw, ten, currentHM))
                {
                    currentHMIndex++;
                    currentHM = new HangMuc
                    {
                        STT = currentHMIndex,
                        RowIndex = r,
                        MaHangMuc = $"HM_{currentHMIndex:D2}",
                        TenHangMuc = ten,
                        LoaiCongTrinh = NhanDienLoaiCongTrinh(ten)
                    };
                    duToan.DanhSachHangMuc.Add(currentHM);
                    currentHMC = null;
                    currentHMCIndex = 0;
                }
                else if (IsLevel2Header(sttRaw, ten))
                {
                    // Level 2: Hạng mục con thực sự
                    if (currentHM == null)
                    {
                        currentHMIndex++;
                        currentHM = new HangMuc
                        {
                            STT = currentHMIndex,
                            RowIndex = r,
                            MaHangMuc = $"HM_{currentHMIndex:D2}",
                            TenHangMuc = "(Chưa phân loại)",
                            LoaiCongTrinh = ""
                        };
                        duToan.DanhSachHangMuc.Add(currentHM);
                    }

                    currentHMCIndex++;
                    currentHMC = new HangMucCon
                    {
                        STT = currentHMCIndex,
                        RowIndex = r,
                        MaHangMucCon = $"{currentHM.MaHangMuc}_{currentHMCIndex:D2}",
                        TenHangMucCon = ten
                    };
                    currentHM.DanhSachHangMucCon.Add(currentHMC);
                }
                else
                {
                    // Các dòng text mô tả khác -> bỏ qua không tạo hạng mục con
                    continue;
                }
                continue;
            }

            if (!isWorkItem) continue;

            // Dòng Công tác
            if (currentHM == null)
            {
                currentHMIndex++;
                currentHM = new HangMuc
                {
                    STT = currentHMIndex,
                    RowIndex = startRow > 6 ? startRow - 1 : 5,
                    MaHangMuc = $"HM_{currentHMIndex:D2}",
                    TenHangMuc = "(Chưa phân loại)",
                    LoaiCongTrinh = ""
                };
                duToan.DanhSachHangMuc.Add(currentHM);
            }

            var dong = new DongDuToan
            {
                STT = r, // Dùng STT tạm bằng row để map ngược lại
                TenHangMucCon = currentHMC?.TenHangMucCon,
                MaHieu = maHieu,
                TenCongTac = ten,
                DonVi = donVi,
                KhoiLuong = khoiLuong,
                DonGiaVL = donGiaVL,
                DonGiaNC = donGiaNC,
                DonGiaMay = donGiaMay
            };

            currentHM.DanhSachCongTac.Add(dong);
            if (currentHMC != null)
            {
                currentHMC.DanhSachCongTac.Add(dong);
            }
        }

        // Đồng bộ Loại công trình mặc định của dự toán nếu chưa có
        if (duToan.DanhSachHangMuc.Count > 0 && string.IsNullOrEmpty(duToan.LoaiCongTrinh))
        {
            duToan.LoaiCongTrinh = duToan.DanhSachHangMuc[0].LoaiCongTrinh;
        }

        return duToan;
    }

    /// <summary>
    /// Gán đơn giá và công thức Thành tiền ngược lại các dòng tương ứng trên Excel,
    /// đồng thời đặt công thức tổng trực tiếp trên dòng tiêu đề Hạng mục và Hạng mục con.
    /// </summary>
    public void WriteDonGiaToExcel(DuToan duToan)
    {
        var app = (Application)ExcelDnaUtil.Application;
        var ws = app.ActiveSheet as Worksheet;
        if (ws == null) return;

        int maxR = 5;

        Action<DongDuToan> writeCongTac = (dong) =>
        {
            int r = dong.STT; // Do lúc đọc ta lưu SoDongExcel vào STT
            if (r < 6) return;
            if (r > maxR) maxR = r;

            // Gán Đơn Giá (Cột F, G, H) nếu ô chưa có công thức liên kết
            string fForm = ws.Cells[r, 6].Formula?.ToString() ?? "";
            if (!fForm.StartsWith("=", StringComparison.OrdinalIgnoreCase))
            {
                ws.Cells[r, 6].Value2 = dong.DonGiaVL;
            }
            string gForm = ws.Cells[r, 7].Formula?.ToString() ?? "";
            if (!gForm.StartsWith("=", StringComparison.OrdinalIgnoreCase))
            {
                ws.Cells[r, 7].Value2 = dong.DonGiaNC;
            }
            string hForm = ws.Cells[r, 8].Formula?.ToString() ?? "";
            if (!hForm.StartsWith("=", StringComparison.OrdinalIgnoreCase))
            {
                ws.Cells[r, 8].Value2 = dong.DonGiaMay;
            }

            // 3 Cột Thành tiền: Vật liệu (I), Nhân công (J), Máy thi công (K)
            ws.Cells[r, 9].Formula = $"=ROUND(E{r}*F{r}, 0)";
            ws.Cells[r, 10].Formula = $"=ROUND(E{r}*G{r}, 0)";
            ws.Cells[r, 11].Formula = $"=ROUND(E{r}*H{r}, 0)";
        };

        foreach (var hm in duToan.DanhSachHangMuc)
        {
            foreach (var dong in hm.DanhSachCongTac)
            {
                writeCongTac(dong);
            }

            if (hm.DanhSachHangMucCon != null && hm.DanhSachHangMucCon.Count > 0)
            {
                foreach (var hmc in hm.DanhSachHangMucCon)
                {
                    foreach (var dong in hmc.DanhSachCongTac)
                    {
                        writeCongTac(dong);
                    }
                }
            }

            // Gán công thức tổng ngay trên dòng tiêu đề Hạng mục con và Hạng mục cha
            if (hm.DanhSachHangMucCon != null && hm.DanhSachHangMucCon.Count > 0)
            {
                foreach (var hmc in hm.DanhSachHangMucCon)
                {
                    if (hmc.RowIndex > 0 && hmc.DanhSachCongTac.Count > 0)
                    {
                        int minR = hmc.DanhSachCongTac.Min(x => x.STT);
                        int maxR_hmc = hmc.DanhSachCongTac.Max(x => x.STT);
                        ws.Cells[hmc.RowIndex, 9].Formula = $"=SUM(I{minR}:I{maxR_hmc})";
                        ws.Cells[hmc.RowIndex, 10].Formula = $"=SUM(J{minR}:J{maxR_hmc})";
                        ws.Cells[hmc.RowIndex, 11].Formula = $"=SUM(K{minR}:K{maxR_hmc})";
                    }
                }

                if (hm.RowIndex > 0)
                {
                    var validHmc = hm.DanhSachHangMucCon.Where(c => c.RowIndex > 0).ToList();
                    if (validHmc.Count > 0)
                    {
                        ws.Cells[hm.RowIndex, 9].Formula = $"={string.Join("+", validHmc.Select(c => $"I{c.RowIndex}"))}";
                        ws.Cells[hm.RowIndex, 10].Formula = $"={string.Join("+", validHmc.Select(c => $"J{c.RowIndex}"))}";
                        ws.Cells[hm.RowIndex, 11].Formula = $"={string.Join("+", validHmc.Select(c => $"K{c.RowIndex}"))}";
                    }
                }
            }
            else if (hm.RowIndex > 0 && hm.DanhSachCongTac.Count > 0)
            {
                int minR = hm.DanhSachCongTac.Min(x => x.STT);
                int maxR_hm = hm.DanhSachCongTac.Max(x => x.STT);
                ws.Cells[hm.RowIndex, 9].Formula = $"=SUM(I{minR}:I{maxR_hm})";
                ws.Cells[hm.RowIndex, 10].Formula = $"=SUM(J{minR}:J{maxR_hm})";
                ws.Cells[hm.RowIndex, 11].Formula = $"=SUM(K{minR}:K{maxR_hm})";
            }
        }

        // Cập nhật dòng TỔNG CỘNG: Quét động theo số dòng thực tế trên Sheet
        Range usedRange = ws.UsedRange;
        int lastSheetRow = usedRange != null ? (usedRange.Rows.Count + usedRange.Row - 1) : maxR;
        int existingTotalRow = -1;
        for (int r = Math.Max(6, maxR - 2); r <= Math.Max(lastSheetRow, maxR + 5); r++)
        {
            string cText = ws.Cells[r, 3]?.Value2?.ToString()?.Trim() ?? "";
            if (cText.Equals("TỔNG CỘNG", StringComparison.OrdinalIgnoreCase) || cText.Equals("CỘNG", StringComparison.OrdinalIgnoreCase))
            {
                existingTotalRow = r;
                break;
            }
        }

        int totalRow = existingTotalRow > 0 ? existingTotalRow : (maxR + 1);
        ws.Cells[totalRow, 3].Value2 = "TỔNG CỘNG";

        var validHms = duToan.DanhSachHangMuc.Where(h => h.RowIndex > 0).ToList();
        if (validHms.Count > 1)
        {
            ws.Cells[totalRow, 9].Formula = $"={string.Join("+", validHms.Select(h => $"I{h.RowIndex}"))}";
            ws.Cells[totalRow, 10].Formula = $"={string.Join("+", validHms.Select(h => $"J{h.RowIndex}"))}";
            ws.Cells[totalRow, 11].Formula = $"={string.Join("+", validHms.Select(h => $"K{h.RowIndex}"))}";
        }
        else if (validHms.Count == 1)
        {
            ws.Cells[totalRow, 9].Formula = $"=I{validHms[0].RowIndex}";
            ws.Cells[totalRow, 10].Formula = $"=J{validHms[0].RowIndex}";
            ws.Cells[totalRow, 11].Formula = $"=K{validHms[0].RowIndex}";
        }
        else
        {
            // Quét động theo số dòng công tác thực tế, tránh cộng trùng các dòng tiêu đề hoặc diễn giải khối lượng
            var allCtRows = duToan.DanhSachHangMuc
                .SelectMany(h => h.DanhSachCongTac.Concat(h.DanhSachHangMucCon?.SelectMany(c => c.DanhSachCongTac) ?? Enumerable.Empty<DongDuToan>()))
                .Where(ct => ct.STT >= 6 && ct.STT <= maxR)
                .Select(ct => ct.STT)
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            if (allCtRows.Count > 0)
            {
                int minRow = allCtRows.Min();
                int maxRow = allCtRows.Max();
                if (maxRow - minRow + 1 == allCtRows.Count)
                {
                    ws.Cells[totalRow, 9].Formula = $"=SUM(I{minRow}:I{maxRow})";
                    ws.Cells[totalRow, 10].Formula = $"=SUM(J{minRow}:J{maxRow})";
                    ws.Cells[totalRow, 11].Formula = $"=SUM(K{minRow}:K{maxRow})";
                }
                else
                {
                    ws.Cells[totalRow, 9].Formula = $"={string.Join("+", allCtRows.Select(rIdx => $"I{rIdx}"))}";
                    ws.Cells[totalRow, 10].Formula = $"={string.Join("+", allCtRows.Select(rIdx => $"J{rIdx}"))}";
                    ws.Cells[totalRow, 11].Formula = $"={string.Join("+", allCtRows.Select(rIdx => $"K{rIdx}"))}";
                }
            }
            else
            {
                ws.Cells[totalRow, 9].Formula = $"=SUM(I6:I{maxR})";
                ws.Cells[totalRow, 10].Formula = $"=SUM(J6:J{maxR})";
                ws.Cells[totalRow, 11].Formula = $"=SUM(K6:K{maxR})";
            }
        }

        ws.Range[ws.Cells[totalRow, 1], ws.Cells[totalRow, 11]].Font.Bold = true;
        ws.Range[ws.Cells[totalRow, 1], ws.Cells[totalRow, 11]].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(200, 225, 250));

        // Kẻ khung toàn bộ bảng và định dạng số
        Range fullTable = ws.Range[ws.Cells[4, 1], ws.Cells[totalRow, 11]];
        fullTable.Borders.LineStyle = XlLineStyle.xlContinuous;
        ExcelFormatHelper.ApplyQuantityFormat(ws.Range[ws.Cells[6, 5], ws.Cells[totalRow, 5]], 2);
        ExcelFormatHelper.ApplyIntegerFormat(ws.Range[ws.Cells[6, 6], ws.Cells[totalRow, 11]]);
        try 
        { 
            ((Range)ws.Columns[10]).Hidden = false; 
            ((Range)ws.Columns[10]).ColumnWidth = 16; 
        } 
        catch { }
    }

    /// <summary>
    /// Ghi toàn bộ dữ liệu DuToan ra Sheet hiện tại (tái tạo bảng tiên lượng từ file .dt).
    /// </summary>
    public void WriteBOQToActiveSheet(DuToan duToan)
    {
        var app = (Application)ExcelDnaUtil.Application;
        
        // Nếu không có Workbook nào đang mở, tự động tạo mới
        if (app.Workbooks.Count == 0)
        {
            app.Workbooks.Add();
        }

        var ws = app.ActiveSheet as Worksheet;
        if (ws == null) throw new Exception("Không có Sheet nào đang mở.");

        app.ScreenUpdating = false;
        try
        {
            ws.Cells.Clear();

            // Set entire sheet font to Times New Roman, 12
            ws.Cells.Font.Name = "Times New Roman";
            ws.Cells.Font.Size = 12;

            // Dòng 1: Tiêu đề sheet
            Range r0 = ws.Range["A1", "K1"];
            r0.Merge();
            r0.Value2 = "BẢNG DỰ TOÁN CHI TIẾT";
            r0.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            r0.VerticalAlignment = XlVAlign.xlVAlignCenter;
            r0.Font.Bold = true;
            r0.Font.Size = 14;

            // Dòng 2: Dự án (đồng bộ với TongMucDauTu/TH_DuToan)
            string tenDuAn = UIHelper.ChuanHoaChuThuong(duToan.TenCongTrinh);
            if (string.IsNullOrEmpty(tenDuAn)) tenDuAn = "................................................................";
            Range r1 = ws.Range["A2", "K2"];
            r1.Merge();
            r1.Value2 = "Dự án: " + tenDuAn;
            r1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            r1.VerticalAlignment = XlVAlign.xlVAlignCenter;
            r1.Font.Bold = true;
            r1.Font.Size = 12;

            // Dòng 3: Địa điểm xây dựng (đồng bộ với TongMucDauTu/TH_DuToan)
            string diaDiem = UIHelper.ChuanHoaChuThuong(duToan.DiaDiem);
            if (string.IsNullOrEmpty(diaDiem)) diaDiem = "................................................................";
            Range r2 = ws.Range["A3", "K3"];
            r2.Merge();
            r2.Value2 = "Địa điểm xây dựng: " + diaDiem;
            r2.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            r2.VerticalAlignment = XlVAlign.xlVAlignCenter;
            r2.Font.Bold = true;
            r2.Font.Size = 12;

            // Dòng 4 & 5: Header 2 dòng chuẩn
            ws.Cells[4, 1] = "STT";
            ws.Cells[4, 2] = "Mã hiệu";
            ws.Cells[4, 3] = "Tên công tác";
            ws.Cells[4, 4] = "Đơn vị";
            ws.Cells[4, 5] = "Khối lượng";
            ws.Cells[4, 6] = "Đơn giá (đồng)";
            ws.Cells[4, 9] = "Thành tiền (đồng)";

            ws.Cells[5, 6] = "Vật liệu";
            ws.Cells[5, 7] = "Nhân công";
            ws.Cells[5, 8] = "Máy thi công";

            ws.Cells[5, 9] = "Vật liệu";
            ws.Cells[5, 10] = "Nhân công";
            ws.Cells[5, 11] = "Máy thi công";

            // Merge header cells
            ws.Range["A4:A5"].Merge();
            ws.Range["B4:B5"].Merge();
            ws.Range["C4:C5"].Merge();
            ws.Range["D4:D5"].Merge();
            ws.Range["E4:E5"].Merge();
            ws.Range["F4:H4"].Merge();
            ws.Range["I4:K4"].Merge();

            Range headerRange = ws.Range[ws.Cells[4, 1], ws.Cells[5, 11]];
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            headerRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(200, 220, 240));
            headerRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            // Column widths
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[1]).ColumnWidth = 5;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[2]).ColumnWidth = 12;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[3]).ColumnWidth = 42;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[4]).ColumnWidth = 8;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[5]).ColumnWidth = 12;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[6]).ColumnWidth = 14;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[7]).ColumnWidth = 14;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[8]).ColumnWidth = 14;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[9]).ColumnWidth = 16;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[10]).ColumnWidth = 16;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[11]).ColumnWidth = 16;

            if (ws.Name.StartsWith("Sheet"))
            {
                ws.Name = "DuToan_" + DateTime.Now.ToString("HHmmss");
            }

            // Dòng 6 trở đi: Dữ liệu đa hạng mục & hạng mục con
            int r = 6;
            int sttCongTac = 1;
            var hmRows = new List<int>();

            foreach (var hm in duToan.DanhSachHangMuc)
            {
                // Dòng Tiêu đề Hạng mục cha (Level 1)
                int hmRow = r;
                hm.RowIndex = hmRow;
                hmRows.Add(hmRow);

                ws.Cells[r, 1] = ToRomanNumeral(hm.STT);
                ws.Cells[r, 3] = hm.TenHangMuc.ToUpper();
                ws.Range[ws.Cells[r, 1], ws.Cells[r, 11]].Font.Bold = true;
                ws.Range[ws.Cells[r, 1], ws.Cells[r, 11]].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(220, 235, 252));
                r++;

                if (hm.DanhSachHangMucCon != null && hm.DanhSachHangMucCon.Count > 0)
                {
                    var hmcRows = new List<int>();
                    foreach (var hmc in hm.DanhSachHangMucCon)
                    {
                        // Dòng Tiêu đề Hạng mục con (Level 2)
                        int hmcRow = r;
                        hmc.RowIndex = hmcRow;
                        hmcRows.Add(hmcRow);

                        ws.Cells[r, 1] = hmc.STT;
                        ws.Cells[r, 3] = hmc.TenHangMucCon;
                        ws.Range[ws.Cells[r, 1], ws.Cells[r, 11]].Font.Bold = true;
                        ws.Range[ws.Cells[r, 1], ws.Cells[r, 11]].Font.Italic = true;
                        ws.Range[ws.Cells[r, 1], ws.Cells[r, 11]].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(242, 245, 249));
                        r++;

                        int hmcStart = r;
                        foreach (var dong in hmc.DanhSachCongTac)
                        {
                            dong.STT = r;
                            ws.Cells[r, 1] = sttCongTac++;
                            ws.Cells[r, 2] = dong.MaHieu;
                            ws.Cells[r, 3] = dong.TenCongTac;
                            ws.Cells[r, 4] = dong.DonVi;
                            ws.Cells[r, 5] = (double)dong.KhoiLuong;

                            if (dong.DonGiaVL > 0 || dong.DonGiaNC > 0 || dong.DonGiaMay > 0)
                            {
                                ws.Cells[r, 6] = (double)dong.DonGiaVL;
                                ws.Cells[r, 7] = (double)dong.DonGiaNC;
                                ws.Cells[r, 8] = (double)dong.DonGiaMay;
                            }

                            ws.Cells[r, 9].Formula = $"=ROUND(E{r}*F{r}, 0)";
                            ws.Cells[r, 10].Formula = $"=ROUND(E{r}*G{r}, 0)";
                            ws.Cells[r, 11].Formula = $"=ROUND(E{r}*H{r}, 0)";
                            r++;
                        }
                        int hmcEnd = r - 1;

                        if (hmcEnd >= hmcStart)
                        {
                            ws.Cells[hmcRow, 9].Formula = $"=SUM(I{hmcStart}:I{hmcEnd})";
                            ws.Cells[hmcRow, 10].Formula = $"=SUM(J{hmcStart}:J{hmcEnd})";
                            ws.Cells[hmcRow, 11].Formula = $"=SUM(K{hmcStart}:K{hmcEnd})";
                        }
                    }

                    if (hmcRows.Count > 0)
                    {
                        ws.Cells[hmRow, 9].Formula = $"={string.Join("+", hmcRows.Select(x => $"I{x}"))}";
                        ws.Cells[hmRow, 10].Formula = $"={string.Join("+", hmcRows.Select(x => $"J{x}"))}";
                        ws.Cells[hmRow, 11].Formula = $"={string.Join("+", hmcRows.Select(x => $"K{x}"))}";
                    }
                }
                else
                {
                    int hmStart = r;
                    foreach (var dong in hm.DanhSachCongTac)
                    {
                        dong.STT = r;
                        ws.Cells[r, 1] = sttCongTac++;
                        ws.Cells[r, 2] = dong.MaHieu;
                        ws.Cells[r, 3] = dong.TenCongTac;
                        ws.Cells[r, 4] = dong.DonVi;
                        ws.Cells[r, 5] = (double)dong.KhoiLuong;

                        if (dong.DonGiaVL > 0 || dong.DonGiaNC > 0 || dong.DonGiaMay > 0)
                        {
                            ws.Cells[r, 6] = (double)dong.DonGiaVL;
                            ws.Cells[r, 7] = (double)dong.DonGiaNC;
                            ws.Cells[r, 8] = (double)dong.DonGiaMay;
                        }

                        ws.Cells[r, 9].Formula = $"=ROUND(E{r}*F{r}, 0)";
                        ws.Cells[r, 10].Formula = $"=ROUND(E{r}*G{r}, 0)";
                        ws.Cells[r, 11].Formula = $"=ROUND(E{r}*H{r}, 0)";
                        r++;
                    }
                    int hmEnd = r - 1;

                    if (hmEnd >= hmStart)
                    {
                        ws.Cells[hmRow, 9].Formula = $"=SUM(I{hmStart}:I{hmEnd})";
                        ws.Cells[hmRow, 10].Formula = $"=SUM(J{hmStart}:J{hmEnd})";
                        ws.Cells[hmRow, 11].Formula = $"=SUM(K{hmStart}:K{hmEnd})";
                    }
                }
            }

            // Dòng TỔNG CỘNG ở cuối bảng DuToan
            if (hmRows.Count > 0)
            {
                ws.Cells[r, 3] = "TỔNG CỘNG";
                ws.Cells[r, 9].Formula = $"={string.Join("+", hmRows.Select(x => $"I{x}"))}";
                ws.Cells[r, 10].Formula = $"={string.Join("+", hmRows.Select(x => $"J{x}"))}";
                ws.Cells[r, 11].Formula = $"={string.Join("+", hmRows.Select(x => $"K{x}"))}";
                ws.Range[ws.Cells[r, 1], ws.Cells[r, 11]].Font.Bold = true;
                ws.Range[ws.Cells[r, 1], ws.Cells[r, 11]].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(200, 225, 250));
                r++;
            }

            // Định dạng số cho cột Khối lượng, đơn giá và Thành tiền
            if (r > 6)
            {
                Range dataRange = ws.Range[$"A4:K{r - 1}"];
                dataRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                dataRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                ws.Range[$"A6:A{r - 1}"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                ws.Range[$"C6:C{r - 1}"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignJustify;
                ws.Range[$"C6:C{r - 1}"].WrapText = true;
                ws.Range[$"D6:D{r - 1}"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                
                Range numberCols = ws.Range[$"E6:K{r - 1}"];
                numberCols.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

                ExcelFormatHelper.ApplyQuantityFormat(ws.Range[$"E6:E{r - 1}"], 2);
                ExcelFormatHelper.ApplyIntegerFormat(ws.Range[$"F6:K{r - 1}"]);
            }
            
            // Freeze panes at row 5 (headers are rows 4 & 5)
            ws.Activate();
            app.ActiveWindow.FreezePanes = false;
            app.ActiveWindow.SplitRow = 5;
            app.ActiveWindow.SplitColumn = 0;
            app.ActiveWindow.FreezePanes = true;
        }
        finally
        {
            app.ScreenUpdating = true;
        }
    }

    private string GetCellValue(Worksheet ws, int row, int col)
    {
        try
        {
            Range range = ws.Cells[row, col];
            if (range.Value2 == null) return string.Empty;
            return range.Value2.ToString().Trim();
        }
        catch
        {
            return string.Empty;
        }
    }

    private decimal GetCellDecimal(Worksheet ws, int row, int col)
    {
        try
        {
            Range range = ws.Cells[row, col];
            if (range == null || range.Value2 == null) return 0m;
            object val = range.Value2;
            if (val is double d) return Convert.ToDecimal(d);
            if (val is int i) return (decimal)i;
            if (val is decimal m) return m;
            if (val is float f) return Convert.ToDecimal(f);
            string str = val.ToString().Trim();
            return UIHelper.ParseFlexibleDecimal(str, isPercentage: false);
        }
        catch
        {
            return 0m;
        }
    }

    public static bool IsRomanNumeral(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return false;
        string clean = s.Trim().TrimEnd('.', ')', ':', '-').ToUpper();
        return clean switch
        {
            "I" or "II" or "III" or "IV" or "V" or "VI" or "VII" or "VIII" or "IX" or "X" or "XI" or "XII" => true,
            _ => false
        };
    }

    public static string ToRomanNumeral(int number)
    {
        return number switch
        {
            1 => "I",
            2 => "II",
            3 => "III",
            4 => "IV",
            5 => "V",
            6 => "VI",
            7 => "VII",
            8 => "VIII",
            9 => "IX",
            10 => "X",
            11 => "XI",
            12 => "XII",
            _ => number.ToString()
        };
    }

    public static bool IsLevel1Header(string stt, string ten, HangMuc currentHM)
    {
        if (currentHM == null) return true; // Dòng tiêu đề đầu tiên luôn là Level 1
        if (IsRomanNumeral(stt)) return true;
        if (ten.StartsWith("HẠNG MỤC", StringComparison.OrdinalIgnoreCase) ||
            ten.StartsWith("PHẦN", StringComparison.OrdinalIgnoreCase) ||
            ten.StartsWith("HM ", StringComparison.OrdinalIgnoreCase) ||
            ten.StartsWith("HM.", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
        return false;
    }

    public static bool IsDienGiaiKhoiLuong(string stt, string ten)
    {
        if (string.IsNullOrWhiteSpace(ten)) return false;
        string t = ten.Trim();
        // Bắt đầu bằng gạch đầu dòng, cộng, nhân, ngoặc
        if (t.StartsWith("-") || t.StartsWith("+") || t.StartsWith("*") || t.StartsWith("/") || t.StartsWith("(")) return true;
        // Chứa phép nhân kích thước đo bóc: số * số hoặc số x số (VD: 2*1.5*0.8, 1.2 x 2.4)
        if (System.Text.RegularExpressions.Regex.IsMatch(t, @"\d+[\.,]?\d*\s*[\*xX]\s*\d+[\.,]?\d*")) return true;
        // Bắt đầu bằng chữ thường khi STT rỗng
        if (string.IsNullOrWhiteSpace(stt) && char.IsLower(t[0])) return true;
        return false;
    }

    public static bool IsLevel2Header(string stt, string ten)
    {
        if (IsDienGiaiKhoiLuong(stt, ten)) return false;

        string s = (stt ?? "").Trim().ToLower();
        string t = (ten ?? "").Trim().ToLower();

        // Có STT dạng 1, 2, A, B, a)...
        if (!string.IsNullOrEmpty(s) && (int.TryParse(s.TrimEnd('.', ')'), out _) || s.Length <= 4))
        {
            return true;
        }

        // Hoặc tên bắt đầu bằng các từ khóa phân đoạn
        if (t.StartsWith("phần") || t.StartsWith("gói") || t.StartsWith("khu") || 
            t.StartsWith("tầng") || t.StartsWith("tuyến") || t.StartsWith("đoạn") ||
            t.StartsWith("hạng mục"))
        {
            return true;
        }

        return false;
    }

    public static string NhanDienLoaiCongTrinh(string ten)
    {
        if (string.IsNullOrWhiteSpace(ten)) return "Dân dụng";
        string lower = ten.ToLower();

        if (lower.Contains("giao thông") || lower.Contains("đường") || lower.Contains("cầu") || (lower.Contains("vỉa hè") && !lower.Contains("thoát")))
            return "Giao thông";
        if (lower.Contains("thoát nước") || lower.Contains("cấp nước") || lower.Contains("chiếu sáng") || lower.Contains("hạ tầng") || lower.Contains("cây xanh"))
            return "Hạ tầng kỹ thuật";
        if (lower.Contains("thủy lợi") || lower.Contains("kênh") || lower.Contains("mương") || lower.Contains("đê") || lower.Contains("đập") || lower.Contains("hồ chứa") || lower.Contains("nông nghiệp"))
            return DinhMucTT38Database.LoaiNongNghiepMoiTruong;
        if (lower.Contains("nhà xưởng") || lower.Contains("trạm biến áp") || lower.Contains("đường dây") || lower.Contains("công nghiệp") || lower.Contains("kho bãi"))
            return "Công nghiệp";
        if (lower.Contains("nhà") || lower.Contains("trường") || lower.Contains("trạm y tế") || lower.Contains("văn phòng") || lower.Contains("dân dụng") || lower.Contains("hội trường"))
            return "Dân dụng";

        return "Dân dụng";
    }

    /// <summary>
    /// Đọc lại hệ số điều chỉnh riêng VL/NC/M theo từng hạng mục từ sheet HeSo_DieuChinh (nếu có).
    /// Dòng 10: Hệ số VL, Dòng 11: Hệ số NC, Dòng 12: Hệ số M. Cột bắt đầu từ cột 3 (C).
    /// </summary>
    public void DocHeSoDieuChinhTuSheet(DuToan duToan)
    {
        if (duToan?.DanhSachHangMuc == null || duToan.DanhSachHangMuc.Count == 0) return;
        try
        {
            var app = (Application)ExcelDnaUtil.Application;
            var wb = app.ActiveWorkbook;
            if (wb == null) return;

            Worksheet ws = null;
            foreach (Worksheet sheet in wb.Sheets)
            {
                if (string.Equals(sheet.Name, "HeSo_DieuChinh", StringComparison.OrdinalIgnoreCase))
                {
                    ws = sheet;
                    break;
                }
            }
            if (ws == null) return;

            int nHm = duToan.DanhSachHangMuc.Count;
            for (int i = 0; i < nHm; i++)
            {
                int col = 3 + i;
                var hm = duToan.DanhSachHangMuc[i];

                if (hm.ChiPhiXD == null) hm.ChiPhiXD = new ChiPhiXayDung();
                decimal cpc = 0m, tt = 0m, tl = 0m, gtgt = 0m, lt = 0m;
                var valCPC = ws.Cells[5, col]?.Value2;
                if (valCPC != null && decimal.TryParse(valCPC.ToString(), out cpc) && cpc > 0)
                {
                    hm.ChiPhiXD.TiLeCPC = cpc;
                }
                var valTT = ws.Cells[6, col]?.Value2;
                if (valTT != null && decimal.TryParse(valTT.ToString(), out tt) && tt >= 0)
                {
                    hm.ChiPhiXD.TiLeTT = tt;
                }
                var valTNCTTT = ws.Cells[7, col]?.Value2;
                if (valTNCTTT != null && decimal.TryParse(valTNCTTT.ToString(), out tl) && tl > 0)
                {
                    hm.ChiPhiXD.TiLeTNCTTT = tl;
                }
                var valGTGT = ws.Cells[8, col]?.Value2;
                if (valGTGT != null && decimal.TryParse(valGTGT.ToString(), out gtgt) && gtgt >= 0)
                {
                    hm.ChiPhiXD.TiLeGTGT = gtgt;
                }
                var valLT = ws.Cells[9, col]?.Value2;
                if (valLT != null && decimal.TryParse(valLT.ToString(), out lt) && lt >= 0)
                {
                    hm.ChiPhiXD.TiLeNhaTam = lt;
                }

                var valVL = ws.Cells[10, col]?.Value2;
                if (valVL != null)
                {
                    if (decimal.TryParse(valVL.ToString(), out decimal hsVL) && hsVL > 0)
                    {
                        hm.HeSoVL = hsVL;
                    }
                }

                var valNC = ws.Cells[11, col]?.Value2;
                if (valNC != null)
                {
                    if (decimal.TryParse(valNC.ToString(), out decimal hsNC) && hsNC > 0)
                    {
                        hm.HeSoNC = hsNC;
                    }
                }

                var valM = ws.Cells[12, col]?.Value2;
                if (valM != null)
                {
                    if (decimal.TryParse(valM.ToString(), out decimal hsM) && hsM > 0)
                    {
                        hm.HeSoM = hsM;
                    }
                }
            }
        }
        catch { }
    }
}
