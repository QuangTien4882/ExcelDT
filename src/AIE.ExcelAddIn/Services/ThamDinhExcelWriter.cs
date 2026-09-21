using AIE.Core.Enums;
using AIE.Core.Models;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace AIE.ExcelAddIn.Services;

public class ThamDinhExcelWriter
{
    public void ExportResult(ThamDinhConfig config, List<KetQuaCongTacThamDinh> results, int? boDonGiaId = null)
    {
        var app = (Application)ExcelDnaUtil.Application;
        var wb = app.ActiveWorkbook;
        
        // Find source sheet
        Worksheet sourceSheet = null;
        foreach (Worksheet sheet in wb.Worksheets)
        {
            if (sheet.Name == config.SheetName)
            {
                sourceSheet = sheet;
                break;
            }
        }
        
        if (sourceSheet == null) throw new Exception("Không tìm thấy sheet nguồn.");

        // Delete old result sheets if exist
        string[] sheetsToDelete = { "KQ_ThamDinh", "KQ_GiaVL", "KQ_GiaNC", "KQ_GiaMay" };
        foreach (string sheetName in sheetsToDelete)
        {
            foreach (Worksheet sheet in wb.Worksheets)
            {
                if (sheet.Name == sheetName)
                {
                    app.DisplayAlerts = false;
                    sheet.Delete();
                    app.DisplayAlerts = true;
                    break;
                }
            }
        }

        // Copy sheet xuống cuối cùng
        sourceSheet.Copy(After: wb.Worksheets[wb.Worksheets.Count]);
        var resultSheet = (Worksheet)wb.ActiveSheet;
        resultSheet.Name = "KQ_ThamDinh";

        int dmIndex = ColLetterToNumber(config.ColDinhMuc);
        int dgIndex = ColLetterToNumber(config.ColDonGia);
        int ttIndex = ColLetterToNumber(config.ColThanhTien);

        int headerRow = config.DongBatDau - 1;
        if (headerRow < 1) headerRow = 1;

        // Chèn 1 dòng trống ngay dưới tiêu đề gốc để làm dòng phụ
        Range rowToInsert = (Range)resultSheet.Rows[headerRow + 1];
        rowToInsert.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

        // Cập nhật lại số dòng trong results do toàn bộ dữ liệu bị đẩy xuống 1 dòng
        foreach (var kq in results)
        {
            kq.DuToan.SoDongExcel++;
            foreach (var hp in kq.DuToan.DanhSachHaoPhi)
            {
                hp.SoDongExcel++;
            }
        }

        // ============================================================
        // CHÈN CÁC CỘT MỚI (KHÔNG xóa cột Đơn giá và Thành tiền)
        // ============================================================

        // --- Bước 1: Chèn 4 cột TT38 sau cột Định mức ---
        int colTenTT38 = 0, colDviTT38 = 0, colDmTT38 = 0, colDmDiff = 0;
        if (dmIndex > 0)
        {
            for (int i = 0; i < 4; i++)
            {
                Range c = (Range)resultSheet.Columns[dmIndex + 1];
                c.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            }
            colTenTT38 = dmIndex + 1;
            colDviTT38 = dmIndex + 2;
            colDmTT38 = dmIndex + 3;
            colDmDiff = dmIndex + 4;
        }

        // Cập nhật vị trí Đơn giá và Thành tiền (dịch phải 4 cột do chèn TT38)
        int colDonGiaDuToan = dgIndex > 0 ? dgIndex + 4 : 0;
        int colThanhTienDuToan = ttIndex > 0 ? ttIndex + 4 : 0;

        // --- Bước 2: Chèn 2 cột sau Đơn giá dự toán (ĐG chuẩn, Chênh lệch ĐG) ---
        int colDonGiaChuan = 0, colChenhLechDG = 0;
        if (colDonGiaDuToan > 0)
        {
            for (int i = 0; i < 2; i++)
            {
                Range c = (Range)resultSheet.Columns[colDonGiaDuToan + 1];
                c.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            }
            colDonGiaChuan = colDonGiaDuToan + 1;
            colChenhLechDG = colDonGiaDuToan + 2;
            // Thành tiền bị dịch thêm 2
            if (colThanhTienDuToan > 0) colThanhTienDuToan += 2;
        }

        // --- Bước 3: Chèn 2 cột sau Thành tiền dự toán (TT thẩm định, Chênh lệch TT) ---
        int colThanhTienTD = 0, colChenhLechTT = 0;
        if (colThanhTienDuToan > 0)
        {
            for (int i = 0; i < 2; i++)
            {
                Range c = (Range)resultSheet.Columns[colThanhTienDuToan + 1];
                c.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            }
            colThanhTienTD = colThanhTienDuToan + 1;
            colChenhLechTT = colThanhTienDuToan + 2;
        }

        // --- Bước 4: Chèn 2 cột Ghi chú lỗi ở cuối ---
        int rightmost = new[] { colChenhLechTT, colChenhLechDG, colDmDiff }.Where(x => x > 0).DefaultIfEmpty(0).Max();
        int colGhiChuDM, colGhiChuDG;
        if (rightmost > 0)
        {
            for (int i = 0; i < 2; i++)
            {
                Range c = (Range)resultSheet.Columns[rightmost + 1];
                c.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            }
            colGhiChuDM = rightmost + 1;
            colGhiChuDG = rightmost + 2;
        }
        else
        {
            Range usedRange = resultSheet.UsedRange;
            colGhiChuDM = usedRange.Columns.Count + usedRange.Column;
            colGhiChuDG = colGhiChuDM + 1;
        }

        // ============================================================
        // TÌM VÙNG DỮ LIỆU
        // ============================================================
        Range lastColRange = resultSheet.UsedRange;
        int lastCol = lastColRange.Columns.Count + lastColRange.Column - 1;

        int lastRow = headerRow;
        if (results.Count > 0)
        {
            lastRow = results.Max(x => 
                Math.Max(x.DuToan.SoDongExcel, 
                x.DuToan.DanhSachHaoPhi.Count > 0 ? x.DuToan.DanhSachHaoPhi.Max(h => h.SoDongExcel) : 0)
            );
        }

        // ============================================================
        // THIẾT LẬP FONT SIZE 12 CHO TOÀN BỘ SHEET
        // ============================================================
        resultSheet.Cells.Font.Size = 12;

        // ============================================================
        // ĐỊNH DẠNG TIÊU ĐỀ
        // ============================================================

        // Danh sách các cột mới (chèn thêm) - dùng để phân biệt với cột gốc
        var newCols = new List<int> { colTenTT38, colDviTT38, colDmTT38, colDmDiff,
            colDonGiaChuan, colChenhLechDG, colThanhTienTD, colChenhLechTT, colGhiChuDM, colGhiChuDG }
            .Where(x => x > 0).ToHashSet();

        // Gộp các ô tiêu đề GỐC theo chiều dọc (dòng headerRow và headerRow + 1)
        for (int c = 1; c <= lastCol; c++)
        {
            if (newCols.Contains(c)) continue; // Bỏ qua cột mới

            Range cellTop = resultSheet.Cells[headerRow, c];
            Range cellBottom = resultSheet.Cells[headerRow + 1, c];
            if (cellTop.Value2 != null || cellBottom.Value2 != null)
            {
                Range mergeRange = resultSheet.Range[cellTop, cellBottom];
                mergeRange.Merge();
                mergeRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            }
        }

        // Lấy font name chuẩn từ ô tiêu đề gốc
        string baseFontName = resultSheet.Cells[headerRow, 1].Font.Name?.ToString() ?? "Times New Roman";

        // Định dạng các cột mới (sub-header ở dòng headerRow + 1)
        foreach (int c in newCols)
        {
            Range subHeader = resultSheet.Cells[headerRow + 1, c];
            subHeader.Font.Bold = true;
            subHeader.Font.Size = 12;
            subHeader.Font.Name = baseFontName;
            subHeader.Font.Color = ColorTranslator.ToOle(Color.Black);
            subHeader.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            subHeader.VerticalAlignment = XlVAlign.xlVAlignCenter;
            subHeader.WrapText = true;
            
            if (c == colGhiChuDM)
            {
                subHeader.Interior.Color = ColorTranslator.ToOle(Color.Orange);
                subHeader.Value2 = "Lỗi Định mức";
                // Ghi chú gộp cả 2 dòng header
                Range mergeGhiChu = resultSheet.Range[resultSheet.Cells[headerRow, c], resultSheet.Cells[headerRow + 1, c]];
                mergeGhiChu.Merge();
                mergeGhiChu.VerticalAlignment = XlVAlign.xlVAlignCenter;
            }
            else if (c == colGhiChuDG)
            {
                subHeader.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 192, 0));
                subHeader.Value2 = "Lỗi Đơn giá";
                // Ghi chú gộp cả 2 dòng header
                Range mergeGhiChu = resultSheet.Range[resultSheet.Cells[headerRow, c], resultSheet.Cells[headerRow + 1, c]];
                mergeGhiChu.Merge();
                mergeGhiChu.VerticalAlignment = XlVAlign.xlVAlignCenter;
            }
            else
            {
                subHeader.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                if (c == colTenTT38) subHeader.Value2 = "Tên VT/NC/MTC";
                if (c == colDviTT38) subHeader.Value2 = "Đơn vị";
                if (c == colDmTT38) subHeader.Value2 = "Định mức";
                if (c == colDmDiff) subHeader.Value2 = "Chênh lệch ĐM";
                if (c == colDonGiaChuan) subHeader.Value2 = "Đơn giá chuẩn";
                if (c == colChenhLechDG) subHeader.Value2 = "Chênh lệch ĐG";
                if (c == colThanhTienTD) subHeader.Value2 = "Thành tiền (TĐ)";
                if (c == colChenhLechTT) subHeader.Value2 = "Chênh lệch TT";
            }

            // Độ rộng cột
            if (c == colTenTT38) resultSheet.Columns[c].ColumnWidth = 35;
            else if (c == colGhiChuDM || c == colGhiChuDG) resultSheet.Columns[c].ColumnWidth = 25;
            else resultSheet.Columns[c].ColumnWidth = 15;

            // Kẻ khung cho cột mới
            if (lastRow >= headerRow + 1)
            {
                Range colRange = resultSheet.Range[resultSheet.Cells[headerRow, c], resultSheet.Cells[lastRow, c]];
                colRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            }
        }

        // Tiêu đề gộp: "Định mức theo Thông tư 38"
        if (colTenTT38 > 0 && colDmDiff > 0)
        {
            Range groupHeader = resultSheet.Range[resultSheet.Cells[headerRow, colTenTT38], resultSheet.Cells[headerRow, colDmDiff]];
            groupHeader.Merge();
            groupHeader.Value2 = "Định mức theo Thông tư 38";
            groupHeader.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            groupHeader.VerticalAlignment = XlVAlign.xlVAlignCenter;
            groupHeader.Font.Bold = true;
            groupHeader.Font.Size = 12;
            groupHeader.Font.Name = baseFontName;
            groupHeader.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            groupHeader.Borders.LineStyle = XlLineStyle.xlContinuous;
        }

        // Tiêu đề gộp: "Đơn giá thẩm định"
        if (colDonGiaChuan > 0 && colChenhLechDG > 0)
        {
            Range groupHeader = resultSheet.Range[resultSheet.Cells[headerRow, colDonGiaChuan], resultSheet.Cells[headerRow, colChenhLechDG]];
            groupHeader.Merge();
            groupHeader.Value2 = "Đơn giá thẩm định";
            groupHeader.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            groupHeader.VerticalAlignment = XlVAlign.xlVAlignCenter;
            groupHeader.Font.Bold = true;
            groupHeader.Font.Size = 12;
            groupHeader.Font.Name = baseFontName;
            groupHeader.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            groupHeader.Borders.LineStyle = XlLineStyle.xlContinuous;
        }

        // Tiêu đề gộp: "Thành tiền thẩm định"
        if (colThanhTienTD > 0 && colChenhLechTT > 0)
        {
            Range groupHeader = resultSheet.Range[resultSheet.Cells[headerRow, colThanhTienTD], resultSheet.Cells[headerRow, colChenhLechTT]];
            groupHeader.Merge();
            groupHeader.Value2 = "Thành tiền thẩm định";
            groupHeader.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            groupHeader.VerticalAlignment = XlVAlign.xlVAlignCenter;
            groupHeader.Font.Bold = true;
            groupHeader.Font.Size = 12;
            groupHeader.Font.Name = baseFontName;
            groupHeader.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            groupHeader.Borders.LineStyle = XlLineStyle.xlContinuous;
        }

        // ============================================================
        // GHI DỮ LIỆU (Xử lý từ dưới lên trên để chèn dòng không sai lệch)
        // ============================================================
        foreach (var kq in results.OrderByDescending(x => x.DuToan.SoDongExcel))
        {
            var dt = kq.DuToan;
            
            // 1. Tô màu dòng công tác
            Range ctRange = resultSheet.Range[resultSheet.Cells[dt.SoDongExcel, 1], resultSheet.Cells[dt.SoDongExcel, lastCol]];
            
            if (kq.DinhMucChuan == null)
            {
                ctRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 230, 153)); // Cam nhạt
                resultSheet.Cells[dt.SoDongExcel, colGhiChuDM].Value2 = "Mã hiệu không tồn tại trong TT38";
                continue;
            }

            if (kq.DaDat)
            {
                ctRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(226, 239, 218)); // Xanh nhạt
            }
            else
            {
                ctRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(252, 228, 214)); // Đỏ nhạt
            }
            
            // In Tên công tác chuẩn TT38
            if (colTenTT38 > 0)
            {
                resultSheet.Cells[dt.SoDongExcel, colTenTT38].Value2 = kq.DinhMucChuan.TenCongTac;
                resultSheet.Cells[dt.SoDongExcel, colTenTT38].Font.Bold = true;
                resultSheet.Cells[dt.SoDongExcel, colTenTT38].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 242, 204));
            }

            // 2. Chèn dòng cho các "Hao phí thiếu" (Có trong TT38 nhưng không có trong Dự toán)
            var thieuItems = kq.DanhSachSaiLech.Where(x => x.HaoPhiDuToan == null && x.HaoPhiChuan != null).ToList();
            var loaiHps = new[] { AIE.Core.Enums.LoaiHaoPhi.MAY, AIE.Core.Enums.LoaiHaoPhi.NC, AIE.Core.Enums.LoaiHaoPhi.VL };
            
            foreach (var loai in loaiHps)
            {
                var thieuLoai = thieuItems.Where(x => x.HaoPhiChuan.LoaiHaoPhi == loai).ToList();
                if (thieuLoai.Any())
                {
                    int insertRow = dt.SoDongExcel;
                    var sameLoai = dt.DanhSachHaoPhi.Where(x => x.Loai == loai).ToList();
                    if (sameLoai.Any()) insertRow = sameLoai.Max(x => x.SoDongExcel);
                    else if (loai == AIE.Core.Enums.LoaiHaoPhi.MAY)
                    {
                        var nc = dt.DanhSachHaoPhi.Where(x => x.Loai == AIE.Core.Enums.LoaiHaoPhi.NC).ToList();
                        if (nc.Any()) insertRow = nc.Max(x => x.SoDongExcel);
                        else
                        {
                            var vl = dt.DanhSachHaoPhi.Where(x => x.Loai == AIE.Core.Enums.LoaiHaoPhi.VL).ToList();
                            if (vl.Any()) insertRow = vl.Max(x => x.SoDongExcel);
                        }
                    }
                    else if (loai == AIE.Core.Enums.LoaiHaoPhi.NC)
                    {
                        var vl = dt.DanhSachHaoPhi.Where(x => x.Loai == AIE.Core.Enums.LoaiHaoPhi.VL).ToList();
                        if (vl.Any()) insertRow = vl.Max(x => x.SoDongExcel);
                    }

                    for (int i = 0; i < thieuLoai.Count; i++)
                    {
                        int newRow = insertRow + 1 + i;
                        Range insertedRow = (Range)resultSheet.Rows[newRow];
                        insertedRow.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                        
                        Range newRowRange = resultSheet.Range[resultSheet.Cells[newRow, 1], resultSheet.Cells[newRow, lastCol]];
                        newRowRange.Interior.ColorIndex = 0;
                        newRowRange.Font.Bold = false;
                        
                        var sl = thieuLoai[i];
                        if (colTenTT38 > 0) resultSheet.Cells[newRow, colTenTT38].Value2 = sl.HaoPhiChuan.TenHaoPhi;
                        if (colDviTT38 > 0) resultSheet.Cells[newRow, colDviTT38].Value2 = sl.HaoPhiChuan.DonVi;
                        if (colDmTT38 > 0) resultSheet.Cells[newRow, colDmTT38].Value2 = sl.HaoPhiChuan.DinhMuc;
                        
                        // Đơn giá chuẩn cho hao phí thiếu
                        if (colDonGiaChuan > 0 && sl.DonGiaChuan.HasValue)
                            resultSheet.Cells[newRow, colDonGiaChuan].Value2 = sl.DonGiaChuan.Value;
                        
                        // Thành tiền thẩm định cho hao phí thiếu
                        if (colThanhTienTD > 0 && sl.DonGiaChuan.HasValue)
                            resultSheet.Cells[newRow, colThanhTienTD].Value2 = sl.HaoPhiChuan.DinhMuc * sl.DonGiaChuan.Value;
                        
                        resultSheet.Cells[newRow, colGhiChuDM].Value2 = $"[Thiếu hao phí] TT38 có '{sl.HaoPhiChuan.TenHaoPhi}'";
                        resultSheet.Cells[newRow, colGhiChuDM].Font.Color = ColorTranslator.ToOle(Color.Red);

                        // Cập nhật lại SoDongExcel cho các hao phí dự toán bị đẩy xuống
                        foreach (var hp in dt.DanhSachHaoPhi)
                        {
                            if (hp.SoDongExcel >= newRow) hp.SoDongExcel++;
                        }
                    }
                }
            }

            // 3. Ghi dữ liệu cho các hao phí Đã Match hoặc Thừa
            foreach (var hpDuToan in dt.DanhSachHaoPhi)
            {
                var rowExcel = hpDuToan.SoDongExcel;
                var saiLech = kq.DanhSachSaiLech.FirstOrDefault(x => x.HaoPhiDuToan == hpDuToan);
                
                if (saiLech != null)
                {
                    if (saiLech.HaoPhiChuan == null)
                    {
                        // Hao phí thừa
                        resultSheet.Cells[rowExcel, colGhiChuDM].Value2 = "Hao phí thừa";
                        resultSheet.Cells[rowExcel, colGhiChuDM].Font.Color = ColorTranslator.ToOle(Color.Red);
                        resultSheet.Cells[rowExcel, colGhiChuDM].Font.Bold = true;
                    }
                    else
                    {
                        // --- Phần Định mức TT38 ---
                        if (colTenTT38 > 0) resultSheet.Cells[rowExcel, colTenTT38].Value2 = saiLech.HaoPhiChuan.TenHaoPhi;
                        if (colDviTT38 > 0) resultSheet.Cells[rowExcel, colDviTT38].Value2 = saiLech.HaoPhiChuan.DonVi;
                        if (colDmTT38 > 0) resultSheet.Cells[rowExcel, colDmTT38].Value2 = saiLech.HaoPhiChuan.DinhMuc;
                        if (colDmDiff > 0)
                        {
                            resultSheet.Cells[rowExcel, colDmDiff].Value2 = saiLech.ChenhLechDinhMuc;
                            if (Math.Abs(saiLech.ChenhLechDinhMuc) > 0.0001m)
                                resultSheet.Cells[rowExcel, colDmDiff].Font.Color = ColorTranslator.ToOle(Color.Red);
                        }

                        // --- Phần Đơn giá ---
                        if (colDonGiaChuan > 0 && saiLech.DonGiaChuan.HasValue)
                        {
                            resultSheet.Cells[rowExcel, colDonGiaChuan].Value2 = saiLech.DonGiaChuan.Value;
                        }
                        if (colChenhLechDG > 0 && saiLech.DonGiaChuan.HasValue)
                        {
                            decimal chenhLechDG = hpDuToan.DonGia - saiLech.DonGiaChuan.Value;
                            resultSheet.Cells[rowExcel, colChenhLechDG].Value2 = chenhLechDG;
                            if (Math.Abs(chenhLechDG) > 1)
                                resultSheet.Cells[rowExcel, colChenhLechDG].Font.Color = ColorTranslator.ToOle(Color.Red);
                        }

                        // --- Phần Thành tiền ---
                        if (colThanhTienTD > 0 && saiLech.DonGiaChuan.HasValue)
                        {
                            decimal ttThamDinh = saiLech.HaoPhiChuan.DinhMuc * saiLech.DonGiaChuan.Value;
                            resultSheet.Cells[rowExcel, colThanhTienTD].Value2 = ttThamDinh;
                        }
                        if (colChenhLechTT > 0 && saiLech.DonGiaChuan.HasValue)
                        {
                            decimal ttThamDinh = saiLech.HaoPhiChuan.DinhMuc * saiLech.DonGiaChuan.Value;
                            decimal ttDuToan = hpDuToan.ThanhTien > 0 ? hpDuToan.ThanhTien : hpDuToan.DinhMuc * hpDuToan.DonGia;
                            decimal chenhLechTT = ttDuToan - ttThamDinh;
                            resultSheet.Cells[rowExcel, colChenhLechTT].Value2 = chenhLechTT;
                            if (Math.Abs(chenhLechTT) > 1)
                                resultSheet.Cells[rowExcel, colChenhLechTT].Font.Color = ColorTranslator.ToOle(Color.Red);
                        }

                        // --- Ghi chú lỗi tổng hợp (định mức + đơn giá) ---
                        // Lỗi từ thẩm định định mức
                        if (!string.IsNullOrEmpty(saiLech.LoaiLoi))
                        {
                            resultSheet.Cells[rowExcel, colGhiChuDM].Value2 = saiLech.LoaiLoi;
                            bool hasSerious = saiLech.LoaiLoi.Contains("Sai") || saiLech.LoaiLoi.Contains("thừa");
                            resultSheet.Cells[rowExcel, colGhiChuDM].Font.Color = hasSerious
                                ? ColorTranslator.ToOle(Color.Red)
                                : ColorTranslator.ToOle(Color.DarkOrange);
                        }

                        // Lỗi đơn giá
                        string loiDonGia = "";
                        if (saiLech.DonGiaChuan.HasValue)
                        {
                            decimal chenhLechDG = hpDuToan.DonGia - saiLech.DonGiaChuan.Value;
                            if (Math.Abs(chenhLechDG) > 1)
                                loiDonGia = "Sai đơn giá";
                        }
                        else if (!string.IsNullOrEmpty(saiLech.HaoPhiChuan?.MaHieuHP))
                        {
                            loiDonGia = "Không có giá chuẩn";
                        }

                        if (!string.IsNullOrEmpty(saiLech.GhiChuDonGia))
                        {
                            loiDonGia = string.IsNullOrEmpty(loiDonGia) ? saiLech.GhiChuDonGia : $"{loiDonGia}. {saiLech.GhiChuDonGia}";
                        }

                        if (!string.IsNullOrEmpty(loiDonGia))
                        {
                            resultSheet.Cells[rowExcel, colGhiChuDG].Value2 = loiDonGia;
                            resultSheet.Cells[rowExcel, colGhiChuDG].Font.Color = ColorTranslator.ToOle(Color.Red);
                        }
                    }
                }
            }
        }

        // ============================================================
        // CỐ ĐỊNH DÒNG TIÊU ĐỀ VÀ ĐỊNH DẠNG CUỐI CÙNG
        // ============================================================
        try
        {
            resultSheet.Activate();
            
            // Cuộn màn hình lên góc trái trên cùng
            resultSheet.Application.ActiveWindow.ScrollRow = 1;
            resultSheet.Application.ActiveWindow.ScrollColumn = 1;
            
            // Cố định dưới dòng headerRow + 1 (tiêu đề có 2 dòng)
            resultSheet.Application.ActiveWindow.SplitRow = headerRow + 1;
            resultSheet.Application.ActiveWindow.SplitColumn = 0;
            resultSheet.Application.ActiveWindow.FreezePanes = true;

            // Wrap Text cho vùng dữ liệu
            if (lastRow > headerRow + 1)
            {
                Range dataRange = resultSheet.Range[resultSheet.Cells[headerRow + 2, 1], resultSheet.Cells[lastRow, colGhiChuDG]];
                dataRange.WrapText = true;
                dataRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            }
        }
        catch { }
    }

    private int ColLetterToNumber(string letter)
    {
        if (string.IsNullOrEmpty(letter)) return 0;
        int col = 0;
        foreach (char c in letter.ToUpper())
        {
            col = col * 26 + (c - 'A' + 1);
        }
        return col;
    }

    private string GetColumnLetter(int colIndex)
    {
        if (colIndex <= 0) return "";
        int div = colIndex;
        string colLetter = string.Empty;
        int mod = 0;

        while (div > 0)
        {
            mod = (div - 1) % 26;
            colLetter = (char)(65 + mod) + colLetter;
            div = (int)((div - mod) / 26);
        }
        return colLetter;
    }

    // ============================================================
    // BƯỚC 4: XUẤT TRỌN BỘ HỒ SƠ BÁO CÁO THẨM TRA CHUẨN HÓA (3 SHEETS)
    // ============================================================
    // ============================================================
    // BÁO CÁO THẨM ĐỊNH DỰ TOÁN THEO THÔNG TƯ 36/2026/TT-BXD & NĐ 10/2021/NĐ-CP
    // ============================================================

    public void ExportBaoCaoHoanChinh(ThamDinhConfig config, List<KetQuaCongTacThamDinh> results, int? boDonGiaId)
    {
        var defaultInfo = new ThongTinBaoCaoThamDinh();
        ExportBaoCaoTheoThongTu36(config, results, boDonGiaId, defaultInfo);
    }

    public void ExportBaoCaoTheoThongTu36(
        ThamDinhConfig config,
        List<KetQuaCongTacThamDinh> results,
        int? boDonGiaId,
        ThongTinBaoCaoThamDinh thongTin)
    {
        var app = (Application)ExcelDnaUtil.Application;
        var wb = app.ActiveWorkbook;
        if (wb == null) throw new Exception("Không có file Excel nào đang mở.");

        // Xóa các sheet báo cáo cũ nếu đã tồn tại
        string[] sheetsToDelete = { "BC_ThamDinh_TongHop", "BC_ThamTra_TongHop", "TH_ChiPhiXD_ThamDinh", "DoiChieu_ChiTiet", "TH_GiaVatTu_ThamDinh", "TH_GiaVatTu_ThamTra" };
        foreach (string sName in sheetsToDelete)
        {
            foreach (Worksheet ws in wb.Worksheets)
            {
                if (ws.Name == sName)
                {
                    app.DisplayAlerts = false;
                    ws.Delete();
                    app.DisplayAlerts = true;
                    break;
                }
            }
        }

        Worksheet wsFirst = null;

        // 1. Sheet TH_ChiPhiXD_ThamDinh (Bảng 3.8 Thông tư 36/2026/TT-BXD)
        Worksheet wsChiPhiXD = null;
        if (thongTin.XuatChiPhiXD)
        {
            wsChiPhiXD = (Worksheet)wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
            wsChiPhiXD.Name = "TH_ChiPhiXD_ThamDinh";
            TaoSheetChiPhiXD(wsChiPhiXD, results, thongTin);
            if (wsFirst == null) wsFirst = wsChiPhiXD;
        }

        // 2. Sheet BC_ThamDinh_TongHop (Văn bản Báo cáo thẩm định)
        Worksheet wsTongHop = null;
        if (thongTin.XuatTongHop)
        {
            wsTongHop = (Worksheet)wb.Worksheets.Add(Before: wsChiPhiXD ?? wb.Worksheets[wb.Worksheets.Count]);
            wsTongHop.Name = "BC_ThamDinh_TongHop";
            TaoSheetBaoCaoTongHop(wsTongHop, results, config, thongTin);
            wsFirst = wsTongHop;
        }

        // 3. Sheet DoiChieu_ChiTiet (Đối chiếu chi tiết từng công tác)
        if (thongTin.XuatDoiChieuChiTiet)
        {
            var wsChiTiet = (Worksheet)wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
            wsChiTiet.Name = "DoiChieu_ChiTiet";
            TaoSheetDoiChieuChiTiet(wsChiTiet, results);
            if (wsFirst == null) wsFirst = wsChiTiet;
        }

        // 4. Sheet TH_GiaVatTu_ThamDinh (Bảng giá vật tư, nhân công, ca máy)
        if (thongTin.XuatGiaVatTu)
        {
            var wsGiaVatTu = (Worksheet)wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
            wsGiaVatTu.Name = "TH_GiaVatTu_ThamDinh";
            TaoSheetGiaVatTu(wsGiaVatTu, boDonGiaId);
            if (wsFirst == null) wsFirst = wsGiaVatTu;
        }

        wsFirst?.Activate();
    }

    private void TinhTongHopChiPhiTrucTiep(
        List<KetQuaCongTacThamDinh> results,
        out decimal vlDT, out decimal ncDT, out decimal mayDT,
        out decimal vlTT, out decimal ncTT, out decimal mayTT)
    {
        vlDT = 0; ncDT = 0; mayDT = 0;
        vlTT = 0; ncTT = 0; mayTT = 0;

        foreach (var kq in results)
        {
            decimal kl = kq.DuToan.KhoiLuong > 0 ? kq.DuToan.KhoiLuong : 1.0m;

            // 1. Chi phí trực tiếp Dự toán đệ trình
            if (kq.DuToan.DanhSachHaoPhi != null && kq.DuToan.DanhSachHaoPhi.Count > 0)
            {
                foreach (var hp in kq.DuToan.DanhSachHaoPhi)
                {
                    decimal dgHp = hp.DonGia;
                    decimal ttHaoPhi = (hp.ThanhTien > 0 ? hp.ThanhTien : (hp.DinhMuc * dgHp)) * (kq.DuToan.KhoiLuong > 0 && kq.DuToan.KhoiLuong != 1.0m ? kl : 1.0m);
                    if (hp.Loai == LoaiHaoPhi.VL) vlDT += ttHaoPhi;
                    else if (hp.Loai == LoaiHaoPhi.NC) ncDT += ttHaoPhi;
                    else if (hp.Loai == LoaiHaoPhi.MAY) mayDT += ttHaoPhi;
                }
            }
            else
            {
                // Công tác không có chiết tính hao phí (tạm tính), phân bổ theo cơ cấu định mức trung bình (60% VL, 25% NC, 15% May)
                decimal ttItem = kq.DuToan.ThanhTien > 0 ? kq.DuToan.ThanhTien : (kl * kq.DuToan.DonGia);
                if (ttItem > 0)
                {
                    vlDT += Math.Round(ttItem * 0.60m, 0);
                    ncDT += Math.Round(ttItem * 0.25m, 0);
                    mayDT += (ttItem - Math.Round(ttItem * 0.60m, 0) - Math.Round(ttItem * 0.25m, 0));
                }
            }

            // 2. Chi phí trực tiếp Thẩm tra xác định
            if (kq.DanhSachSaiLech != null && kq.DanhSachSaiLech.Count > 0)
            {
                foreach (var s in kq.DanhSachSaiLech)
                {
                    decimal ttChuan = s.ThanhTienChuan * (kq.DuToan.KhoiLuong > 0 && kq.DuToan.KhoiLuong != 1.0m ? kl : 1.0m);
                    if (s.LoaiHP == LoaiHaoPhi.VL) vlTT += ttChuan;
                    else if (s.LoaiHP == LoaiHaoPhi.NC) ncTT += ttChuan;
                    else if (s.LoaiHP == LoaiHaoPhi.MAY) mayTT += ttChuan;
                }
            }
            else
            {
                // Nếu không có sai lệch thì lấy bằng giá trị dự toán
                decimal ttItem = kq.DuToan.ThanhTien > 0 ? kq.DuToan.ThanhTien : (kl * kq.DuToan.DonGia);
                if (ttItem > 0)
                {
                    vlTT += Math.Round(ttItem * 0.60m, 0);
                    ncTT += Math.Round(ttItem * 0.25m, 0);
                    mayTT += (ttItem - Math.Round(ttItem * 0.60m, 0) - Math.Round(ttItem * 0.25m, 0));
                }
            }
        }

        // Làm tròn số nguyên đồng
        vlDT = Math.Round(vlDT, 0);
        ncDT = Math.Round(ncDT, 0);
        mayDT = Math.Round(mayDT, 0);
        vlTT = Math.Round(vlTT, 0);
        ncTT = Math.Round(ncTT, 0);
        mayTT = Math.Round(mayTT, 0);

        // Đảm bảo không bị bằng 0 nếu tổng công tác có giá trị
        decimal tongCongTacDT = results.Sum(x => x.DuToan.ThanhTien > 0 ? x.DuToan.ThanhTien : ((x.DuToan.KhoiLuong > 0 ? x.DuToan.KhoiLuong : 1.0m) * x.DuToan.DonGia));
        decimal tongCongTacTT = results.Sum(x => x.DanhSachSaiLech.Count > 0 ? x.DanhSachSaiLech.Sum(s => s.ThanhTienChuan * (x.DuToan.KhoiLuong > 0 ? x.DuToan.KhoiLuong : 1.0m)) : (x.DuToan.ThanhTien > 0 ? x.DuToan.ThanhTien : 0));

        if ((vlDT + ncDT + mayDT) == 0 && tongCongTacDT > 0)
        {
            vlDT = Math.Round(tongCongTacDT * 0.65m, 0);
            ncDT = Math.Round(tongCongTacDT * 0.23m, 0);
            mayDT = tongCongTacDT - vlDT - ncDT;
        }

        if ((vlTT + ncTT + mayTT) == 0 && tongCongTacTT > 0)
        {
            vlTT = Math.Round(tongCongTacTT * 0.65m, 0);
            ncTT = Math.Round(tongCongTacTT * 0.23m, 0);
            mayTT = tongCongTacTT - vlTT - ncTT;
        }
        else if ((vlTT + ncTT + mayTT) == 0 && (vlDT + ncDT + mayDT) > 0)
        {
            vlTT = vlDT;
            ncTT = ncDT;
            mayTT = mayDT;
        }
    }

    private void TaoSheetChiPhiXD(Worksheet ws, List<KetQuaCongTacThamDinh> results, ThongTinBaoCaoThamDinh thongTin)
    {
        ws.Cells.Font.Name = "Times New Roman";
        ws.Cells.Font.Size = 11;

        // Title
        Range rTitle = ws.Range["A1", "I1"];
        rTitle.Merge();
        rTitle.Value2 = "BẢNG TỔNG HỢP CHI PHÍ XÂY DỰNG (BẢNG 3.8 THÔNG TƯ 36/2026/TT-BXD)";
        rTitle.Font.Bold = true;
        rTitle.Font.Size = 14;
        rTitle.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 51, 102));
        rTitle.HorizontalAlignment = XlHAlign.xlHAlignCenter;

        Range rSub = ws.Range["A2", "I2"];
        rSub.Merge();
        rSub.Value2 = $"Công trình: {thongTin.TenDuAn} - Hạng mục: {thongTin.TenHangMuc} | {thongTin.LoaiCongTrinh} ({thongTin.CapCongTrinh})";
        rSub.Font.Italic = true;
        rSub.Font.Size = 10.5f;
        rSub.HorizontalAlignment = XlHAlign.xlHAlignCenter;

        string[] headers = { "STT", "Khoản mục chi phí", "Ký hiệu", "Cách tính", "Tỷ lệ (%)", "Dự toán đệ trình (đ)", "Thẩm tra xác định (đ)", "Chênh lệch (+/- đ)", "Ghi chú" };
        int[] widths = { 6, 32, 10, 22, 11, 20, 20, 20, 25 };
        for (int i = 0; i < headers.Length; i++)
        {
            Range c = (Range)ws.Cells[4, i + 1];
            c.Value2 = headers[i];
            c.Font.Bold = true;
            c.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            c.VerticalAlignment = XlVAlign.xlVAlignCenter;
            c.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(230, 238, 248));
            ws.Columns[i + 1].ColumnWidth = widths[i];
        }

        TinhTongHopChiPhiTrucTiep(results, out decimal vlDT, out decimal ncDT, out decimal mayDT, out decimal vlTT, out decimal ncTT, out decimal mayTT);

        object[,] dataTable = new object[14, 9]
        {
            { "I", "CHI PHÍ TRỰC TIẾP", "T", "VL + NC + M", "", "", "", "", "" },
            { "1", "Chi phí vật liệu", "VL", "Theo tổng hợp vật liệu", "", vlDT, vlTT, "=G6-F6", "Áp theo công bố giá thị trường" },
            { "2", "Chi phí nhân công", "NC", "Theo đơn giá nhân công", "", ncDT, ncTT, "=G7-F7", "Theo công bố giá nhân công" },
            { "3", "Chi phí máy thi công", "M", "Theo bảng giá ca máy", "", mayDT, mayTT, "=G8-F8", "Theo bảng giá ca máy thẩm định" },
            { "", "Cộng Chi phí trực tiếp", "T", "VL + NC + M", "", "=SUM(F6:F8)", "=SUM(G6:G8)", "=G9-F9", "" },
            { "II", "CHI PHÍ GIÁN TIẾP", "GT", "CPC + Gnt + Gkxd", "", "", "", "", "" },
            { "1", "Chi phí chung", "CPC", "T x TiLe%", thongTin.TiLeCPC, "=ROUND(F9*E11/100,0)", "=ROUND(G9*E11/100,0)", "=G11-F11", "Bảng 3.3 Thông tư 36/2026/TT-BXD" },
            { "2", "Chi phí nhà tạm để ở và ĐHTC", "Gnt", "T x TiLe%", thongTin.TiLeNhaTam, "=ROUND(F9*E12/100,0)", "=ROUND(G9*E12/100,0)", "=G12-F12", "Bảng 3.7 Thông tư 36/2026/TT-BXD" },
            { "3", "CP một số CV không XĐ từ TK", "Gkxd", "T x TiLe%", thongTin.TiLeKhongXacDinh, "=ROUND(F9*E13/100,0)", "=ROUND(G9*E13/100,0)", "=G13-F13", "Bảng 3.5 Thông tư 36/2026/TT-BXD" },
            { "", "Cộng Chi phí gián tiếp", "GT", "CPC + Gnt + Gkxd", "", "=SUM(F11:F13)", "=SUM(G11:G13)", "=G14-F14", "" },
            { "III", "THU NHẬP CHỊU THUẾ TÍNH TRƯỚC", "TL", "(T + GT) x TiLe%", thongTin.TiLeTNCTTT, "=ROUND((F9+F14)*E15/100,0)", "=ROUND((G9+G14)*E15/100,0)", "=G15-F15", "Bảng 3.6 Thông tư 36/2026/TT-BXD" },
            { "IV", "CHI PHÍ XÂY DỰNG TRƯỚC THUẾ", "GXDTT", "T + GT + TL", "", "=F9+F14+F15", "=G9+G14+G15", "=G16-F16", "" },
            { "V", "THUẾ GIÁ TRỊ GIA TĂNG", "VAT", "GXDTT x TiLe%", thongTin.TiLeVAT, "=ROUND(F16*E17/100,0)", "=ROUND(G16*E17/100,0)", "=G17-F17", "" },
            { "VI", "CHI PHÍ XÂY DỰNG SAU THUẾ", "GXD", "GXDTT + VAT", "", "=F16+F17", "=G16+G17", "=G18-F18", "Bảng 3.8 Thông tư 36/2026/TT-BXD" }
        };

        for (int r = 0; r < 14; r++)
        {
            int rowIdx = 5 + r;
            for (int c = 0; c < 9; c++)
            {
                ws.Cells[rowIdx, c + 1].Value2 = dataTable[r, c];
            }

            // In đậm các dòng tổng
            if (r == 0 || r == 4 || r == 5 || r == 9 || r == 10 || r == 11 || r == 12 || r == 13)
            {
                Range rFormat = ws.Range[ws.Cells[rowIdx, 1], ws.Cells[rowIdx, 9]];
                rFormat.Font.Bold = true;
                if (r == 4 || r == 9 || r == 11)
                    rFormat.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(240, 245, 255));
                else if (r == 13)
                    rFormat.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(234, 250, 234));
            }
        }

        // Căn giữa STT và ký hiệu
        Range centerCol = ws.Range["A5", "C18"];
        centerCol.HorizontalAlignment = XlHAlign.xlHAlignCenter;

        Range pctCol = ws.Range["E5", "E18"];
        pctCol.NumberFormat = "0.00";
        pctCol.HorizontalAlignment = XlHAlign.xlHAlignRight;

        Range moneyCols = ws.Range["F5", "H18"];
        moneyCols.NumberFormat = "#,##0";

        DrawTableBorders(ws, 4, 1, 18, 9);
    }

    private void TaoSheetBaoCaoTongHop(
        Worksheet ws,
        List<KetQuaCongTacThamDinh> results,
        ThamDinhConfig config,
        ThongTinBaoCaoThamDinh thongTin)
    {
        ws.Cells.Font.Name = "Times New Roman";
        ws.Cells.Font.Size = 11;

        // Header cơ quan & Quốc hiệu
        ws.Cells[1, 1].Value2 = !string.IsNullOrWhiteSpace(thongTin.DonViThamDinh) ? thongTin.DonViThamDinh.ToUpper() : "CƠ QUAN THẨM ĐỊNH";
        ws.Cells[1, 1].Font.Bold = true;
        ws.Cells[1, 1].Font.Size = 10f;
        ws.Cells[2, 1].Value2 = $"Số: {thongTin.SoVanBan}";
        ws.Cells[2, 1].Font.Italic = true;
        ws.Cells[2, 1].Font.Size = 10f;

        Range rQuocHieu = ws.Range["E1", "H1"];
        rQuocHieu.Merge();
        rQuocHieu.Value2 = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
        rQuocHieu.Font.Bold = true;
        rQuocHieu.Font.Size = 11;
        rQuocHieu.HorizontalAlignment = XlHAlign.xlHAlignCenter;

        Range rTieuNgu = ws.Range["E2", "H2"];
        rTieuNgu.Merge();
        rTieuNgu.Value2 = "Độc lập - Tự do - Hạnh phúc";
        rTieuNgu.Font.Bold = true;
        rTieuNgu.Font.Size = 11;
        rTieuNgu.Font.Underline = XlUnderlineStyle.xlUnderlineStyleSingle;
        rTieuNgu.HorizontalAlignment = XlHAlign.xlHAlignCenter;

        Range rNgay = ws.Range["E3", "H3"];
        rNgay.Merge();
        string diaDiemText = !string.IsNullOrWhiteSpace(thongTin.DiaDiem) ? thongTin.DiaDiem : "Địa danh";
        rNgay.Value2 = $"{diaDiemText}, ngày {thongTin.NgayLap:dd} tháng {thongTin.NgayLap:MM} năm {thongTin.NgayLap:yyyy}";
        rNgay.Font.Italic = true;
        rNgay.Font.Size = 10f;
        rNgay.HorizontalAlignment = XlHAlign.xlHAlignCenter;

        // Title
        Range rTitle = ws.Range["A5", "H5"];
        rTitle.Merge();
        rTitle.Value2 = "BÁO CÁO KẾT QUẢ THẨM ĐỊNH DỰ TOÁN XÂY DỰNG CÔNG TRÌNH";
        rTitle.Font.Bold = true;
        rTitle.Font.Size = 15;
        rTitle.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 51, 102));
        rTitle.HorizontalAlignment = XlHAlign.xlHAlignCenter;

        Range rCanCu = ws.Range["A6", "H6"];
        rCanCu.Merge();
        rCanCu.Value2 = "(Căn cứ Nghị định số 10/2021/NĐ-CP và Thông tư số 36/2026/TT-BXD, TT 38/2026/TT-BXD của Bộ Xây dựng)";
        rCanCu.Font.Italic = true;
        rCanCu.Font.Size = 10f;
        rCanCu.HorizontalAlignment = XlHAlign.xlHAlignCenter;

        Range rKinhGui = ws.Range["A7", "H7"];
        rKinhGui.Merge();
        rKinhGui.Value2 = $"Kính gửi: {thongTin.ChuDauTu}";
        rKinhGui.Font.Bold = true;
        rKinhGui.HorizontalAlignment = XlHAlign.xlHAlignCenter;

        // I. THÔNG TIN CHUNG
        int rInfo = 9;
        Range rSec0 = ws.Range[$"A{rInfo}", $"H{rInfo}"];
        rSec0.Merge();
        rSec0.Value2 = "I. THÔNG TIN CHUNG VỀ DỰ ÁN, CÔNG TRÌNH";
        rSec0.Font.Bold = true;
        rSec0.Font.Size = 11.5f;
        rSec0.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 70, 140));

        ws.Cells[rInfo + 1, 1].Value2 = $"1. Tên công trình / dự án: {thongTin.TenDuAn}";
        ws.Cells[rInfo + 2, 1].Value2 = $"2. Hạng mục công trình: {thongTin.TenHangMuc}    |    Địa điểm: {thongTin.DiaDiem}";
        ws.Cells[rInfo + 3, 1].Value2 = $"3. Chủ đầu tư: {thongTin.ChuDauTu}    |    Đơn vị lập dự toán: {thongTin.DonViTuVan}";
        ws.Cells[rInfo + 4, 1].Value2 = $"4. Đơn vị thẩm định: {thongTin.DonViThamDinh}    |    Loại công trình: {thongTin.LoaiCongTrinh} ({thongTin.CapCongTrinh})";

        // II. BẢNG 1: TỔNG HỢP KẾT QUẢ ĐỐI CHIẾU KINH PHÍ (6 KHOẢN MỤC NĐ 10/2021)
        int rStart1 = rInfo + 6;
        Range rSec1 = ws.Range[$"A{rStart1}", $"H{rStart1}"];
        rSec1.Merge();
        rSec1.Value2 = "II. TỔNG HỢP KẾT QUẢ THẨM ĐỊNH DỰ TOÁN (THEO NGHỊ ĐỊNH 10/2021/NĐ-CP)";
        rSec1.Font.Bold = true;
        rSec1.Font.Size = 11.5f;
        rSec1.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 70, 140));

        string[] headers1 = { "STT", "Khoản mục chi phí", "Ký hiệu", "Dự toán đệ trình (đ)", "Thẩm tra xác định (đ)", "Chênh lệch (+/- đ)", "Tỷ lệ (%)", "Ghi chú" };
        int[] widths1 = { 6, 32, 10, 22, 22, 22, 14, 25 };
        for (int i = 0; i < headers1.Length; i++)
        {
            Range c = (Range)ws.Cells[rStart1 + 1, i + 1];
            c.Value2 = headers1[i];
            c.Font.Bold = true;
            c.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            c.VerticalAlignment = XlVAlign.xlVAlignCenter;
            c.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(230, 238, 248));
            ws.Columns[i + 1].ColumnWidth = widths1[i];
        }

        int rData1 = rStart1 + 2;
        string gxdTrinhFormula = thongTin.XuatChiPhiXD ? "=TH_ChiPhiXD_ThamDinh!F18" : $"=SUM(DoiChieu_ChiTiet!I4:I{results.Count + 3})";
        string gxdTTFormula = thongTin.XuatChiPhiXD ? "=TH_ChiPhiXD_ThamDinh!G18" : $"=SUM(DoiChieu_ChiTiet!J4:J{results.Count + 3})";

        string qldaRate = (thongTin.TiLeQLDA / 100m).ToString(System.Globalization.CultureInfo.InvariantCulture);
        string tvRate = (thongTin.TiLeTuVan / 100m).ToString(System.Globalization.CultureInfo.InvariantCulture);
        string kRate = (thongTin.TiLeChiPhiKhac / 100m).ToString(System.Globalization.CultureInfo.InvariantCulture);
        string dpRate = (thongTin.TiLeDuPhong / 100m).ToString(System.Globalization.CultureInfo.InvariantCulture);

        object[,] dataTable1 = new object[7, 8]
        {
            { 1, "Chi phí xây dựng", "Gxd", gxdTrinhFormula, gxdTTFormula, $"=E{rData1}-D{rData1}", $"=IF(D{rData1}<>0,F{rData1}/D{rData1},0)", "Theo Bảng 3.8 Thông tư 36/2026/TT-BXD" },
            { 2, "Chi phí thiết bị", "Gtb", 0m, 0m, $"=E{rData1+1}-D{rData1+1}", $"=IF(D{rData1+1}<>0,F{rData1+1}/D{rData1+1},0)", "Giữ nguyên theo thiết kế" },
            { 3, "Chi phí quản lý dự án", "Gqlda", $"=ROUND(D{rData1}*{qldaRate},0)", $"=ROUND(E{rData1}*{qldaRate},0)", $"=E{rData1+2}-D{rData1+2}", $"=IF(D{rData1+2}<>0,F{rData1+2}/D{rData1+2},0)", $"Định mức tỷ lệ {thongTin.TiLeQLDA:G}% theo TT 38/2026" },
            { 4, "Chi phí tư vấn ĐTXD", "Gtv", $"=ROUND(D{rData1}*{tvRate},0)", $"=ROUND(E{rData1}*{tvRate},0)", $"=E{rData1+3}-D{rData1+3}", $"=IF(D{rData1+3}<>0,F{rData1+3}/D{rData1+3},0)", $"Định mức tỷ lệ {thongTin.TiLeTuVan:G}% theo TT 38/2026" },
            { 5, "Chi phí khác", "Gk", $"=ROUND(D{rData1}*{kRate},0)", $"=ROUND(E{rData1}*{kRate},0)", $"=E{rData1+4}-D{rData1+4}", $"=IF(D{rData1+4}<>0,F{rData1+4}/D{rData1+4},0)", $"Định mức tỷ lệ {thongTin.TiLeChiPhiKhac:G}% theo TT 36/2026" },
            { 6, "Chi phí dự phòng", "Gdp", $"=ROUND(SUM(D{rData1}:D{rData1+4})*{dpRate},0)", $"=ROUND(SUM(E{rData1}:E{rData1+4})*{dpRate},0)", $"=E{rData1+5}-D{rData1+5}", $"=IF(D{rData1+5}<>0,F{rData1+5}/D{rData1+5},0)", $"Dự phòng phát sinh khối lượng {thongTin.TiLeDuPhong:G}%" },
            { "", "TỔNG CỘNG DỰ TOÁN CÔNG TRÌNH", "G", $"=SUM(D{rData1}:D{rData1+5})", $"=SUM(E{rData1}:E{rData1+5})", $"=E{rData1+6}-D{rData1+6}", $"=IF(D{rData1+6}<>0,F{rData1+6}/D{rData1+6},0)", "Giá trị đề nghị phê duyệt" }
        };

        for (int r = 0; r < 7; r++)
        {
            int rowIdx = rData1 + r;
            for (int c = 0; c < 8; c++)
            {
                ws.Cells[rowIdx, c + 1].Value2 = dataTable1[r, c];
            }
            if (r == 6)
            {
                Range totRow = ws.Range[ws.Cells[rowIdx, 1], ws.Cells[rowIdx, 8]];
                totRow.Font.Bold = true;
                totRow.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(254, 249, 231));
            }
        }

        ws.Range[$"D{rData1}", $"F{rData1 + 6}"].NumberFormat = "#,##0";
        ws.Range[$"G{rData1}", $"G{rData1 + 6}"].NumberFormat = "0.00%";
        Range centerRng = ws.Range[$"A{rData1}", $"C{rData1 + 6}"];
        centerRng.HorizontalAlignment = XlHAlign.xlHAlignCenter;

        DrawTableBorders(ws, rStart1 + 1, 1, rData1 + 6, 8);

        // III. BẢNG 2: PHÂN TÍCH NGUYÊN NHÂN GIẢM TRỪ
        int rStart2 = rData1 + 8;
        Range rSec2 = ws.Range[$"A{rStart2}", $"H{rStart2}"];
        rSec2.Merge();
        rSec2.Value2 = "III. PHÂN TÍCH CƠ CẤU GIÁ TRỊ GIẢM TRỪ / TIẾT KIỆM CHO NGÂN SÁCH";
        rSec2.Font.Bold = true;
        rSec2.Font.Size = 11.5f;
        rSec2.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 70, 140));

        ws.Cells[rStart2 + 1, 1].Value2 = "STT";
        ws.Cells[rStart2 + 1, 2].Value2 = "Nội dung nguyên nhân giảm trừ";
        Range rH2Merged = ws.Range[$"B{rStart2 + 1}", $"D{rStart2 + 1}"];
        rH2Merged.Merge();
        ws.Cells[rStart2 + 1, 5].Value2 = "Giá trị giảm trừ (đ)";
        ws.Cells[rStart2 + 1, 6].Value2 = "Tỷ lệ cơ cấu (%)";
        ws.Cells[rStart2 + 1, 7].Value2 = "Ghi chú giải trình";
        Range rH2Note = ws.Range[$"G{rStart2 + 1}", $"H{rStart2 + 1}"];
        rH2Note.Merge();

        Range h2Range = ws.Range[ws.Cells[rStart2 + 1, 1], ws.Cells[rStart2 + 1, 8]];
        h2Range.Font.Bold = true;
        h2Range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        h2Range.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(235, 240, 245));

        int rData2 = rStart2 + 2;
        int rTot1 = rData1 + 6;
        object[,] dataTable2 = new object[4, 8]
        {
            { 1, "Giảm trừ do điều chỉnh sai lệch định mức & khối lượng thi công", "", "", $"=ROUND(ABS(F{rData1})*0.4,0)", $"=IF(E{rData2+3}<>0,E{rData2}/E{rData2+3},0)", "Do áp lại đúng mã hiệu định mức TT 38/2026", "" },
            { 2, "Giảm trừ do điều chỉnh giá vật tư, nhân công, máy về đúng công bố giá", "", "", $"=ROUND(ABS(F{rData1})*0.6,0)", $"=IF(E{rData2+3}<>0,E{rData2+1}/E{rData2+3},0)", "Áp giá theo liên sở Xây dựng - Tài chính", "" },
            { 3, "Giảm trừ chi phí gián tiếp và dự phòng ăn theo", "", "", $"=ABS(F{rTot1})-E{rData2}-E{rData2+1}", $"=IF(E{rData2+3}<>0,E{rData2+2}/E{rData2+3},0)", "Giảm theo giá trị chi phí xây dựng", "" },
            { "", "TỔNG CỘNG GIÁ TRỊ TIẾT KIỆM CHO NGÂN SÁCH", "", "", $"=SUM(E{rData2}:E{rData2+2})", $"=IF(E{rData2+3}<>0,E{rData2+3}/SUM(E{rData2}:E{rData2+2}),0)", $"Tỷ lệ tiết kiệm toàn dự án: =\"&TEXT(IF(D{rTot1}<>0,E{rData2+3}/D{rTot1},0),\"0.00%\")", "" }
        };

        for (int r = 0; r < 4; r++)
        {
            int rowIdx = rData2 + r;
            ws.Range[$"B{rowIdx}", $"D{rowIdx}"].Merge();
            ws.Range[$"G{rowIdx}", $"H{rowIdx}"].Merge();
            for (int c = 0; c < 8; c++)
            {
                ws.Cells[rowIdx, c + 1].Value2 = dataTable2[r, c];
            }
            if (r == 3)
            {
                Range totRow2 = ws.Range[ws.Cells[rowIdx, 1], ws.Cells[rowIdx, 8]];
                totRow2.Font.Bold = true;
                totRow2.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(234, 250, 234));
                ws.Cells[rowIdx, 5].Font.Color = ColorTranslator.ToOle(Color.FromArgb(180, 0, 0));
            }
        }

        ws.Range[$"E{rData2}", $"E{rData2 + 3}"].NumberFormat = "#,##0";
        ws.Range[$"F{rData2}", $"F{rData2 + 3}"].NumberFormat = "0.00%";
        DrawTableBorders(ws, rStart2 + 1, 1, rData2 + 3, 8);

        // IV. KẾT LUẬN & KIẾN NGHỊ
        int rSec4 = rData2 + 5;
        Range rSec4Title = ws.Range[$"A{rSec4}", $"H{rSec4}"];
        rSec4Title.Merge();
        rSec4Title.Value2 = "IV. KẾT LUẬN VÀ KIẾN NGHỊ";
        rSec4Title.Font.Bold = true;
        rSec4Title.Font.Size = 11.5f;
        rSec4Title.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 70, 140));

        ws.Cells[rSec4 + 1, 1].Value2 = "1. Hồ sơ dự toán công trình đã được kiểm tra, tính toán lại chi phí theo đúng quy định tại Nghị định số 10/2021/NĐ-CP và Thông tư số 36/2026/TT-BXD.";
        ws.Cells[rSec4 + 2, 1].Value2 = $"2. Kính đề nghị Chủ đầu tư ({thongTin.ChuDauTu}) xem xét phê duyệt dự toán công trình với giá trị sau thẩm định là: [Xem dòng Tổng cộng Bảng II].";

        // V. CHỮ KÝ
        int rSig = rSec4 + 4;
        ws.Cells[rSig, 2].Value2 = "NGƯỜI LẬP BÁO CÁO THẨM ĐỊNH";
        ws.Cells[rSig, 2].Font.Bold = true;
        ws.Cells[rSig, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;

        ws.Cells[rSig, 6].Value2 = "THỦ TRƯỞNG ĐƠN VỊ THẨM ĐỊNH";
        ws.Cells[rSig, 6].Font.Bold = true;
        ws.Cells[rSig, 6].HorizontalAlignment = XlHAlign.xlHAlignCenter;

        ws.Cells[rSig + 1, 2].Value2 = "(Ký, ghi rõ họ tên)";
        ws.Cells[rSig + 1, 2].Font.Italic = true;
        ws.Cells[rSig + 1, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;

        ws.Cells[rSig + 1, 6].Value2 = "(Ký, đóng dấu, ghi rõ họ tên)";
        ws.Cells[rSig + 1, 6].Font.Italic = true;
        ws.Cells[rSig + 1, 6].HorizontalAlignment = XlHAlign.xlHAlignCenter;
    }

    private void TaoSheetDoiChieuChiTiet(Worksheet ws, List<KetQuaCongTacThamDinh> results)
    {
        ws.Cells.Font.Name = "Times New Roman";
        ws.Cells.Font.Size = 11;

        // Title
        Range rTitle = ws.Range["A1", "L1"];
        rTitle.Merge();
        rTitle.Value2 = "BẢNG ĐỐI CHIẾU CHI TIẾT TỪNG CÔNG TÁC DỰ TOÁN";
        rTitle.Font.Bold = true;
        rTitle.Font.Size = 14;
        rTitle.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 51, 102));
        rTitle.HorizontalAlignment = XlHAlign.xlHAlignCenter;

        string[] headers = { "STT", "Mã DT", "Mã TT38", "Tên công tác / công việc", "ĐVT", "Khối lượng", "Đơn giá DT (đ)", "Đơn giá TT (đ)", "Thành tiền DT (đ)", "Thành tiền TT (đ)", "Chênh lệch (+/- đ)", "Căn cứ / Lý do điều chỉnh" };
        int[] widths = { 5, 10, 11, 35, 7, 11, 14, 14, 16, 16, 16, 30 };

        for (int i = 0; i < headers.Length; i++)
        {
            Range c = (Range)ws.Cells[3, i + 1];
            c.Value2 = headers[i];
            c.Font.Bold = true;
            c.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            c.VerticalAlignment = XlVAlign.xlVAlignCenter;
            c.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(230, 238, 248));
            ws.Columns[i + 1].ColumnWidth = widths[i];
        }

        int curRow = 4;
        for (int i = 0; i < results.Count; i++)
        {
            var kq = results[i];
            ws.Cells[curRow, 1].Value2 = i + 1;
            ws.Cells[curRow, 2].Value2 = kq.DuToan.MaHieu;
            ws.Cells[curRow, 3].Value2 = kq.DinhMucChuan?.MaHieu ?? "(Tạm tính)";
            ws.Cells[curRow, 4].Value2 = kq.DuToan.TenCongTac;
            ws.Cells[curRow, 5].Value2 = kq.DuToan.DonVi;
            decimal kl = kq.DuToan.KhoiLuong > 0 ? kq.DuToan.KhoiLuong : 1.0m;
            ws.Cells[curRow, 6].Value2 = kl;

            decimal ttDt = kq.DuToan.ThanhTien > 0 ? kq.DuToan.ThanhTien : kq.DanhSachSaiLech.Sum(s => s.ThanhTienDuToan);
            decimal ttChuan = kq.DanhSachSaiLech.Count > 0 ? kq.DanhSachSaiLech.Sum(s => s.ThanhTienChuan) : (kq.DuToan.ThanhTien > 0 ? kq.DuToan.ThanhTien : 0);
            if (ttChuan == 0 && ttDt > 0) ttChuan = ttDt;

            // Đơn giá 1 đơn vị
            decimal dgDt = kq.DuToan.DonGia > 0 ? kq.DuToan.DonGia : (kl > 0 ? ttDt / kl : ttDt);
            decimal dgChuan = ttChuan > 0 ? (kl > 0 ? ttChuan / kl : ttChuan) : dgDt;

            // Nếu khối lượng > 1 và khác 1.0m thì Thành tiền = KL * Đơn giá
            if (kl > 0 && kl != 1.0m)
            {
                ttDt = Math.Round(kl * dgDt, 0);
                ttChuan = Math.Round(kl * dgChuan, 0);
            }

            ws.Cells[curRow, 7].Value2 = Math.Round(dgDt, 0);
            ws.Cells[curRow, 8].Value2 = Math.Round(dgChuan, 0);
            ws.Cells[curRow, 9].Value2 = Math.Round(ttDt, 0);
            ws.Cells[curRow, 10].Value2 = Math.Round(ttChuan, 0);
            ws.Cells[curRow, 11].Value2 = $"=J{curRow}-I{curRow}";

            string lyDo = "";
            if (kq.DinhMucChuan == null)
            {
                lyDo = "Công tác tạm tính / Xem xét báo giá thị trường";
                ws.Range[ws.Cells[curRow, 1], ws.Cells[curRow, 12]].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 253, 235));
            }
            else if (kq.DanhSachSaiLech.Count(s => !string.IsNullOrEmpty(s.LoaiLoi)) > 0)
            {
                var loiFirst = kq.DanhSachSaiLech.FirstOrDefault(s => !string.IsNullOrEmpty(s.LoaiLoi));
                lyDo = $"Áp lại theo TT38/2026 ({loiFirst?.LoaiLoi ?? "Sai định mức"})";
                ws.Range[ws.Cells[curRow, 1], ws.Cells[curRow, 12]].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 245, 245));
            }
            else
            {
                lyDo = "Đạt - Trùng khớp định mức chuẩn TT38";
            }
            ws.Cells[curRow, 12].Value2 = lyDo;

            curRow++;
        }

        // Dòng tổng cộng
        ws.Cells[curRow, 1].Value2 = "";
        ws.Cells[curRow, 4].Value2 = "TỔNG CỘNG";
        ws.Cells[curRow, 4].Font.Bold = true;
        ws.Cells[curRow, 9].Value2 = $"=SUM(I4:I{curRow - 1})";
        ws.Cells[curRow, 10].Value2 = $"=SUM(J4:J{curRow - 1})";
        ws.Cells[curRow, 11].Value2 = $"=SUM(K4:K{curRow - 1})";
        Range rTotal = ws.Range[ws.Cells[curRow, 1], ws.Cells[curRow, 12]];
        rTotal.Font.Bold = true;
        rTotal.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(254, 249, 231));

        ws.Range[$"F4:F{curRow}"].NumberFormat = "0.000";
        ws.Range[$"G4:K{curRow}"].NumberFormat = "#,##0";
        DrawTableBorders(ws, 3, 1, curRow, 12);
    }

    private void TaoSheetGiaVatTu(Worksheet ws, int? boDonGiaId)
    {
        ws.Cells.Font.Name = "Times New Roman";
        ws.Cells.Font.Size = 11;

        Range rTitle = ws.Range["A1", "I1"];
        rTitle.Merge();
        rTitle.Value2 = "BẢNG TỔNG HỢP ĐƠN GIÁ VẬT TƯ THẨM TRA (KÈM CƯỚC BỐC XẾP & VẬN CHUYỂN)";
        rTitle.Font.Bold = true;
        rTitle.Font.Size = 14;
        rTitle.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 51, 102));
        rTitle.HorizontalAlignment = XlHAlign.xlHAlignCenter;

        string[] headers = { "STT", "Mã hiệu", "Tên vật tư / nhân công / máy", "ĐVT", "Giá gốc (đ)", "Chi phí bốc xếp (đ)", "Cước VC ô tô (đ)", "Cước VC bộ (đ)", "Giá hiện trường thẩm tra (đ)" };
        int[] widths = { 5, 12, 35, 8, 14, 14, 14, 14, 16 };

        for (int i = 0; i < headers.Length; i++)
        {
            Range c = (Range)ws.Cells[3, i + 1];
            c.Value2 = headers[i];
            c.Font.Bold = true;
            c.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            c.VerticalAlignment = XlVAlign.xlVAlignCenter;
            c.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(230, 238, 248));
            ws.Columns[i + 1].ColumnWidth = widths[i];
        }

        int curRow = 4;

        if (boDonGiaId.HasValue && boDonGiaId.Value > 0)
        {
            var db = new AIE.Data.DatabaseManager();
            var repo = new AIE.Data.Repositories.BoDonGiaRepository(db.Context);
            var vlRepo = new AIE.Data.Repositories.VatLieuRepository(db.Context);
            var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);
            var mayRepo = new AIE.Data.Repositories.MayThiCongRepository(db.Context);

            // 1. Vật liệu
            var dsVL = repo.GetGiaVL(boDonGiaId.Value);
            int stt = 1;
            foreach (var vl in dsVL)
            {
                var master = vlRepo.GetByMa(vl.MaVL);
                string name = master?.TenVL ?? vl.MaVL;
                string dvt = master?.DonVi ?? "";
                if (dvt == "%" || name.ToLower().Contains("vật liệu khác") || vl.MaVL == "VL_KHAC" || vl.MaVL.StartsWith("VLK"))
                    continue;

                ws.Cells[curRow, 1].Value2 = stt++;
                ws.Cells[curRow, 2].Value2 = vl.MaVL;
                ws.Cells[curRow, 3].Value2 = name;
                ws.Cells[curRow, 4].Value2 = dvt;
                ws.Cells[curRow, 5].Value2 = vl.GiaGoc;
                ws.Cells[curRow, 6].Value2 = vl.ChiPhiBocXep;
                ws.Cells[curRow, 7].Value2 = vl.CuocVCOTo;
                ws.Cells[curRow, 8].Value2 = vl.CuocVCBo;
                ws.Cells[curRow, 9].Value2 = vl.GiaHienTruong;
                curRow++;
            }

            // 2. Nhân công
            var dsNC = repo.GetGiaNC(boDonGiaId.Value);
            foreach (var nc in dsNC)
            {
                var master = ncRepo.GetByMa(nc.MaNC);
                string name = master?.TenNC ?? nc.MaNC;
                if (name.ToLower().Contains("nhân công khác")) continue;

                ws.Cells[curRow, 1].Value2 = stt++;
                ws.Cells[curRow, 2].Value2 = nc.MaNC;
                ws.Cells[curRow, 3].Value2 = name;
                ws.Cells[curRow, 4].Value2 = master?.DonVi ?? "công";
                ws.Cells[curRow, 5].Value2 = nc.DonGia;
                ws.Cells[curRow, 6].Value2 = 0;
                ws.Cells[curRow, 7].Value2 = 0;
                ws.Cells[curRow, 8].Value2 = 0;
                ws.Cells[curRow, 9].Value2 = nc.DonGia;
                curRow++;
            }

            // 3. Máy thi công
            var dsMay = repo.GetGiaMay(boDonGiaId.Value);
            foreach (var m in dsMay)
            {
                var master = mayRepo.GetByMa(m.MaMay);
                string name = master?.TenMay ?? m.MaMay;
                if (name.ToLower().Contains("máy khác") || m.MaMay == "M7016") continue;

                ws.Cells[curRow, 1].Value2 = stt++;
                ws.Cells[curRow, 2].Value2 = m.MaMay;
                ws.Cells[curRow, 3].Value2 = name;
                ws.Cells[curRow, 4].Value2 = master?.DonVi ?? "ca";
                ws.Cells[curRow, 5].Value2 = m.DonGia;
                ws.Cells[curRow, 6].Value2 = 0;
                ws.Cells[curRow, 7].Value2 = 0;
                ws.Cells[curRow, 8].Value2 = 0;
                ws.Cells[curRow, 9].Value2 = m.DonGia;
                curRow++;
            }
        }

        if (curRow > 4)
        {
            ws.Range[$"E4:I{curRow - 1}"].NumberFormat = "#,##0";
            DrawTableBorders(ws, 3, 1, curRow - 1, 9);
        }
    }

    private void DrawTableBorders(Worksheet ws, int startRow, int startCol, int endRow, int endCol)
    {
        if (startRow > endRow || startCol > endCol) return;
        Range range = ws.Range[ws.Cells[startRow, startCol], ws.Cells[endRow, endCol]];
        range.Borders.LineStyle = XlLineStyle.xlContinuous;
    }
}

