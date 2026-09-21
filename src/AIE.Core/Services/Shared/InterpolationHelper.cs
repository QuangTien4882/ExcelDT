using AIE.Core.Models;

namespace AIE.Core.Services.Shared;

/// <summary>
/// Nội suy tuyến tính cho định mức % QLDA / Tư vấn.
/// Khi quy mô nằm giữa 2 mốc, nội suy tuyến tính để tính tỉ lệ % chính xác.
/// </summary>
public static class InterpolationHelper
{
    /// <summary>
    /// Tìm tỉ lệ % QLDA theo quy mô dự án.
    /// Nội suy tuyến tính nếu quy mô nằm giữa 2 mốc.
    /// </summary>
    /// <param name="danhSachDinhMuc">Danh sách định mức % theo mốc quy mô (đã sắp xếp tăng)</param>
    /// <param name="quyMo">Quy mô dự án (tỷ đồng)</param>
    /// <returns>Tỉ lệ % sau nội suy</returns>
    public static decimal NoiSuyTiLe(IReadOnlyList<DinhMucQLDA> danhSachDinhMuc, decimal quyMo)
    {
        if (danhSachDinhMuc.Count == 0)
            throw new ArgumentException("Danh sách định mức rỗng.");

        // Sắp xếp theo QuyMoMin tăng dần
        var sorted = danhSachDinhMuc.OrderBy(x => x.QuyMoMin).ToList();

        // Nếu quy mô nhỏ hơn mốc đầu tiên → lấy tỉ lệ đầu
        if (quyMo <= sorted[0].QuyMoMin)
            return sorted[0].TiLe;

        // Nếu quy mô lớn hơn mốc cuối → lấy tỉ lệ cuối
        var last = sorted[sorted.Count - 1];
        if (last.QuyMoMax.HasValue && quyMo >= last.QuyMoMax.Value)
            return last.TiLe;

        // Tìm 2 mốc kẹp
        for (int i = 0; i < sorted.Count; i++)
        {
            var dm = sorted[i];
            decimal min = dm.QuyMoMin;
            decimal max = dm.QuyMoMax ?? decimal.MaxValue;

            if (quyMo >= min && quyMo <= max)
            {
                // Nằm đúng trong 1 khoảng → lấy tỉ lệ đó
                // Nhưng nếu có mốc kế tiếp, nội suy giữa 2 mốc
                if (i + 1 < sorted.Count && quyMo > min)
                {
                    var next = sorted[i + 1];
                    return NoiSuyTuyenTinh(min, dm.TiLe, next.QuyMoMin, next.TiLe, quyMo);
                }
                return dm.TiLe;
            }
        }

        // Fallback: lấy mốc cuối
        return sorted[sorted.Count - 1].TiLe;
    }

    /// <summary>
    /// Tìm tỉ lệ % tư vấn theo quy mô.
    /// </summary>
    public static decimal NoiSuyTiLeTuVan(IReadOnlyList<DinhMucTuVan> danhSachDinhMuc, decimal quyMo)
    {
        if (danhSachDinhMuc.Count == 0)
            throw new ArgumentException("Danh sách định mức rỗng.");

        var sorted = danhSachDinhMuc.OrderBy(x => x.QuyMoMin).ToList();

        if (quyMo <= sorted[0].QuyMoMin)
            return sorted[0].TiLe;

        var last = sorted[sorted.Count - 1];
        if (last.QuyMoMax.HasValue && quyMo >= last.QuyMoMax.Value)
            return last.TiLe;

        for (int i = 0; i < sorted.Count; i++)
        {
            var dm = sorted[i];
            decimal min = dm.QuyMoMin;
            decimal max = dm.QuyMoMax ?? decimal.MaxValue;

            if (quyMo >= min && quyMo <= max)
            {
                if (i + 1 < sorted.Count && quyMo > min)
                {
                    var next = sorted[i + 1];
                    return NoiSuyTuyenTinh(min, dm.TiLe, next.QuyMoMin, next.TiLe, quyMo);
                }
                return dm.TiLe;
            }
        }

        return sorted[sorted.Count - 1].TiLe;
    }

    /// <summary>
    /// Nội suy tuyến tính giữa 2 điểm.
    /// </summary>
    public static decimal NoiSuyTuyenTinh(decimal x1, decimal y1, decimal x2, decimal y2, decimal x)
    {
        if (x2 == x1) return y1;
        decimal result = y1 + (y2 - y1) * (x - x1) / (x2 - x1);
        return Math.Round(result, 3);
    }

    /// <summary>
    /// Tìm tỉ lệ % CPC theo quy mô chi phí XD.
    /// Nội suy theo công thức Thông tư 36/2026/TT-BXD:
    /// N_t = N_d - (N_d - N_t) * (G_t - G_d) / (G_t - G_d)
    /// </summary>
    public static decimal NoiSuyTiLeCPC(IReadOnlyList<DinhMucCPC> danhSachDinhMuc, decimal quyMo)
    {
        if (danhSachDinhMuc.Count == 0)
            throw new ArgumentException("Danh sách định mức rỗng.");

        var sorted = danhSachDinhMuc.OrderBy(x => x.QuyMoMax ?? decimal.MaxValue).ToList();

        // Nếu quy mô nhỏ hơn hoặc bằng mốc nhỏ nhất
        if (quyMo <= (sorted[0].QuyMoMax ?? 0))
            return sorted[0].TiLe;

        for (int i = 1; i < sorted.Count; i++)
        {
            var current = sorted[i];
            var prev = sorted[i - 1];
            
            decimal maxPrev = prev.QuyMoMax ?? 0;
            decimal maxCurrent = current.QuyMoMax ?? decimal.MaxValue;

            if (quyMo > maxPrev && quyMo <= maxCurrent)
            {
                // Nội suy tuyến tính giữa (maxPrev, prev.TiLe) và (maxCurrent, current.TiLe)
                if (maxCurrent == decimal.MaxValue) 
                    return current.TiLe; // Không có mốc trên, giữ nguyên

                return NoiSuyTuyenTinh(maxPrev, prev.TiLe, maxCurrent, current.TiLe, quyMo);
            }
        }

        // Lớn hơn mốc cuối cùng (trường hợp > 1000)
        return sorted[sorted.Count - 1].TiLe;
    }

    /// <summary>
    /// Tìm tỉ lệ % chi phí nhà tạm để ở và điều hành thi công theo Bảng 3.7 TT 36/2026/TT-BXD.
    /// Nội suy tuyến tính theo quy mô chi phí XD trước thuế trong TMĐT (tỷ đồng).
    /// </summary>
    /// <param name="loaiNhaTam">"Công trình xây dựng theo tuyến" hoặc "Công trình xây dựng còn lại"</param>
    /// <param name="quyMo">Quy mô chi phí XD trước thuế trong TMĐT (tỷ đồng)</param>
    /// <returns>Tỷ lệ % chi phí nhà tạm</returns>
    public static decimal NoiSuyTiLeNhaTam(string loaiNhaTam, decimal quyMo)
    {
        bool laTheoTuyen = !string.IsNullOrEmpty(loaiNhaTam) && 
                           loaiNhaTam.IndexOf("tuyến", StringComparison.OrdinalIgnoreCase) >= 0;

        // Các mốc Bảng 3.7: ≤15, ≤100, ≤500, ≤1000, >1000 (tỷ đồng)
        (decimal Moc, decimal TiLe)[] mocs = laTheoTuyen
            ? new (decimal, decimal)[] { (15m, 2.2m), (100m, 2.0m), (500m, 1.9m), (1000m, 1.8m), (decimal.MaxValue, 1.7m) }
            : new (decimal, decimal)[] { (15m, 1.1m), (100m, 1.0m), (500m, 0.95m), (1000m, 0.9m), (decimal.MaxValue, 0.85m) };

        if (quyMo <= mocs[0].Moc)
            return mocs[0].TiLe;

        for (int i = 1; i < mocs.Length; i++)
        {
            var prev = mocs[i - 1];
            var curr = mocs[i];

            if (quyMo <= curr.Moc)
            {
                if (curr.Moc == decimal.MaxValue)
                    return curr.TiLe;

                return NoiSuyTuyenTinh(prev.Moc, prev.TiLe, curr.Moc, curr.TiLe, quyMo);
            }
        }

        return mocs[mocs.Length - 1].TiLe;
    }

    /// <summary>
    /// Lấy định mức tỷ lệ Chi phí chung giai đoạn Tổng mức đầu tư (Bảng 3.2 TT 36/2026/TT-BXD).
    /// Tỷ lệ cố định theo loại công trình, không phụ thuộc quy mô chi phí.
    /// </summary>
    public static decimal LayTiLeCPCTMDT(string loaiCongTrinh, string? phanLoaiPhu = null)
    {
        string normCT = DinhMucTT38Database.ChuanHoaLoaiCongTrinh(loaiCongTrinh);
        string normPhu = phanLoaiPhu?.Trim() ?? "";

        if (normCT.IndexOf("Dân dụng", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            if (normPhu.IndexOf("di tích", StringComparison.OrdinalIgnoreCase) >= 0)
                return 11.6m;
            return 7.3m;
        }

        if (normCT.IndexOf("Công nghiệp", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            if (normPhu.IndexOf("hầm", StringComparison.OrdinalIgnoreCase) >= 0)
                return 7.3m;
            return 6.2m;
        }

        if (normCT.IndexOf("Giao thông", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            if (normPhu.IndexOf("hầm", StringComparison.OrdinalIgnoreCase) >= 0)
                return 7.3m;
            return 6.2m;
        }

        if (normCT.IndexOf("Nông nghiệp", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            if (normPhu.IndexOf("hầm", StringComparison.OrdinalIgnoreCase) >= 0)
                return 7.3m;
            return 6.1m;
        }

        if (normCT.IndexOf("Hạ tầng", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return 5.5m;
        }

        return 7.3m; // Mặc định dân dụng
    }

    /// <summary>
    /// Lấy định mức tỷ lệ Thu nhập chịu thuế tính trước (Bảng 3.6 TT 36/2026/TT-BXD).
    /// </summary>
    public static decimal LayTiLeTNCTTT(string loaiCongTrinh)
    {
        string norm = DinhMucTT38Database.ChuanHoaLoaiCongTrinh(loaiCongTrinh);
        if (norm.IndexOf("Công nghiệp", StringComparison.OrdinalIgnoreCase) >= 0 ||
            norm.IndexOf("Giao thông", StringComparison.OrdinalIgnoreCase) >= 0 ||
            norm.IndexOf("Lắp đặt thiết bị", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return 6.0m;
        }

        return 5.5m; // Dân dụng, Nông nghiệp & Môi trường, Hạ tầng kỹ thuật
    }
}

