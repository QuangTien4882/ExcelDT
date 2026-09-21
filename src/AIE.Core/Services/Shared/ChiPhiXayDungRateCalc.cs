using System;
using System.Collections.Generic;
using System.Linq;
using AIE.Core.Models;
using AIE.Core.Services.Shared;

namespace AIE.Core.Services
{
    /// <summary>
    /// Kết quả tính tỷ lệ % chi phí xây dựng (CPC, TT, TNCTTT, GTGT, Nhà tạm).
    /// </summary>
    public class ChiPhiXayDungRateResult
    {
        public decimal TiLeCPC { get; set; }
        public decimal TiLeTT { get; set; }
        public decimal TiLeTNCTTT { get; set; }
        public decimal TiLeGTGT { get; set; }
        public decimal TiLeNhaTam { get; set; }

        /// <summary>Danh sách cảnh báo khi phải dùng giá trị mặc định/thiếu CSDL (rỗng = tính chuẩn).</summary>
        public List<string> CanhBao { get; set; } = new();

        public bool CoCanhBao => CanhBao.Count > 0;
    }

    /// <summary>
    /// Tính tỷ lệ % chi phí xây dựng theo Thông tư 36/2026 &amp; 38/2026 — thuần túy, không phụ thuộc CSDL/Excel
    /// (dễ kiểm thử tự động). Nguồn dữ liệu định mức theo quy mô (Bảng 3.3/3.5) do tầng gọi nạp sẵn từ CSDL
    /// và truyền vào; nếu thiếu sẽ tự dùng bảng chuẩn (Bảng 3.2/3.6/3.7) và báo cảnh báo.
    /// </summary>
    public static class ChiPhiXayDungRateCalc
    {
        public const decimal TyDong = 1_000_000_000m;

        /// <summary>
        /// Tính quy mô chi phí XD trước thuế (tỷ đồng) dùng để nội suy tỷ lệ:
        /// dùng GXDTT có sẵn; nếu = 0 thì ước lượng từ tổng chi phí trực tiếp T và các tỷ lệ đã có.
        /// </summary>
        public static decimal TinhQuyMoTuChiPhiXD(decimal gxdtt, decimal tongT, decimal tiLeCPC, decimal tiLeTT, decimal tiLeTNCTTT)
        {
            decimal value = gxdtt;
            if (value <= 0m)
            {
                decimal estT = tongT > 0m ? tongT : 0m;
                if (estT > 0m)
                {
                    decimal c0 = tiLeCPC > 0m ? tiLeCPC : 7.3m;
                    decimal t0 = tiLeTT > 0m ? tiLeTT : 2.5m;
                    decimal n0 = tiLeTNCTTT > 0m ? tiLeTNCTTT : 5.5m;
                    value = estT * (1m + (c0 + t0) / 100m) * (1m + n0 / 100m);
                }
            }
            return value > 0m ? Math.Round(value / TyDong, 4, MidpointRounding.AwayFromZero) : 0m;
        }

        /// <summary>
        /// Tính 5 tỷ lệ % chi phí xây dựng.
        /// </summary>
        /// <param name="loaiCongTrinh">Loại công trình (Dân dụng, Giao thông,...); rỗng → mặc định "Dân dụng"</param>
        /// <param name="phanLoai">Phân loại phụ (di tích, hầm,...) — có thể null</param>
        /// <param name="gxdtt">Chi phí XD trước thuế (đồng); 0 = chưa biết (ước lượng từ tongT)</param>
        /// <param name="tongT">Tổng chi phí trực tiếp T = VL + NC + M (đồng)</param>
        /// <param name="soBuocThietKe">0/2/3 bước (lập dự toán), 1 bước (BC KT-KT)</param>
        /// <param name="laVungSauXa">Công trình vùng sâu, xa, núi, biên giới, hải đảo → CPC ×1.1</param>
        /// <param name="loaiCongTrinhNhaTam">"Công trình xây dựng theo tuyến" hoặc "Công trình xây dựng còn lại"</param>
        /// <param name="tiLeCPC">Tỷ lệ CPC hiện có (giữ nguyên nếu &gt; 0)</param>
        /// <param name="tiLeTT">Tỷ lệ TT hiện có (giữ nguyên nếu &gt; 0)</param>
        /// <param name="tiLeTNCTTT">Tỷ lệ TNCTTT hiện có (giữ nguyên nếu &gt; 0)</param>
        /// <param name="tiLeGTGT">Thuế suất GTGT hiện có (giữ nguyên nếu &gt; 0; mặc định 10%)</param>
        /// <param name="tiLeNhaTam">Tỷ lệ nhà tạm hiện có (giữ nguyên nếu &gt; 0)</param>
        /// <param name="cpcTheoQuyMo">Định mức CPC Bảng 3.3 (từ CSDL); null/rỗng → dùng Bảng 3.2 tỷ lệ cố định</param>
        /// <param name="ttTheoLoai">Định mức TT Bảng 3.5 (từ CSDL); null → mặc định 2.5%</param>
        public static ChiPhiXayDungRateResult TinhTyLe(
            string loaiCongTrinh,
            string? phanLoai,
            decimal gxdtt,
            decimal tongT,
            int soBuocThietKe,
            bool laVungSauXa,
            string loaiCongTrinhNhaTam,
            decimal tiLeCPC,
            decimal tiLeTT,
            decimal tiLeTNCTTT,
            decimal tiLeGTGT,
            decimal tiLeNhaTam,
            IReadOnlyList<DinhMucCPC>? cpcTheoQuyMo = null,
            DinhMucTT? ttTheoLoai = null)
        {
            var r = new ChiPhiXayDungRateResult();
            string loaiCT = string.IsNullOrEmpty(loaiCongTrinh) ? "Dân dụng" : loaiCongTrinh;
            string? phanLoaiCT = string.IsNullOrEmpty(phanLoai) ? null : phanLoai;

            if (string.IsNullOrEmpty(loaiCongTrinh))
                r.CanhBao.Add("Loại công trình chưa xác định → mặc định tính theo \"Dân dụng\".");

            decimal quyMo = TinhQuyMoTuChiPhiXD(gxdtt, tongT, tiLeCPC, tiLeTT, tiLeTNCTTT);
            if (quyMo <= 0m)
                r.CanhBao.Add("Không xác định được quy mô chi phí XD (GXDTT và T đều bằng 0) → tỷ lệ CPC lấy mức cố định.");

            // CPC theo quy mô (Bảng 3.3) nếu có CSDL, ngược lại tỷ lệ cố định Bảng 3.2
            decimal cpc = tiLeCPC;
            if (cpc <= 0m && (cpcTheoQuyMo == null || cpcTheoQuyMo.Count == 0))
            {
                cpc = InterpolationHelper.LayTiLeCPCTMDT(loaiCT, phanLoaiCT);
                r.CanhBao.Add($"Không có bảng định mức Chi phí chung theo quy mô (Bảng 3.3) → dùng tỷ lệ cố định {cpc}%.");
            }
            else if (cpc <= 0m)
            {
                if (soBuocThietKe == 1)
                {
                    cpc = cpcTheoQuyMo!.OrderBy(x => x.QuyMoMin).First().TiLe;
                }
                else if (soBuocThietKe >= 3)
                {
                    cpc = InterpolationHelper.LayTiLeCPCTMDT(loaiCT, phanLoaiCT);
                }
                else
                {
                    cpc = quyMo > 0m
                        ? InterpolationHelper.NoiSuyTiLeCPC(cpcTheoQuyMo!, quyMo)
                        : InterpolationHelper.LayTiLeCPCTMDT(loaiCT, phanLoaiCT);
                }
            }

            // TT (Bảng 3.5)
            decimal tt = tiLeTT;
            if (tt <= 0m)
            {
                if (ttTheoLoai != null && ttTheoLoai.TiLe > 0m)
                    tt = ttTheoLoai.TiLe;
                else
                {
                    tt = 2.5m;
                    r.CanhBao.Add("Không có định mức Chi phí công việc không xác định KL (Bảng 3.5) → dùng mặc định 2.5%.");
                }
            }

            // TNCTTT (Bảng 3.6)
            decimal tncttt = tiLeTNCTTT;
            if (tncttt <= 0m)
            {
                tncttt = InterpolationHelper.LayTiLeTNCTTT(loaiCT);
                if (tncttt <= 0m)
                    tncttt = (loaiCT == "Công nghiệp" || loaiCT == "Giao thông") ? 6.0m : 5.5m;
            }

            decimal gtgt = tiLeGTGT > 0m ? tiLeGTGT : 10.0m;

            // Nhà tạm (Bảng 3.7)
            decimal nhaTam = tiLeNhaTam;
            string loaiNhaTamKq = !string.IsNullOrEmpty(loaiCongTrinhNhaTam) ? loaiCongTrinhNhaTam : "Công trình xây dựng còn lại";
            if (nhaTam <= 0m)
            {
                nhaTam = InterpolationHelper.NoiSuyTiLeNhaTam(loaiNhaTamKq, quyMo);
                if (nhaTam <= 0m)
                {
                    nhaTam = 1.2m;
                    r.CanhBao.Add("Không xác định được tỷ lệ nhà tạm (Bảng 3.7) → dùng mức 1.2%.");
                }
            }

            // Khoản 3.1.4 TT 36: vùng núi, biên giới, hải đảo ×1.1
            if (laVungSauXa && cpc > 0m)
                cpc = Math.Round(cpc * 1.1m, 3, MidpointRounding.AwayFromZero);

            r.TiLeCPC = cpc;
            r.TiLeTT = tt;
            r.TiLeTNCTTT = tncttt;
            r.TiLeGTGT = gtgt;
            r.TiLeNhaTam = nhaTam;
            return r;
        }
    }
}