using AIE.Core.Models;
using AIE.Core.Enums;
using AIE.Data.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace AIE.ExcelAddIn.Services;

public class ThamDinhEngine
{
    private readonly CongTacRepository _repository;
    private readonly VatLieuRepository _vlRepo;
    private readonly NhanCongRepository _ncRepo;
    private readonly MayThiCongRepository _mayRepo;

    public ThamDinhEngine(
        CongTacRepository repository,
        VatLieuRepository vlRepo,
        NhanCongRepository ncRepo,
        MayThiCongRepository mayRepo)
    {
        _repository = repository;
        _vlRepo = vlRepo;
        _ncRepo = ncRepo;
        _mayRepo = mayRepo;
    }

    public List<KetQuaCongTacThamDinh> KiemTra(List<CongTacThamDinh> duToanList, int? boDonGiaId = null, Vung vung = Vung.VungII)
    {
        var ketQua = new List<KetQuaCongTacThamDinh>();
        
        var db = new AIE.Data.DatabaseManager();
        var dmMayRepo = new DinhMucCaMayRepository(db.Context);
        var allNcList = _ncRepo.GetAll().ToList();

        var giaVlDict = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        var giaNcDict = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        var giaMayDict = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

        decimal giaXang = 22150m, giaDiezel = 28090m, giaDien = 2204m;

        if (boDonGiaId.HasValue && boDonGiaId.Value > 0)
        {
            var bdgRepo = new BoDonGiaRepository(db.Context);
            var boInfo = bdgRepo.GetById(boDonGiaId.Value);
            if (boInfo != null)
            {
                if (boInfo.GiaXang > 0) giaXang = boInfo.GiaXang;
                if (boInfo.GiaDiezel > 0) giaDiezel = boInfo.GiaDiezel;
                if (boInfo.GiaDien > 0) giaDien = boInfo.GiaDien;
            }

            foreach (var vl in bdgRepo.GetGiaVL(boDonGiaId.Value))
            {
                if (!string.IsNullOrEmpty(vl.MaVL)) giaVlDict[vl.MaVL.Trim()] = vl.GiaHienTruong;
            }

            foreach (var nc in bdgRepo.GetGiaNC(boDonGiaId.Value))
            {
                if (!string.IsNullOrEmpty(nc.MaNC)) giaNcDict[nc.MaNC.Trim()] = nc.DonGia;
            }

            foreach (var m in bdgRepo.GetGiaMay(boDonGiaId.Value))
            {
                if (!string.IsNullOrEmpty(m.MaMay)) giaMayDict[m.MaMay.Trim()] = m.DonGia;
            }
        }
        else
        {
            try
            {
                var settingsDir = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "AIE_DuToan");
                var fuelFile = System.IO.Path.Combine(settingsDir, "fuel_prices.json");
                if (System.IO.File.Exists(fuelFile))
                {
                    var json = System.IO.File.ReadAllText(fuelFile);
                    dynamic doc = Newtonsoft.Json.JsonConvert.DeserializeObject(json);
                    if (doc != null)
                    {
                        if ((decimal)doc.GiaXang > 0) giaXang = (decimal)doc.GiaXang;
                        if ((decimal)doc.GiaDiezel > 0) giaDiezel = (decimal)doc.GiaDiezel;
                        if ((decimal)doc.GiaDien > 0) giaDien = (decimal)doc.GiaDien;
                    }
                }
            }
            catch { }
        }

        foreach (var ctDuToan in duToanList)
        {
            var kq = new KetQuaCongTacThamDinh { DuToan = ctDuToan };
            var ctChuan = _repository.GetByMaHieu(ctDuToan.MaHieu);

            if (ctChuan == null)
            {
                kq.DanhSachSaiLech.Add(new SaiLechDinhMuc
                {
                    LoaiLoi = "Mã không tồn tại",
                    MoTa = $"Mã hiệu '{ctDuToan.MaHieu}' không tồn tại trong CSDL định mức.",
                    SoDongExcel = ctDuToan.SoDongExcel
                });
                ketQua.Add(kq);
                continue;
            }

            kq.DinhMucChuan = ctChuan;

            // Tạo bản sao danh sách chuẩn để đánh dấu những mục đã map
            var dsHpChuan = ctChuan.DanhSachHaoPhi.ToList();

            foreach (var hpDuToan in ctDuToan.DanhSachHaoPhi)
            {
                var hpChuanMatched = TimHaoPhiTuongDuong(hpDuToan, dsHpChuan, ctChuan);
                
                var saiLech = new SaiLechDinhMuc
                {
                    HaoPhiDuToan = hpDuToan,
                    LoaiHP = hpDuToan.Loai,
                    SoDongExcel = hpDuToan.SoDongExcel
                };

                if (hpChuanMatched == null)
                {
                    saiLech.LoaiLoi = "Hao phí thừa";
                    saiLech.MoTa = $"Dự toán có '{hpDuToan.TenHaoPhi}' nhưng TT38 không có (hoặc tên không khớp).";
                }
                else
                {
                    dsHpChuan.Remove(hpChuanMatched); // Đã map
                    
                    saiLech.HaoPhiChuan = hpChuanMatched;
                    saiLech.DonGiaChuan = GetDonGiaChuan(hpChuanMatched, vung, giaVlDict, giaNcDict, giaMayDict, allNcList, dmMayRepo, giaXang, giaDiezel, giaDien);

                    // Kiểm tra khác tên gọi (do Smart Mapping)
                    string tenDt = ChuanHoaTen(hpDuToan.TenHaoPhi);
                    string tenCh = ChuanHoaTen(hpChuanMatched.TenHaoPhi);
                    if (tenDt != tenCh && !tenDt.Contains(tenCh) && !tenCh.Contains(tenDt))
                    {
                        saiLech.LoaiLoi = "Khác tên gọi";
                        saiLech.MoTa = $"DT: '{hpDuToan.TenHaoPhi}' | TT38: '{hpChuanMatched.TenHaoPhi}'";
                    }

                    // Kiểm tra định mức
                    if (Math.Abs(saiLech.ChenhLechDinhMuc) > 0.0001m)
                    {
                        if (string.IsNullOrEmpty(saiLech.LoaiLoi)) saiLech.LoaiLoi = "Sai định mức";
                        else saiLech.LoaiLoi += ", Sai định mức";
                        
                        if (string.IsNullOrEmpty(saiLech.MoTa)) saiLech.MoTa = $"DT = {hpDuToan.DinhMuc:G}, TT38 = {hpChuanMatched.DinhMuc:G}";
                        else saiLech.MoTa += $"\nSai ĐM: DT = {hpDuToan.DinhMuc:G}, TT38 = {hpChuanMatched.DinhMuc:G}";
                    }

                    // Kiểm tra Đơn vị
                    if (!string.Equals(hpDuToan.DonVi, hpChuanMatched.DonVi, StringComparison.OrdinalIgnoreCase))
                    {
                        if (string.IsNullOrEmpty(saiLech.LoaiLoi)) saiLech.LoaiLoi = "Sai Đơn vị";
                        else saiLech.LoaiLoi += ", Sai Đơn vị";
                    }
                }
                
                kq.DanhSachSaiLech.Add(saiLech);
            }

            // Những hao phí chuẩn còn lại (chưa được map) là bị thiếu
            foreach (var hpThieu in dsHpChuan)
            {
                var sl = new SaiLechDinhMuc
                {
                    LoaiLoi = "Thiếu hao phí",
                    MoTa = $"TT38 có '{hpThieu.TenHaoPhi}' nhưng Dự toán bị thiếu.",
                    LoaiHP = hpThieu.LoaiHaoPhi,
                    SoDongExcel = ctDuToan.SoDongExcel, // Báo ở dòng công tác vì không có dòng hao phí
                    HaoPhiChuan = hpThieu
                };
                
                sl.DonGiaChuan = GetDonGiaChuan(hpThieu, vung, giaVlDict, giaNcDict, giaMayDict, allNcList, dmMayRepo, giaXang, giaDiezel, giaDien);
                kq.DanhSachSaiLech.Add(sl);
            }

            ketQua.Add(kq);
        }

        return ketQua;
    }

    public List<VatTuGiaModel> TrichXuatVatTu(List<KetQuaCongTacThamDinh> ketQuaDinhMuc)
    {
        var dict = new Dictionary<string, VatTuGiaModel>();

        foreach (var kq in ketQuaDinhMuc)
        {
            foreach (var saiLech in kq.DanhSachSaiLech)
            {
                if (saiLech.HaoPhiChuan == null || string.IsNullOrEmpty(saiLech.HaoPhiChuan.MaHieuHP)) continue;
                
                string maHieu = saiLech.HaoPhiChuan.MaHieuHP;
                string tenHp = (saiLech.HaoPhiChuan.TenHaoPhi ?? "").ToLower();
                string donVi = (saiLech.HaoPhiChuan.DonVi ?? "").Trim();

                // Loại bỏ các hao phí tính theo tỷ lệ % hoặc mang tính chất 'khác'
                if (saiLech.HaoPhiChuan.LoaiHaoPhi == LoaiHaoPhi.VL)
                {
                    if (donVi == "%" || tenHp.Contains("vật liệu khác") || maHieu == "VL_KHAC" || maHieu.StartsWith("VLK"))
                        continue;
                }
                else if (saiLech.HaoPhiChuan.LoaiHaoPhi == LoaiHaoPhi.MAY)
                {
                    if (donVi == "%" || tenHp.Contains("máy khác") || maHieu == "M7016")
                        continue;
                }
                else if (saiLech.HaoPhiChuan.LoaiHaoPhi == LoaiHaoPhi.NC)
                {
                    if (donVi == "%" || tenHp.Contains("nhân công khác"))
                        continue;
                }

                string key = maHieu + "_" + saiLech.HaoPhiChuan.LoaiHaoPhi.ToString();

                decimal klHaoPhi = saiLech.HaoPhiDuToan?.DinhMuc ?? 0;

                if (!dict.ContainsKey(key))
                {
                    dict[key] = new VatTuGiaModel
                    {
                        TenVatTu = saiLech.HaoPhiChuan.TenHaoPhi,
                        DonVi = saiLech.HaoPhiChuan.DonVi,
                        LoaiHP = saiLech.HaoPhiChuan.LoaiHaoPhi,
                        GiaDuToan = saiLech.HaoPhiDuToan?.DonGia ?? 0,
                        MaHieu = maHieu,
                        GiaChuan = saiLech.DonGiaChuan,
                        KhoiLuong = klHaoPhi
                    };
                }
                else
                {
                    dict[key].KhoiLuong += klHaoPhi;
                }
            }
        }

        return dict.Values.OrderBy(x => x.LoaiHP).ThenBy(x => x.TenVatTu).ToList();
    }
    
    private decimal? GetDonGiaChuan(
        HaoPhi hpChuan,
        Vung vung,
        Dictionary<string, decimal> giaVlDict,
        Dictionary<string, decimal> giaNcDict,
        Dictionary<string, decimal> giaMayDict,
        List<NhanCong> allNC,
        DinhMucCaMayRepository dmMayRepo,
        decimal gx, decimal gdz, decimal gdi)
    {
        if (string.IsNullOrEmpty(hpChuan.MaHieuHP)) return null;
        string ma = hpChuan.MaHieuHP.Trim();

        // 1. VẬT LIỆU
        if (hpChuan.LoaiHaoPhi == LoaiHaoPhi.VL)
        {
            if (giaVlDict.TryGetValue(ma, out decimal gvl) && gvl > 0)
                return gvl;
            var vl = _vlRepo.GetByMa(ma);
            if (vl != null && vl.DonGia > 0) return vl.DonGia;
            return null;
        }

        // 2. NHÂN CÔNG
        if (hpChuan.LoaiHaoPhi == LoaiHaoPhi.NC)
        {
            if (giaNcDict.TryGetValue(ma, out decimal gnc) && gnc > 0)
                return gnc;

            // Fallback: Tìm theo nhóm nhân công (N97789 -> 1, N97790 -> 2, N97791 -> 3, N97792 -> 4...)
            int nhom = 0;
            if (ma.StartsWith("N97789")) nhom = 1;
            else if (ma.StartsWith("N97790")) nhom = 2;
            else if (ma.StartsWith("N97791")) nhom = 3;
            else if (ma.StartsWith("N97792")) nhom = 4;
            else if (ma.StartsWith("N97793")) nhom = 5;
            else if (ma.StartsWith("N97794")) nhom = 6;
            else if (ma.StartsWith("N97795")) nhom = 7;

            if (nhom == 0 && !string.IsNullOrEmpty(hpChuan.TenHaoPhi))
            {
                string t = hpChuan.TenHaoPhi.ToLower();
                if (t.Contains("nhóm 1") || t.Contains("nhóm i")) nhom = 1;
                else if (t.Contains("nhóm 2") || t.Contains("nhóm ii")) nhom = 2;
                else if (t.Contains("nhóm 3") || t.Contains("nhóm iii")) nhom = 3;
                else if (t.Contains("nhóm 4") || t.Contains("nhóm iv")) nhom = 4;
                else if (t.Contains("nhóm 5") || t.Contains("nhóm v")) nhom = 5;
                else if (t.Contains("nhóm 6") || t.Contains("nhóm vi")) nhom = 6;
                else if (t.Contains("nhóm 7") || t.Contains("nhóm vii")) nhom = 7;
            }

            if (nhom > 0)
            {
                var nc = allNC.FirstOrDefault(x => x.LoaiNhanCong == LoaiNhanCong.XayDung && x.Nhom == nhom);
                if (nc != null) return nc.GetDonGia(vung);
            }

            var ncDirect = _ncRepo.GetByMa(ma);
            if (ncDirect != null) return ncDirect.GetDonGia(vung);

            return null;
        }

        // 3. MÁY THI CÔNG
        if (hpChuan.LoaiHaoPhi == LoaiHaoPhi.MAY)
        {
            if (giaMayDict.TryGetValue(ma, out decimal gm) && gm > 0)
                return gm;

            // Fallback: Tính toán giá ca máy chuẩn từ Định mức ca máy TT37
            var dmMay = dmMayRepo.GetByMaMay(ma);
            if (dmMay != null)
            {
                int soCaNam = dmMay.SoCaNam > 0 ? dmMay.SoCaNam : 250;
                decimal g_th = dmMay.NguyenGia >= 30000000m ? dmMay.NguyenGia * 0.1m : 0m;
                decimal khauHao = ((dmMay.NguyenGia - g_th) * dmMay.KhauHao / 100m) / soCaNam;
                decimal suaChua = (dmMay.NguyenGia * dmMay.SuaChua / 100m) / soCaNam;
                decimal chiPhiKhac = (dmMay.NguyenGia * dmMay.ChiPhiKhac / 100m) / soCaNam;

                decimal cpNhienLieu = (dmMay.DinhMucXang * gx + dmMay.DinhMucDiezel * gdz + dmMay.DinhMucDien * gdi) * (1m + dmMay.HeSoNhienLieuPhu / 100m);

                decimal luongTho = 0;
                var allNCForMay = _ncRepo.GetAll();
                var tpNC = dmMay.GetThanhPhanNhanCong();
                foreach (var tp in tpNC)
                {
                    var ncMay = allNCForMay.FirstOrDefault(n => n.LoaiNhanCong == LoaiNhanCong.VanHanhMay && n.Nhom == tp.Nhom);
                    if (ncMay != null) luongTho += ncMay.GetDonGia(vung) * tp.SoLuong;
                }

                decimal tongCaMay = khauHao + suaChua + chiPhiKhac + cpNhienLieu + luongTho;
                if (tongCaMay > 0) return tongCaMay;
            }

            var may = _mayRepo.GetByMa(ma);
            if (may != null && may.DonGia > 0) return may.DonGia;

            return null;
        }

        return null;
    }

    private HaoPhi? TimHaoPhiTuongDuong(HaoPhiThamDinh hpDuToan, List<HaoPhi> dsHpChuan, CongTacXayDung ctChuan)
    {
        string tenDt = ChuanHoaTen(hpDuToan.TenHaoPhi);

        // 1. Khớp chính xác hoàn toàn
        var exactMatch = dsHpChuan.FirstOrDefault(x => ChuanHoaTen(x.TenHaoPhi) == tenDt && x.LoaiHaoPhi == hpDuToan.Loai);
        if (exactMatch != null) return exactMatch;

        // 2. Khớp gần đúng (chứa chuỗi)
        var containsMatch = dsHpChuan.FirstOrDefault(x =>
            (ChuanHoaTen(x.TenHaoPhi).Contains(tenDt) || tenDt.Contains(ChuanHoaTen(x.TenHaoPhi)))
            && x.LoaiHaoPhi == hpDuToan.Loai);
            
        if (containsMatch != null) return containsMatch;

        // 3. Khớp theo mức độ trùng lặp từ vựng (Word Overlap)
        var wordsDt = tenDt.Split(new[] { ' ', '-', ',', '.', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
        HaoPhi bestMatch = null;
        int maxOverlap = 0;

        foreach (var hp in dsHpChuan.Where(x => x.LoaiHaoPhi == hpDuToan.Loai))
        {
            var wordsCh = ChuanHoaTen(hp.TenHaoPhi).Split(new[] { ' ', '-', ',', '.', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
            int overlap = wordsDt.Intersect(wordsCh).Count();
            
            if (overlap > maxOverlap)
            {
                maxOverlap = overlap;
                bestMatch = hp;
            }
        }

        // Bắt cặp nếu trùng ít nhất 2 từ (VD: "Máy đào", "Máy đầm", "Vật liệu")
        // Nếu tên quá ngắn (chỉ 1 từ) thì cần trùng 1 từ
        int minRequiredOverlap = wordsDt.Length <= 1 ? 1 : 2;
        if (bestMatch != null && maxOverlap >= minRequiredOverlap)
        {
            return bestMatch;
        }

        // 4. Smart Mapping: Nếu công tác chỉ có 1 hao phí loại này, và chuẩn cũng có 1 -> Bắt cặp
        // Nhưng BẮT BUỘC phải có ít nhất 1 từ trùng để tránh ghép nhầm (VD: "Nhựa bitum" với "Bột đá")
        var countChuanLoaiNay = ctChuan.DanhSachHaoPhi.Count(x => x.LoaiHaoPhi == hpDuToan.Loai);
        if (countChuanLoaiNay == 1)
        {
            var smartMatch = dsHpChuan.FirstOrDefault(x => x.LoaiHaoPhi == hpDuToan.Loai);
            if (smartMatch != null)
            {
                // Kiểm tra có ít nhất 1 từ trùng
                var w1 = tenDt.Split(new[] { ' ', '-', ',', '.', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                var w2 = ChuanHoaTen(smartMatch.TenHaoPhi).Split(new[] { ' ', '-', ',', '.', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                if (w1.Intersect(w2).Any()) return smartMatch;
            }
        }
        
        // 5. Fallback cuối cùng: Nếu chỉ còn 1 hao phí TT38 chưa map cùng loại, bắt cặp NẾU có từ trùng.
        // Tuyệt đối KHÔNG bắt cặp khi tên hoàn toàn khác nhau → phải để thành "Hao phí thừa".
        var remainingChuanLoaiNay = dsHpChuan.Where(x => x.LoaiHaoPhi == hpDuToan.Loai).ToList();
        if (remainingChuanLoaiNay.Count == 1)
        {
            var candidate = remainingChuanLoaiNay.First();
            var w1 = tenDt.Split(new[] { ' ', '-', ',', '.', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
            var w2 = ChuanHoaTen(candidate.TenHaoPhi).Split(new[] { ' ', '-', ',', '.', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
            if (w1.Intersect(w2).Any()) return candidate;
        }

        return null;
    }

    private string ChuanHoaTen(string input)
    {
        if (string.IsNullOrWhiteSpace(input)) return string.Empty;
        var result = input.Trim().ToLower();
        if (result.StartsWith("-")) result = result.Substring(1).Trim();
        // Loại bỏ khoảng trắng thừa
        result = Regex.Replace(result, @"\s+", " ");
        return result;
    }
}
