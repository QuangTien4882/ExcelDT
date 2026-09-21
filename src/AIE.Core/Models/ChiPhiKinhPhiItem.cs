using System;
using System.Collections.Generic;
using System.Linq;

namespace AIE.Core.Models
{
    /// <summary>
    /// Nhóm chi phí trong Tổng mức đầu tư và Dự toán xây dựng công trình
    /// (Theo Thông tư 36/2026/TT-BXD)
    /// </summary>
    public enum NhomChiPhi
    {
        BoiThuong_TDC,    // G_BT,TĐC (Chi phí bồi thường, hỗ trợ và tái định cư)
        ChiPhiXayDung,    // G_XD (Chi phí xây dựng)
        ChiPhiThietBi,    // G_TB (Chi phí thiết bị)
        QuanLyDuAn,       // G_QLDA (Chi phí quản lý dự án)
        TuVanDauTuXD,     // G_TV (Chi phí tư vấn đầu tư xây dựng)
        ChiPhiKhac,       // G_K (Chi phí khác)
        ChiPhiDuPhong     // G_DP (Chi phí dự phòng)
    }

    /// <summary>
    /// Phương pháp xác định giá trị chi phí
    /// </summary>
    public enum CachTinhChiPhi
    {
        TheoTyLeDinhMuc,  // Tỷ lệ % × Cơ sở tính × Hệ số k
        NhapTruocThue,    // Nhập trực tiếp giá trị tiền trước thuế
        TheoDuToanRieng   // Lập dự toán riêng
    }

    /// <summary>
    /// Cơ sở tính chi phí theo định mức tỷ lệ %
    /// </summary>
    public enum CoSoTinhChiPhi
    {
        ChiPhiXayDung,             // G_XD (trước thuế)
        ChiPhiThietBi,             // G_TB (trước thuế)
        ChiPhiXD_Va_ThietBi,       // G_XD + G_TB (trước thuế)
        TongChiPhiTruocDuPhong,    // G_XD + G_TB + G_QLDA + G_TV + G_K (hoặc các nhóm trước DP)
        TongMucDauTu,              // Toàn bộ Tổng mức đầu tư
        ChiPhiBoiThuong,           // G_BT,TĐC
        TungKhoanMucRieng          // Tùy biến
    }

    /// <summary>
    /// Đại diện cho một khoản mục chi phí cụ thể trong Bảng Tổng mức đầu tư hoặc Bảng Tổng hợp dự toán
    /// </summary>
    public class ChiPhiKinhPhiItem
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();

        /// <summary>Số thứ tự hiển thị (VD: "I", "4.1", "5.3")</summary>
        public string STT { get; set; } = string.Empty;

        /// <summary>Mã định danh khoản mục (VD: "G_XD", "TV_01", "K_02")</summary>
        public string MaChiPhi { get; set; } = string.Empty;

        /// <summary>Tên khoản mục chi phí</summary>
        public string TenChiPhi { get; set; } = string.Empty;

        /// <summary>Phân nhóm chi phí cha</summary>
        public NhomChiPhi Nhom { get; set; }

        /// <summary>Cách tính chi phí</summary>
        public CachTinhChiPhi CachTinh { get; set; } = CachTinhChiPhi.TheoTyLeDinhMuc;

        /// <summary>Cơ sở tính chi phí</summary>
        public CoSoTinhChiPhi CoSoTinh { get; set; } = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi;

        /// <summary>Tỷ lệ định mức % (VD: 3.25%)</summary>
        public decimal TyLePhanTram { get; set; }

        /// <summary>Hệ số điều chỉnh k (VD: 1.1, 0.8, 1.35...)</summary>
        public decimal HeSoDieuChinh { get; set; } = 1.0m;

        /// <summary>Giá trị trước thuế (đồng)</summary>
        public decimal GiaTriTruocThue { get; set; }

        /// <summary>Thuế suất GTGT (VD: 0.10 cho 10%, 0.08 cho 8%, 0 cho Phí/Lệ phí)</summary>
        public decimal ThueSuatGTGT { get; set; } = 0.10m;

        /// <summary>Tiền thuế GTGT (đồng) = GiaTriTruocThue * ThueSuatGTGT</summary>
        public decimal TienThueGTGT => Math.Round(GiaTriTruocThue * ThueSuatGTGT, 0, MidpointRounding.AwayFromZero);

        /// <summary>Giá trị sau thuế (đồng) = GiaTriTruocThue + TienThueGTGT</summary>
        public decimal GiaTriSauThue => GiaTriTruocThue + TienThueGTGT;

        /// <summary>Ký hiệu theo quy chuẩn TT 36 (VD: Gxd, Gtb, Gqlda, Gtv1, Gk1, Gdp...)</summary>
        public string KyHieu { get; set; } = string.Empty;

        /// <summary>Diễn giải cách tính hiển thị trên bảng</summary>
        public string GhiChuCachTinh { get; set; } = string.Empty;

        /// <summary>Mức khống chế tối thiểu (đồng, null nếu không có)</summary>
        public decimal? MinValue { get; set; }

        /// <summary>Mức khống chế tối đa (đồng, null nếu không có)</summary>
        public decimal? MaxValue { get; set; }

        /// <summary>Bật/Tắt khoản mục chi phí (nếu false thì không cộng vào tổng)</summary>
        public bool IsActive { get; set; } = true;

        /// <summary>Khoản mục do người dùng tự thêm mới vào</summary>
        public bool IsUserAdded { get; set; } = false;

        /// <summary>Khoản mục cốt lõi (G_XD, G_TB...) không cho phép xóa</summary>
        public bool IsReadOnly { get; set; } = false;

        /// <summary>Nhân bản item</summary>
        public ChiPhiKinhPhiItem Clone()
        {
            return (ChiPhiKinhPhiItem)this.MemberwiseClone();
        }
    }

    /// <summary>
    /// Model tổng hợp toàn bộ các khoản mục chi phí của Bảng TMĐT và Tổng hợp dự toán
    /// </summary>
    public class BangTongHopKinhPhiModel
    {
        // --- Thông số thiết lập dự án ---
        public string LoaiCongTrinh { get; set; } = "";         // Dân dụng, Công nghiệp, Giao thông, Nông nghiệp & Môi trường, Hạ tầng kỹ thuật
        public string CapCongTrinh { get; set; } = "";          // Cấp đặc biệt, Cấp I, Cấp II, Cấp III, Cấp IV
        public int SoBuocThietKe { get; set; } = 0;             // 0: Chưa chọn, 1: 1 bước, 2: 2 bước, 3: 3 bước

        // --- Các điều kiện áp dụng hệ số điều chỉnh ---
        public bool ThietBiTren50Pct { get; set; } = false;      // Thiết bị >= 50% (k = 0.7 cho kiểm toán, thẩm tra quyết toán)
        public bool DaKiemToanDocLap { get; set; } = false;      // Dự án đã kiểm toán (k = 0.5 cho thẩm tra quyết toán)
        public bool YeuCauThueThamTra { get; set; } = false;     // Cơ quan thẩm định yêu cầu thuê thẩm tra (k = 0.5 cho phí thẩm định)
        public bool CdtTuQuanLy { get; set; } = false;           // CĐT tự QLDA (k = 0.8 cho QLDA)
        public bool VungKhoKhan { get; set; } = false;           // Vùng sâu xa hải đảo (k = 1.35)
        public bool TuyenQuaNhieuTinh { get; set; } = false;     // Tuyến dài qua nhiều tỉnh (k = 1.1)
        public bool CaiTaoSuaChua { get; set; } = false;         // Cải tạo sửa chữa (k = 1.1 cho thiết kế)
        public bool ThietKeLapLai { get; set; } = false;         // Thiết kế lặp lại (k = 0.36)

        // --- Quy mô chi phí đầu vào ---
        public decimal ChiPhiXDTruocThue { get; set; }           // G_XD (đồng) - Chi phí xây dựng chưa bao gồm nhà tạm
        public decimal ChiPhiNhaTamTruocThue { get; set; }       // G_NT (đồng) - Chi phí nhà tạm để ở và điều hành thi công
        public decimal ChiPhiTBTruocThue { get; set; }           // G_TB (đồng)
        public decimal ChiPhiBTTruocThue { get; set; }           // G_BT,TĐC (đồng)

        // --- Danh sách các khoản mục chi phí ---
        public List<ChiPhiKinhPhiItem> Items { get; set; } = new List<ChiPhiKinhPhiItem>();

        /// <summary>
        /// Tính toán lại giá trị của tất cả các dòng chi phí theo cơ sở tính
        /// </summary>
        public void TinhToanLai()
        {
            // 1. Cập nhật các dòng chính: G_XD, G_NHA_TAM, G_TB, G_BT
            var itemBT = Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.BoiThuong_TDC);
            if (itemBT != null) itemBT.GiaTriTruocThue = ChiPhiBTTruocThue;

            var itemXD = Items.FirstOrDefault(x => x.MaChiPhi == "G_XD");
            if (itemXD != null) itemXD.GiaTriTruocThue = ChiPhiXDTruocThue;

            var itemNT = Items.FirstOrDefault(x => x.MaChiPhi == "G_NHA_TAM");
            if (itemNT != null) itemNT.GiaTriTruocThue = ChiPhiNhaTamTruocThue;

            var itemTB = Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiThietBi);
            if (itemTB != null) itemTB.GiaTriTruocThue = ChiPhiTBTruocThue;

            // Cơ sở tính theo tỷ lệ định mức (QLDA, Giám sát, Thiết kế, Thẩm tra...):
            // CHÚ Ý: Theo quy định, Chi phí xây dựng làm cơ sở tính KHÔNG bao gồm chi phí nhà tạm!
            decimal gXD = ChiPhiXDTruocThue;
            decimal gTB = ChiPhiTBTruocThue;
            decimal gXDTB = gXD + gTB;

            // 2. Tính QLDA và Tư vấn (các khoản có cơ sở tính XD / TB / XD+TB)
            foreach (var item in Items)
            {
                if (!item.IsActive || item.Nhom == NhomChiPhi.ChiPhiXayDung || item.Nhom == NhomChiPhi.ChiPhiThietBi || item.Nhom == NhomChiPhi.BoiThuong_TDC || item.Nhom == NhomChiPhi.ChiPhiDuPhong)
                    continue;

                // Khoản tính theo Tổng mức đầu tư (phí thẩm định dự án) được tính riêng ở bước 3
                if (item.CoSoTinh == CoSoTinhChiPhi.TongMucDauTu)
                    continue;

                if (item.CachTinh == CachTinhChiPhi.TheoTyLeDinhMuc)
                {
                    decimal coSo = 0;
                    switch (item.CoSoTinh)
                    {
                        case CoSoTinhChiPhi.ChiPhiXayDung:
                            coSo = gXD;
                            break;
                        case CoSoTinhChiPhi.ChiPhiThietBi:
                            coSo = gTB;
                            break;
                        case CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi:
                        default:
                            coSo = gXDTB;
                            break;
                    }

                    decimal val = (coSo * (item.TyLePhanTram / 100m)) * item.HeSoDieuChinh;

                    // Áp dụng giới hạn Min/Max nếu có
                    if (item.MinValue.HasValue && val < item.MinValue.Value && coSo > 0)
                        val = item.MinValue.Value;
                    if (item.MaxValue.HasValue && val > item.MaxValue.Value)
                        val = item.MaxValue.Value;

                    item.GiaTriTruocThue = Math.Round(val, 0, MidpointRounding.AwayFromZero);
                }
            }

            // 3. Tính các khoản có cơ sở tính theo Tổng mức đầu tư (Ví dụ: Phí thẩm định dự án - K_TD_DA)
            // Cơ sở = tổng các khoản mục đang kích hoạt trừ chính khoản phí đang tính và trừ dự phòng
            foreach (var item in Items)
            {
                if (!item.IsActive || item.CachTinh != CachTinhChiPhi.TheoTyLeDinhMuc)
                    continue;
                if (item.CoSoTinh != CoSoTinhChiPhi.TongMucDauTu)
                    continue;

                decimal coSo = Items
                    .Where(x => x.IsActive
                        && x.Id != item.Id
                        && x.Nhom != NhomChiPhi.ChiPhiDuPhong)
                    .Sum(x => x.GiaTriTruocThue);

                decimal val = (coSo * (item.TyLePhanTram / 100m)) * item.HeSoDieuChinh;

                if (item.MinValue.HasValue && val < item.MinValue.Value && coSo > 0)
                    val = item.MinValue.Value;
                if (item.MaxValue.HasValue && val > item.MaxValue.Value)
                    val = item.MaxValue.Value;

                item.GiaTriTruocThue = Math.Round(val, 0, MidpointRounding.AwayFromZero);
            }

            // 4. Tính Chi phí Dự phòng (G_DP)
            // Tổng chi phí các nhóm từ I đến V (trước DP)
            decimal tongTruocDP = Items
                .Where(x => x.IsActive && x.Nhom != NhomChiPhi.ChiPhiDuPhong)
                .Sum(x => x.GiaTriTruocThue);

            var itemDP = Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiDuPhong);
            if (itemDP != null && itemDP.IsActive)
            {
                if (itemDP.CachTinh == CachTinhChiPhi.TheoTyLeDinhMuc)
                {
                    decimal valDP = tongTruocDP * (itemDP.TyLePhanTram / 100m) * itemDP.HeSoDieuChinh;
                    itemDP.GiaTriTruocThue = Math.Round(valDP, 0, MidpointRounding.AwayFromZero);
                }
            }
        }

        /// <summary>Tổng chi phí trước thuế của tất cả các khoản mục đang kích hoạt</summary>
        public decimal TongTruocThue => Items.Where(x => x.IsActive).Sum(x => x.GiaTriTruocThue);

        /// <summary>Tổng tiền thuế GTGT</summary>
        public decimal TongTienThueGTGT => Items.Where(x => x.IsActive).Sum(x => x.TienThueGTGT);

        /// <summary>Tổng chi phí sau thuế của từng nhóm (đã làm tròn đến hàng nghìn đồng theo quy định chuẩn)</summary>
        public decimal GetTongSauThueNhom(NhomChiPhi nhom)
        {
            var sum = Items.Where(x => x.IsActive && x.Nhom == nhom).Sum(x => x.GiaTriSauThue);
            return Math.Round(sum / 1000m, 0, MidpointRounding.AwayFromZero) * 1000m;
        }

        /// <summary>
        /// Tổng chi phí sau thuế (Tổng mức đầu tư hoặc Tổng hợp dự toán) - làm tròn đến hàng nghìn đồng theo quy định.
        /// Tính bằng tổng các nhóm chi phí (mỗi nhóm đã làm tròn đến hàng nghìn) và làm tròn tổng thể đến hàng nghìn đồng như công thức Excel.
        /// </summary>
        public decimal TongSauThue
        {
            get
            {
                var nhoms = Items.Where(x => x.IsActive).Select(x => x.Nhom).Distinct();
                decimal tong = 0m;
                foreach (var n in nhoms)
                {
                    tong += GetTongSauThueNhom(n);
                }
                return Math.Round(tong / 1000m, 0, MidpointRounding.AwayFromZero) * 1000m;
            }
        }

        /// <summary>Tổng chi phí sau thuế cộng gộp chưa làm tròn</summary>
        public decimal TongSauThueChuaLamTron => Items.Where(x => x.IsActive).Sum(x => x.GiaTriSauThue);
    }
}
