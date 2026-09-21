using AIE.Data.ImportExport;
using AIE.Data.Repositories;
using System;
using System.IO;

namespace AIE.Data
{
    /// <summary>
    /// Quản lý tập trung database SQLite: khởi tạo, import dữ liệu.
    /// </summary>
    public class DatabaseManager : IDisposable
    {
        private readonly string _dbPath;
        private AieDbContext _context;

        public DatabaseManager()
        {
            // Lưu DB tại thư mục AppData của người dùng
            var appDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AIE_DuToan");
            if (!Directory.Exists(appDataFolder))
                Directory.CreateDirectory(appDataFolder);

            _dbPath = Path.Combine(appDataFolder, "aie_database.sqlite");
            _context = new AieDbContext(_dbPath);
            _context.InitializeDatabase();
        }

        public string DbPath { get { return _dbPath; } }
        public AieDbContext Context { get { return _context; } }

        public void Dispose()
        {
            _context?.Dispose();
            _context = null;
        }

        /// <summary>
        /// Import file 1_DinhMucCongTac.xlsx
        /// </summary>
        public ImportResult ImportDinhMucCongTac(string filePath)
        {
            var importService = new ExcelImportService();
            ImportResult result;
            var danhSachCT = importService.ImportDinhMucCongTac(filePath, out result);

            var repo = new CongTacRepository(_context);
            int newCount = 0;
            int updateCount = 0;
            foreach (var ct in danhSachCT)
            {
                if (repo.Insert(ct))
                    newCount++;
                else
                    updateCount++;
            }
            
            result.SoLuongMoi += newCount;
            result.SoLuongCapNhat += updateCount;

            return result;
        }

        /// <summary>
        /// Import trực tiếp từ file F1 xuất ra
        /// </summary>
        public ImportResult ImportF1DinhMucCongTac(string filePath)
        {
            var importer = new F1DatabaseImporter();
            var danhSachCT = importer.ParseF1File(filePath);

            var repo = new CongTacRepository(_context);
            int countCT = 0;
            int countHP = 0;
            int newCount = 0;
            int updateCount = 0;
            
            foreach (var ct in danhSachCT)
            {
                if (repo.Insert(ct))
                    newCount++;
                else
                    updateCount++;
                    
                countCT++;
                countHP += ct.DanhSachHaoPhi.Count;
            }

            var result = new ImportResult();
            result.SoLuongThanhCong = countCT + countHP;
            result.SoLuongMoi = newCount;
            result.SoLuongCapNhat = updateCount;
            return result;
        }

        /// <summary>
        /// Import file 2_GiaVatLieu.xlsx
        /// </summary>
        public ImportResult ImportGiaVatLieu(string filePath)
        {
            var importService = new ExcelImportService();
            ImportResult result;
            var data = importService.ImportGiaVatLieu(filePath, out result);

            var repo = new VatLieuRepository(_context);
            foreach (var item in data)
            {
                repo.Upsert(item);
            }

            return result;
        }

        /// <summary>
        /// Import file 3_GiaNhanCong.xlsx
        /// </summary>
        public ImportResult ImportGiaNhanCong(string filePath)
        {
            var importService = new ExcelImportService();
            ImportResult result;
            var data = importService.ImportGiaNhanCong(filePath, out result);

            var repo = new NhanCongRepository(_context);
            foreach (var item in data)
            {
                repo.Upsert(item);
            }

            return result;
        }

        /// <summary>
        /// Import file 4_GiaMayThiCong.xlsx
        /// </summary>
        public ImportResult ImportGiaMayThiCong(string filePath)
        {
            var importService = new ExcelImportService();
            ImportResult result;
            var data = importService.ImportGiaMayThiCong(filePath, out result);

            var repo = new MayThiCongRepository(_context);
            foreach (var item in data)
            {
                repo.Upsert(item);
            }

            return result;
        }

        public void ClearAllDinhMuc()
        {
            using (var connection = _context.GetConnection())
            {
                var sql = @"
                    DELETE FROM HaoPhi;
                    DELETE FROM CongTacXayDung;
                    UPDATE sqlite_sequence SET seq = 0 WHERE name = 'HaoPhi' OR name = 'CongTacXayDung';
                ";
                Dapper.SqlMapper.Execute(connection, sql);
            }
        }

        public void ClearAllVatLieu()
        {
            using (var connection = _context.GetConnection())
            {
                var sql = @"
                    DELETE FROM VatLieu;
                    UPDATE sqlite_sequence SET seq = 0 WHERE name = 'VatLieu';
                ";
                Dapper.SqlMapper.Execute(connection, sql);
            }
        }

        public void ClearAllNhanCong()
        {
            using (var connection = _context.GetConnection())
            {
                var sql = @"
                    DELETE FROM NhanCong;
                    UPDATE sqlite_sequence SET seq = 0 WHERE name = 'NhanCong';
                ";
                Dapper.SqlMapper.Execute(connection, sql);
            }
        }

        public void ClearAllMayThiCong()
        {
            using (var connection = _context.GetConnection())
            {
                var sql = @"
                    DELETE FROM MayThiCong;
                    UPDATE sqlite_sequence SET seq = 0 WHERE name = 'MayThiCong';
                ";
                Dapper.SqlMapper.Execute(connection, sql);
            }
        }

        /// <summary>
        /// Lấy thống kê số lượng dữ liệu trong DB.
        /// </summary>
        public DatabaseStats GetStats()
        {
            var stats = new DatabaseStats();
            var vlRepo = new VatLieuRepository(_context);
            var ncRepo = new NhanCongRepository(_context);
            var mayRepo = new MayThiCongRepository(_context);
            var ctRepo = new CongTacRepository(_context);

            stats.SoCongTac = ctRepo.Count();
            stats.SoVatLieu = vlRepo.GetAll().ToList().Count;
            stats.SoNhanCong = ncRepo.GetAll().ToList().Count;
            stats.SoMayThiCong = mayRepo.GetAll().ToList().Count;

            return stats;
        }

        /// <summary>
        /// Seed dữ liệu Nhân công theo TT37/2026 cho Đà Nẵng.
        /// Bao gồm: Xây dựng (nhóm 1-4), Vận hành máy (nhóm 1-6), Khác (nhóm 1-3).
        /// Giá theo 4 vùng: II, III, IV, Cù Lao Chàm.
        /// </summary>
        public void SeedNhanCong()
        {
            var ncRepo = new NhanCongRepository(_context);
            var existing = ncRepo.GetAll().ToList();

            // Dữ liệu Nhân công Xây dựng (Nhóm 1-4) - Giá theo vùng (đồng/công)
            var xdData = new[]
            {
                // MaNC, TenNC, Nhom, DonGiaVung2, DonGiaVung3, DonGiaVung4, DonGiaCLC
                ("NC_XD_1", "Nhân công xây dựng nhóm I",   1, 254498m, 250383m, 242061m, 306897m),
                ("NC_XD_2", "Nhân công xây dựng nhóm II",  2, 294557m, 290753m, 283826m, 356204m),
                ("NC_XD_3", "Nhân công xây dựng nhóm III", 3, 315144m, 311218m, 304215m, 381250m),
                ("NC_XD_4", "Nhân công xây dựng nhóm IV",  4, 342049m, 339770m, 329801m, 412055m),
            };

            // Dữ liệu Nhân công Vận hành máy (Nhóm 1-6) - Giá theo vùng
            var vhmData = new[]
            {
                ("NC_VHM_1", "Nhân công vận hành máy, điều khiển máy",   1, 336697m, 332892m, 324086m, 405515m),
                ("NC_VHM_2", "Lái xe",  2, 321698m, 318062m, 309649m, 387450m),
                ("NC_VHM_3", "Thủy thủ, thợ máy, thợ điện", 3, 336200m, 325000m, 310100m, 400400m),
                ("NC_VHM_4", "Máy trưởng, máy I, máy II, điện trưởng, kỹ thuật viên cuốc I, kỹ thuật viên cuốc II tàu biển",  4, 360400m, 338300m, 329700m, 410200m),
                ("NC_VHM_5", "Máy trưởng, máy I, máy II, điện trưởng, kỹ thuật viên cuốc I, kỹ thuật viên cuốc II tàu sông",   5, 392500m, 359500m, 0m, 471000m),
                ("NC_VHM_6", "Thuyền trưởng, thuyền phó",  6, 416700m, 405100m, 394500m, 502400m),
            };

            // Dữ liệu Nhân công Khác (Nhóm 1-3)
            var khacData = new[]
            {
                ("NC_K_1", "Kỹ sư thực hiện khảo sát, thí nghiệm",   1, 313400m, 304600m, 295200m, 353300m),
                ("NC_K_2", "Thợ lặn",  2, 580900m, 557400m, 529900m, 641700m),
                ("NC_K_3", "Nghệ nhân", 3, 578300m, 549500m, 522000m, 602500m),
            };

            void UpsertNC(string ma, string ten, int nhom, AIE.Core.Enums.LoaiNhanCong loai, decimal v2, decimal v3, decimal v4, decimal clc)
            {
                var nc = existing.FirstOrDefault(x => x.MaNC == ma);
                if (nc == null)
                {
                    nc = new AIE.Core.Models.NhanCong();
                }
                nc.MaNC = ma;
                nc.TenNC = ten;
                nc.Nhom = nhom;
                nc.LoaiNhanCong = loai;
                nc.DonVi = "công";
                nc.DonGiaVung2 = v2;
                nc.DonGiaVung3 = v3;
                nc.DonGiaVung4 = v4;
                nc.DonGiaCLC = clc;
                nc.NgayCapNhat = System.DateTime.Now;
                ncRepo.Upsert(nc);
            }

            foreach (var (ma, ten, nhom, v2, v3, v4, clc) in xdData)
                UpsertNC(ma, ten, nhom, AIE.Core.Enums.LoaiNhanCong.XayDung, v2, v3, v4, clc);

            foreach (var (ma, ten, nhom, v2, v3, v4, clc) in vhmData)
                UpsertNC(ma, ten, nhom, AIE.Core.Enums.LoaiNhanCong.VanHanhMay, v2, v3, v4, clc);

            foreach (var (ma, ten, nhom, v2, v3, v4, clc) in khacData)
                UpsertNC(ma, ten, nhom, AIE.Core.Enums.LoaiNhanCong.Khac, v2, v3, v4, clc);
        }
    }

    public class DatabaseStats
    {
        public int SoCongTac { get; set; }
        public int SoVatLieu { get; set; }
        public int SoNhanCong { get; set; }
        public int SoMayThiCong { get; set; }

        public string TomTat
        {
            get
            {
                return string.Format(
                    "Công tác: {0}\nVật liệu: {1}\nNhân công: {2}\nMáy thi công: {3}",
                    SoCongTac, SoVatLieu, SoNhanCong, SoMayThiCong);
            }
        }
    }
}
