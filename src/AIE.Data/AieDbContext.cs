using System;
using System.Data;
using System.Data.SQLite;
using System.IO;

namespace AIE.Data
{
    public class AieDbContext : IDisposable
    {
        /// <summary>Thời gian chờ tối đa (ms) khi CSDL đang bị thao tác khác giữ khóa.</summary>
        private const int BusyTimeoutMs = 5000;

        private readonly string _connectionString;

        public AieDbContext(string dbPath)
        {
            // Pooling: tái sử dụng kết nối; BusyTimeout: chờ thay vì báo lỗi "database is locked".
            _connectionString = $"Data Source={dbPath};Version=3;Pooling=True;BusyTimeout={BusyTimeoutMs};";
        }

        public IDbConnection GetConnection()
        {
            var conn = new SQLiteConnection(_connectionString);
            conn.Open();

            // Áp PRAGMA cho từng kết nối (an toàn kể cả khi driver không hỗ trợ BusyTimeout trong chuỗi kết nối).
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"PRAGMA busy_timeout={BusyTimeoutMs}; PRAGMA foreign_keys=ON;";
                cmd.ExecuteNonQuery();
            }

            return conn;
        }

        /// <summary>
        /// AieDbContext không giữ kết nối thường trực (mỗi thao tác mở/đóng riêng), nên không có tài nguyên cần giải phóng.
        /// Triển khai IDisposable để sẵn sàng cho các thay đổi sau này và để dùng được với using.
        /// </summary>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        public void InitializeDatabase()
        {
            using var connection = GetConnection();
            using var command = connection.CreateCommand();

            // Bảng định mức công tác
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS CongTacXayDung (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    MaHieu TEXT NOT NULL UNIQUE,
                    TenCongTac TEXT NOT NULL,
                    DonVi TEXT NOT NULL
                );
            ";
            command.ExecuteNonQuery();

            // Bảng định mức hao phí (VL, NC, Máy)
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS HaoPhi (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    CongTacId INTEGER NOT NULL,
                    LoaiHaoPhi INTEGER NOT NULL,
                    MaHieuHP TEXT NOT NULL,
                    TenHaoPhi TEXT NOT NULL,
                    DonVi TEXT NOT NULL,
                    DinhMuc REAL NOT NULL,
                    HeSo REAL DEFAULT 1.0,
                    FOREIGN KEY(CongTacId) REFERENCES CongTacXayDung(Id)
                );
            ";
            command.ExecuteNonQuery();

            // Bảng giá Vật Liệu
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS VatLieu (
                    MaVL TEXT PRIMARY KEY,
                    TenVL TEXT NOT NULL,
                    DonVi TEXT NOT NULL,
                    DonGia REAL NOT NULL,
                    CuocVanChuyen REAL DEFAULT 0,
                    NhaSanXuat TEXT,
                    GhiChu TEXT,
                    NgayCapNhat TEXT NOT NULL
                );
            ";
            command.ExecuteNonQuery();

            // Cập nhật cấu trúc nếu đã tồn tại
            try
            {
                command.CommandText = "ALTER TABLE VatLieu ADD COLUMN CuocVanChuyen REAL DEFAULT 0;";
                command.ExecuteNonQuery();
            }
            catch { /* Bỏ qua nếu cột đã tồn tại */ }

            // Bảng giá Nhân Công
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS NhanCong (
                    MaNC TEXT PRIMARY KEY,
                    TenNC TEXT NOT NULL,
                    Nhom INTEGER NOT NULL,
                    LoaiNhanCong INTEGER DEFAULT 1,
                    DonVi TEXT NOT NULL,
                    DonGiaVung2 REAL DEFAULT 0,
                    DonGiaVung3 REAL DEFAULT 0,
                    DonGiaVung4 REAL DEFAULT 0,
                    DonGiaCLC REAL DEFAULT 0,
                    GhiChu TEXT,
                    NgayCapNhat TEXT NOT NULL
                );
            ";
            command.ExecuteNonQuery();

            try { command.CommandText = "ALTER TABLE NhanCong ADD COLUMN LoaiNhanCong INTEGER DEFAULT 1;"; command.ExecuteNonQuery(); } catch { }
            try { command.CommandText = "ALTER TABLE NhanCong ADD COLUMN DonGiaVung2 REAL DEFAULT 0;"; command.ExecuteNonQuery(); } catch { }
            try { command.CommandText = "ALTER TABLE NhanCong ADD COLUMN DonGiaVung3 REAL DEFAULT 0;"; command.ExecuteNonQuery(); } catch { }
            try { command.CommandText = "ALTER TABLE NhanCong ADD COLUMN DonGiaVung4 REAL DEFAULT 0;"; command.ExecuteNonQuery(); } catch { }
            try { command.CommandText = "ALTER TABLE NhanCong ADD COLUMN DonGiaCLC REAL DEFAULT 0;"; command.ExecuteNonQuery(); } catch { }

            // Bảng giá Máy Thi Công
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS MayThiCong (
                    MaMay TEXT PRIMARY KEY,
                    TenMay TEXT NOT NULL,
                    DonVi TEXT NOT NULL,
                    DonGia REAL NOT NULL,
                    GhiChu TEXT,
                    NgayCapNhat TEXT NOT NULL
                );
            ";
            command.ExecuteNonQuery();

            // Bảng cấu hình Tỉ lệ phần trăm
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS TiLePhanTram (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    LoaiCongTrinh TEXT NOT NULL,
                    CapCongTrinh TEXT,
                    LoaiTiLe TEXT NOT NULL,
                    GiaTri REAL NOT NULL,
                    CoSoTinh TEXT NOT NULL
                );
            ";
            command.ExecuteNonQuery();

            // Bảng định mức Tỉ lệ Tư vấn
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS DinhMucTuVan (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    LoaiTuVan TEXT NOT NULL,
                    LoaiCongTrinh TEXT NOT NULL,
                    QuyMoMin REAL NOT NULL,
                    QuyMoMax REAL,
                    TiLe REAL NOT NULL,
                    GhiChu TEXT
                );
            ";
            command.ExecuteNonQuery();

            // Bảng Từ điển đồng nghĩa
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS TuDienDongNghia (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    TenGoc TEXT NOT NULL,
                    TenChuan TEXT NOT NULL,
                    LoaiVatTu TEXT NOT NULL, -- VL, NC, MAY
                    NgayTao TEXT NOT NULL
                );
            ";
            command.ExecuteNonQuery();

            // --- CÁC BẢNG CHO MODULE BỘ ĐƠN GIÁ VÀ ĐỊNH MỨC CA MÁY TT37 ---

            // Bảng quản lý Bộ đơn giá (chỉ dành cho VL và Máy)
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS BoDonGia (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    TenBo TEXT NOT NULL,         
                    GiaXang REAL,                
                    GiaDiezel REAL,
                    GiaDien REAL,
                    GhiChu TEXT,
                    NgayTao TEXT
                );
            ";
            command.ExecuteNonQuery();

            // Bảng Tính Giá Hiện Trường Vật Liệu (Hỗ trợ nhiều nguồn)
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS GiaVatLieuTheoBo (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    BoDonGiaId INTEGER,
                    MaVL TEXT,
                    NguonCungCap TEXT,
                    GiaGoc REAL,
                    CuLy_Km REAL,
                    LoaiDuong TEXT,
                    CuocVC REAL,
                    GiaHienTruong REAL,
                    DuocChon INTEGER,
                    ChiPhiBocXep REAL DEFAULT 0,
                    CuocVCOTo REAL DEFAULT 0,
                    CuocVCBo REAL DEFAULT 0,
                    FOREIGN KEY(BoDonGiaId) REFERENCES BoDonGia(Id)
                );
            ";
            command.ExecuteNonQuery();

            // Migration cho GiaVatLieuTheoBo (nếu bảng đã tồn tại từ trước)
            try { command.CommandText = "ALTER TABLE GiaVatLieuTheoBo ADD COLUMN ChiPhiBocXep REAL DEFAULT 0;"; command.ExecuteNonQuery(); } catch { }
            try { command.CommandText = "ALTER TABLE GiaVatLieuTheoBo ADD COLUMN CuocVCOTo REAL DEFAULT 0;"; command.ExecuteNonQuery(); } catch { }
            try { command.CommandText = "ALTER TABLE GiaVatLieuTheoBo ADD COLUMN CuocVCBo REAL DEFAULT 0;"; command.ExecuteNonQuery(); } catch { }

            // Bảng Định mức Ca máy theo TT 37
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS DinhMucCaMay_TT37 (
                    MaMay TEXT PRIMARY KEY,
                    NguyenGia REAL,
                    KhauHao REAL,
                    SuaChua REAL,
                    ChiPhiKhac REAL,
                    DinhMucXang REAL,
                    DinhMucDiezel REAL,
                    DinhMucDien REAL,
                    SoLuongNhanCong REAL,
                    NhomNhanCong INTEGER,
                    SoCaNam INTEGER DEFAULT 250,
                    NhanCongString TEXT,
                    ThanhPhanNhanCong TEXT
                );
            ";
            command.ExecuteNonQuery();

            // Migration cho DinhMucCaMay_TT37 (nếu bảng đã tồn tại từ trước)
            try { command.CommandText = "ALTER TABLE DinhMucCaMay_TT37 ADD COLUMN NhanCongString TEXT;"; command.ExecuteNonQuery(); } catch { }
            try { command.CommandText = "ALTER TABLE DinhMucCaMay_TT37 ADD COLUMN ThanhPhanNhanCong TEXT;"; command.ExecuteNonQuery(); } catch { }

            // Bảng Kết quả Giá Máy
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS GiaMayTheoBo (
                    BoDonGiaId INTEGER,
                    MaMay TEXT,
                    DonGia REAL,
                    PRIMARY KEY(BoDonGiaId, MaMay)
                );
            ";
            command.ExecuteNonQuery();

            // Bảng Giá Nhân công theo Bộ
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS GiaNhanCongTheoBo (
                    BoDonGiaId INTEGER,
                    MaNC TEXT,
                    DonGia REAL,
                    PRIMARY KEY(BoDonGiaId, MaNC)
                );
            ";
            command.ExecuteNonQuery();

            // Bảng 3.3: Định mức Chi phí chung (CPC) theo TT 36/2026
            // CPC phụ thuộc loại CT + giá trị CP XD trong TMĐT (tỷ đồng)
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS DinhMucCPC (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    LoaiCongTrinh TEXT NOT NULL,
                    PhanLoaiPhu TEXT,
                    QuyMoMin REAL NOT NULL,
                    QuyMoMax REAL,
                    TiLe REAL NOT NULL
                );
            ";
            command.ExecuteNonQuery();

            // Bảng 3.5: Định mức CP một số công việc không XĐ được KL từ TK
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS DinhMucTT (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    LoaiCongTrinh TEXT NOT NULL,
                    PhanLoaiPhu TEXT,
                    TiLe REAL NOT NULL
                );
            ";
            command.ExecuteNonQuery();

            // Seed dữ liệu Bảng 3.3 và 3.5
            SeedDinhMucChiPhiChung(connection);
        }

        /// <summary>
        /// Seed dữ liệu Bảng 3.3 (CPC) và Bảng 3.5 (TT) theo TT 36/2026/TT-BXD.
        /// Chỉ insert nếu bảng chưa có dữ liệu.
        /// </summary>
        private void SeedDinhMucChiPhiChung(IDbConnection connection)
        {
            using var checkCmd = connection.CreateCommand();
            checkCmd.CommandText = "SELECT COUNT(*) FROM DinhMucCPC";
            var count = (long)checkCmd.ExecuteScalar();
            if (count > 0) return; // Đã có dữ liệu

            // === BẢNG 3.3: CPC (%) theo loại CT và quy mô CP XD (tỷ đồng) ===
            // Các mốc: ≤40, ≤60, ≤100, ≤300, ≤500, ≤750, ≤1000, >1000
            var cpcData = new[]
            {
                // Công trình dân dụng
                ("Dân dụng", (string)null, new[] { (0m, 40m, 7.3m), (40m, 60m, 7.1m), (60m, 100m, 6.7m), (100m, 300m, 6.5m), (300m, 500m, 6.2m), (500m, 750m, 6.1m), (750m, 1000m, 6.0m), (1000m, (decimal?)null, 5.8m) }),
                // Riêng CT tu bổ di tích lịch sử, văn hóa
                ("Dân dụng", "Tu bổ di tích", new[] { (0m, 40m, 11.6m), (40m, 60m, 11.1m), (60m, 100m, 10.3m), (100m, 300m, 10.1m), (300m, 500m, 9.9m), (500m, 750m, 9.8m), (750m, 1000m, 9.6m), (1000m, (decimal?)null, 9.4m) }),
                // Công trình công nghiệp
                ("Công nghiệp", (string)null, new[] { (0m, 40m, 6.2m), (40m, 60m, 6.0m), (60m, 100m, 5.6m), (100m, 300m, 5.3m), (300m, 500m, 5.1m), (500m, 750m, 5.0m), (750m, 1000m, 4.9m), (1000m, (decimal?)null, 4.6m) }),
                // Riêng CT XD đường hầm thủy điện, hầm lò
                ("Công nghiệp", "Đường hầm thủy điện, hầm lò", new[] { (0m, 40m, 7.3m), (40m, 60m, 7.2m), (60m, 100m, 7.1m), (100m, 300m, 6.9m), (300m, 500m, 6.7m), (500m, 750m, 6.6m), (750m, 1000m, 6.5m), (1000m, (decimal?)null, 6.4m) }),
                // Công trình giao thông
                ("Giao thông", (string)null, new[] { (0m, 40m, 6.2m), (40m, 60m, 6.0m), (60m, 100m, 5.6m), (100m, 300m, 5.3m), (300m, 500m, 5.1m), (500m, 750m, 5.0m), (750m, 1000m, 4.9m), (1000m, (decimal?)null, 4.6m) }),
                // Riêng công trình hầm giao thông
                ("Giao thông", "Hầm giao thông", new[] { (0m, 40m, 7.3m), (40m, 60m, 7.2m), (60m, 100m, 7.1m), (100m, 300m, 6.9m), (300m, 500m, 6.7m), (500m, 750m, 6.6m), (750m, 1000m, 6.5m), (1000m, (decimal?)null, 6.4m) }),
                // Công trình nông nghiệp và môi trường
                ("Nông nghiệp và môi trường", (string)null, new[] { (0m, 40m, 6.1m), (40m, 60m, 5.9m), (60m, 100m, 5.5m), (100m, 300m, 5.3m), (300m, 500m, 5.1m), (500m, 750m, 5.0m), (750m, 1000m, 4.8m), (1000m, (decimal?)null, 4.6m) }),
                // Riêng công trình đường hầm
                ("Nông nghiệp và môi trường", "Đường hầm", new[] { (0m, 40m, 7.3m), (40m, 60m, 7.2m), (60m, 100m, 7.1m), (100m, 300m, 6.9m), (300m, 500m, 6.7m), (500m, 750m, 6.6m), (750m, 1000m, 6.5m), (1000m, (decimal?)null, 6.4m) }),
                // Công trình hạ tầng kỹ thuật
                ("Hạ tầng kỹ thuật", (string)null, new[] { (0m, 40m, 5.5m), (40m, 60m, 5.3m), (60m, 100m, 5.0m), (100m, 300m, 4.8m), (300m, 500m, 4.5m), (500m, 750m, 4.4m), (750m, 1000m, 4.3m), (1000m, (decimal?)null, 4.0m) }),
            };

            using var insertCmd = connection.CreateCommand();
            foreach (var (loaiCT, phanLoai, mocs) in cpcData)
            {
                foreach (var (min, max, tiLe) in mocs)
                {
                    insertCmd.CommandText = $@"
                        INSERT INTO DinhMucCPC (LoaiCongTrinh, PhanLoaiPhu, QuyMoMin, QuyMoMax, TiLe)
                        VALUES (@loaiCT, @phanLoai, @min, @max, @tiLe)";
                    insertCmd.Parameters.Clear();
                    
                    var p1 = insertCmd.CreateParameter(); p1.ParameterName = "@loaiCT"; p1.Value = loaiCT; insertCmd.Parameters.Add(p1);
                    var p2 = insertCmd.CreateParameter(); p2.ParameterName = "@phanLoai"; p2.Value = (object)phanLoai ?? System.DBNull.Value; insertCmd.Parameters.Add(p2);
                    var p3 = insertCmd.CreateParameter(); p3.ParameterName = "@min"; p3.Value = min; insertCmd.Parameters.Add(p3);
                    var p4 = insertCmd.CreateParameter(); p4.ParameterName = "@max"; p4.Value = max.HasValue ? (object)max.Value : System.DBNull.Value; insertCmd.Parameters.Add(p4);
                    var p5 = insertCmd.CreateParameter(); p5.ParameterName = "@tiLe"; p5.Value = tiLe; insertCmd.Parameters.Add(p5);
                    
                    insertCmd.ExecuteNonQuery();
                }
            }

            // === BẢNG 3.5: TT (%) — cố định theo loại CT ===
            var ttData = new[]
            {
                ("Dân dụng", (string)null, 2.5m),
                ("Công nghiệp", (string)null, 2.0m),
                ("Công nghiệp", "Đường hầm thủy điện, hầm lò", 6.5m),
                ("Giao thông", (string)null, 2.0m),
                ("Giao thông", "Hầm giao thông", 6.5m),
                ("Nông nghiệp và môi trường", (string)null, 2.0m),
                ("Nông nghiệp và môi trường", "Đường hầm", 6.5m),
                ("Hạ tầng kỹ thuật", (string)null, 2.0m),
            };

            using var insertTT = connection.CreateCommand();
            foreach (var (loaiCT, phanLoai, tiLe) in ttData)
            {
                insertTT.CommandText = @"
                    INSERT INTO DinhMucTT (LoaiCongTrinh, PhanLoaiPhu, TiLe)
                    VALUES (@loaiCT, @phanLoai, @tiLe)";
                insertTT.Parameters.Clear();

                var p1 = insertTT.CreateParameter(); p1.ParameterName = "@loaiCT"; p1.Value = loaiCT; insertTT.Parameters.Add(p1);
                var p2 = insertTT.CreateParameter(); p2.ParameterName = "@phanLoai"; p2.Value = (object)phanLoai ?? System.DBNull.Value; insertTT.Parameters.Add(p2);
                var p3 = insertTT.CreateParameter(); p3.ParameterName = "@tiLe"; p3.Value = tiLe; insertTT.Parameters.Add(p3);

                insertTT.ExecuteNonQuery();
            }
        }
    }
}
