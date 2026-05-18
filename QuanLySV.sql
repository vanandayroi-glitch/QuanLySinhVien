-- 1. Tạo Database

USE QuanLySV;
GO

DROP TABLE IF EXISTS SinhVien;
DROP TABLE IF EXISTS Lop;
DROP TABLE IF EXISTS Khoa;


-- 2. Tạo bảng Khoa (Bảng cha lớn nhất)
CREATE TABLE Khoa(
    MaKhoa CHAR(10) PRIMARY KEY, -- Khóa chính
    TenKhoa NVARCHAR(50),
    DiaChi CHAR(10),
);
GO
-- 3. Tạo bảng Lop (Bảng con của Khoa, cha của Sinh viên)
CREATE TABLE Lop(
    MaLop CHAR(10) PRIMARY KEY, -- Khóa chính
    TenLop NVARCHAR(50),
    Khoa INT,
    MaKhoa CHAR(10), -- Cột liên kết bảng Khoa

    -- Thiết lập khóa ngoại liên kết với bảng Khoa 
    CONSTRAINT FK_Lop_Khoa 
    FOREIGN KEY(MaKhoa) 
    REFERENCES Khoa(MaKhoa)
);
GO
-- 4. Tạo bảng Sinh viên (Bảng con thấp nhất)
CREATE TABLE SinhVien (
    MaSV CHAR(20) PRIMARY KEY, -- Khóa chính
    HoTen NVARCHAR(50),
    NgaySinh DATE,
    AnhSV IMAGE,
    MaLop CHAR(10), -- Cột liên kết sang bảng Lớp

    -- Thiết lập khóa ngoại liên kết với bảng Lớp
    CONSTRAINT FK_SinhVien_Lop 
    FOREIGN KEY(MaLop) 
    REFERENCES Lop(MaLop)
);
GO


-- 5.Kiểm tra kết quả sau khi nạp dữ liệu mẫu
SELECT * FROM Khoa;
SELECT * FROM Lop;
SELECT * FROM SinhVien;
