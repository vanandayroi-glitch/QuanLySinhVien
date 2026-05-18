using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.IO.Ports;

namespace QuanLySinhVien
{
    public partial class Form1 : Form
    {
        // Khai báo chuỗi kết nối (ConnectionString) và đối tượng kết nối SQL
        string strCon = @"Data Source=MSI\SQLEXPRESS;Initial Catalog=QuanLySV;Integrated Security=True;TrustServerCertificate=True";
        SqlConnection sqlCon = null;

        // KHAI BÁO THÊM BIẾN CỔNG COM ĐỂ QUẢN LÝ
        SerialPort mySerialPort = null;

        public Form1()
        {
            InitializeComponent();
            NapMaKhoaVaoComboBox();
            NapMaLopVaoComboBox();
            HienThiDanhSachCongCOM();
        }

        // KHỐI KẾT NỐI DATABASE DÙNG CHUNG
        void MoKetNoi()
        {
            if (sqlCon == null) sqlCon = new SqlConnection(strCon);
            if (sqlCon.State == ConnectionState.Closed) sqlCon.Open();
        }

        void DongKetNoi()
        {
            if (sqlCon != null && sqlCon.State == ConnectionState.Open) sqlCon.Close();
        }

        // KHỐI RESET SAU KHI THỰC HIỆN XONG THAO TÁC
        void ResetKhoa()
        {
            txtMaKhoa.Clear();
            txtTenKhoa.Clear();
            txtDiaChi.Clear();
            txtMaKhoa.ReadOnly = false; // Mở khóa ô mã khoa để sẵn sàng thêm mới
            txtMaKhoa.Focus();
        }
        void ResetLop()
        {
            txtMaLop.Clear();
            txtKhoa.Clear();
            txtTenLop.Clear();
            cboMaKhoa.SelectedIndex = -1; // Đưa hộp chọn về trắng trơn
            txtMaLop.ReadOnly = false;   // Mở khóa ô mã lớp để sẵn sàng thêm mới
            txtMaLop.Focus();
        }
        void ResetSinhVien()
        {
            txtMaSV.Clear();
            txtHoTen.Clear();
            dtpNgaySinh.Value = DateTime.Now;
            cboMaLop.SelectedIndex = -1; // Đưa hộp chọn về trắng trơn
            picAnhSV.Image = null;      // Xóa sạch khung ảnh cũ
            txtMaSV.ReadOnly = false;   // Mở khóa ô mã sinh viên để sẵn sàng thêm mới
            txtMaSV.Focus();
        }
        // KHỐI XỬ LÝ ẢNH 
        byte[] ChuyenAnhThanhByte(Image img)
        {
            if (img == null) return null;
            using (MemoryStream ms = new MemoryStream())
            {
                using (Bitmap bmp = new Bitmap(img))
                {
                    bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                    return ms.ToArray();
                }
            }
        }
        Image ByteThanhAnh(byte[] data)
        {
            using (MemoryStream ms = new MemoryStream(data))
            {
                return new Bitmap(Image.FromStream(ms)); // Trả về dạng bản sao Bitmap tránh lỗi GDI+ khi Sửa
            }
        }
        // Tìm các cổng COM ảo
        void HienThiDanhSachCongCOM()
        {
            string[] ports = SerialPort.GetPortNames();
            cboComPort.Items.Clear();
            cboComPort.Items.AddRange(ports);
            if (cboComPort.Items.Count > 0) cboComPort.SelectedIndex = -1; 
        }

        // ==========================================
        // TAB 1: QUẢN LÝ KHOA (BẢNG CHA LỚN)
        // ==========================================
        void LayDuLieuKhoa()
        {
            try
            {
                MoKetNoi();
                SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM Khoa", sqlCon);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                dgvKhoa.DataSource = dt;
                dgvKhoa.AllowUserToAddRows = false;
                dgvKhoa.RowTemplate.Height = 35; // Hàng chữ đặt cao vừa phải
            }
            catch (Exception ex) { MessageBox.Show("Lỗi tải dữ liệu khoa: " + ex.Message); }
            finally { DongKetNoi(); }
        }
        void NapMaKhoaVaoComboBox()
        {
            try
            {
                MoKetNoi();
                SqlDataAdapter adapter = new SqlDataAdapter("SELECT MaKhoa FROM Khoa", sqlCon);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                cboMaKhoa.DataSource = dt;
                cboMaKhoa.DisplayMember = "MaKhoa"; // Hiển thị Mã khoa lên thanh xổ
                cboMaKhoa.ValueMember = "MaKhoa";   // Giá trị lấy ra ngầm cũng là Mã khoa
                cboMaKhoa.SelectedIndex = -1;     // Trắng trơn ban đầu
            }
            catch (Exception ex) { MessageBox.Show("Lỗi nạp danh sách khoa: " + ex.Message); }
            finally { DongKetNoi(); }
        }

        private void btnHienThiKhoa_Click(object sender, EventArgs e)
        {
            LayDuLieuKhoa();
            ResetKhoa();
        }

        private void btnThemKhoa_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtMaKhoa.Text.Trim() == "" || txtTenKhoa.Text.Trim() == "" || txtDiaChi.Text.Trim() == "")
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin khoa!", "Thông báo");
                    return;
                }

                MoKetNoi();
                SqlCommand cmd = new SqlCommand("INSERT INTO Khoa(MaKhoa, TenKhoa, DiaChi) VALUES(@MaKhoa, @TenKhoa, @DiaChi)", sqlCon);
                cmd.Parameters.AddWithValue("@MaKhoa", txtMaKhoa.Text.Trim());
                cmd.Parameters.AddWithValue("@TenKhoa", txtTenKhoa.Text.Trim());
                cmd.Parameters.AddWithValue("@DiaChi", txtDiaChi.Text.Trim());
                if (cmd.ExecuteNonQuery() > 0)
                {
                    MessageBox.Show("Thêm khoa thành công!", "Thành công");
                    LayDuLieuKhoa();
                    NapMaKhoaVaoComboBox(); // Cập nhật lại ComboBox ở Tab Lớp ngay lập tức
                    ResetKhoa();
                }
            }
            catch (Exception ex) { MessageBox.Show("Lỗi thêm khoa (Trùng mã khóa chính): " + ex.Message); }
            finally { DongKetNoi(); }
        }

        private void btnSuaKhoa_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtMaKhoa.Text.Trim() == "") { MessageBox.Show("Vui lòng nhập mã hoặc chọn khoa cần sửa!"); txtMaKhoa.Focus(); return; }
                if (txtTenKhoa.Text.Trim() == "") { MessageBox.Show("Vui lòng nhập Tên khoa!"); txtTenKhoa.Focus(); return; }
                if (txtDiaChi.Text.Trim() == "") { MessageBox.Show("Vui lòng nhập Địa Chỉ!"); txtDiaChi.Focus(); return; }
                MoKetNoi();
                SqlCommand cmd = new SqlCommand("UPDATE Khoa SET TenKhoa=@Ten, DiaChi=@DiaChi WHERE MaKhoa=@Ma", sqlCon);
                cmd.Parameters.AddWithValue("@Ma", txtMaKhoa.Text.Trim());
                cmd.Parameters.AddWithValue("@Ten", txtTenKhoa.Text.Trim());
                cmd.Parameters.AddWithValue("@DiaChi", txtDiaChi.Text.Trim());


                if (cmd.ExecuteNonQuery() > 0)
                {
                    MessageBox.Show("Cập nhật khoa thành công!");
                    LayDuLieuKhoa();
                    ResetKhoa();
                }
                else
                {
                    MessageBox.Show("Không tìm thấy mã khoa trong hệ thống!");
                }
            }
            catch (Exception ex) { MessageBox.Show("Lỗi sửa Khoa: " + ex.Message); }
            finally { DongKetNoi(); }
        }

        private void btnXoaKhoa_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtMaKhoa.Text.Trim() == "") { MessageBox.Show("Vui lòng chọn khoa cần xóa!"); return; }

                if (MessageBox.Show("Bạn có chắc chắn muốn xóa khoa này vĩnh viễn?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    MoKetNoi();
                    SqlCommand cmd = new SqlCommand("DELETE FROM Khoa WHERE MaKhoa=@Ma", sqlCon);
                    cmd.Parameters.AddWithValue("@Ma", txtMaKhoa.Text.Trim());

                    if (cmd.ExecuteNonQuery() > 0)
                    {
                        MessageBox.Show("Xóa khoa thành công!");
                        LayDuLieuKhoa();
                        NapMaKhoaVaoComboBox(); // Cập nhật lại Combobox ở tab Lớp
                        ResetKhoa();
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Không thể xóa! Khoa này đang có các Lớp trực thuộc quản lý.\nChi tiết: " + ex.Message, "Lỗi ràng buộc khóa ngoại"); }
            finally { DongKetNoi(); }
        }
        private void dgvKhoa_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dgvKhoa.Rows[e.RowIndex];
                txtMaKhoa.Text = row.Cells["MaKhoa"].Value.ToString().Trim();
                txtTenKhoa.Text = row.Cells["TenKhoa"].Value.ToString();
                txtDiaChi.Text = row.Cells["DiaChi"].Value.ToString();
                txtMaKhoa.ReadOnly = true; // Khóa khóa chính MaKhoa lại
            }
        }

        // ==========================================
        // TAB 2: QUẢN LÝ LỚP (BẢNG CHA NHỎ)
        // ==========================================
        void LayDuLieuLop()
        {
            try
            {
                MoKetNoi();
                SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM Lop", sqlCon);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                dgvLop.DataSource = dt;
                dgvLop.AllowUserToAddRows = false;
                dgvLop.RowTemplate.Height = 35;

            }
            catch (Exception ex) { MessageBox.Show("Lỗi tải danh sách Lớp: " + ex.Message); }
            finally { DongKetNoi(); }
        }
        void NapMaLopVaoComboBox()
        {
            try
            {
                MoKetNoi();
                SqlDataAdapter adapter = new SqlDataAdapter("SELECT MaLop FROM Lop", sqlCon);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                cboMaLop.DataSource = dt;
                cboMaLop.DisplayMember = "MaLop"; // Hiển thị Mã Lopứ lên thanh xổ
                cboMaLop.ValueMember = "MaLop";   // Giá trị lấy ra ngầm cũng là Mã Lớp
                cboMaLop.SelectedIndex = -1;     // Trắng trơn ban đầu
            }
            catch (Exception ex) { MessageBox.Show("Lỗi nạp danh sách lớp: " + ex.Message); }
            finally { DongKetNoi(); }
        }
        private void btnHienThiLop_Click(object sender, EventArgs e)
        {
            LayDuLieuLop();
            ResetLop();
        }
        private void btnThemLop_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtMaLop.Text.Trim() == "" || txtTenLop.Text.Trim() == "" || txtKhoa.Text.Trim() == "" || cboMaKhoa.SelectedIndex == -1)
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ Mã, Tên và Khóa!", "Thông báo");
                    return;
                }

                MoKetNoi();
                SqlCommand cmd = new SqlCommand("INSERT INTO Lop(MaLop, TenLop, MaKhoa, Khoa) VALUES(@MaLop, @TenLop, @MaKhoa, @Khoa)", sqlCon);
                cmd.Parameters.AddWithValue("@MaLop", txtMaLop.Text.Trim());
                cmd.Parameters.AddWithValue("@TenLop", txtTenLop.Text.Trim());
                cmd.Parameters.AddWithValue("@Khoa", txtKhoa.Text.Trim());
                cmd.Parameters.AddWithValue("@MaKhoa", cboMaKhoa.SelectedValue.ToString());

                if (cmd.ExecuteNonQuery() > 0)
                {
                    MessageBox.Show("Thêm hồ sơ Lớp thành công!", "Thành công");
                    LayDuLieuLop();
                    NapMaLopVaoComboBox();
                    ResetLop();
                }
            }
            catch (Exception ex) { MessageBox.Show("Lỗi mã lớp (trùng mã khóa chính): " + ex.Message); }
            finally { DongKetNoi(); }
        }

        private void btnSuaLop_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtMaLop.Text.Trim() == "") { MessageBox.Show("Vui lòng nhập Mã Lớp cần sửa!", "Thông báo"); txtMaLop.Focus(); return; }
                if (txtTenLop.Text.Trim() == "") { MessageBox.Show("Vui lòng nhập tên Lớp cần sửa!", "Thông báo"); txtTenLop.Focus(); return; }
                if (txtKhoa.Text.Trim() == "") { MessageBox.Show("Vui lòng nhập Khóa cần sửa!", "Thông báo"); txtKhoa.Focus(); return; }
                if (cboMaKhoa.SelectedIndex == -1) { MessageBox.Show("Vui lòng chọn Mã Khoa!", "Thông báo"); cboMaKhoa.Focus(); return; }

                MoKetNoi();
                SqlCommand cmd = new SqlCommand("UPDATE Lop SET TenLop=@Ten, Khoa=@Khoa, MaKhoa=@MaKhoa WHERE MaLop=@MaLop", sqlCon);
                cmd.Parameters.AddWithValue("@MaLop", txtMaLop.Text.Trim());
                cmd.Parameters.AddWithValue("@Ten", txtTenLop.Text.Trim());
                cmd.Parameters.AddWithValue("@Khoa", txtKhoa.Text.Trim());
                cmd.Parameters.AddWithValue("@MaKhoa", cboMaKhoa.SelectedValue.ToString());

                if (cmd.ExecuteNonQuery() > 0)
                {
                    MessageBox.Show("Cập nhật hồ sơ Lớp thành công!");
                    LayDuLieuLop();
                    NapMaLopVaoComboBox();
                    ResetLop();
                }
                else
                {
                    MessageBox.Show("Không tìm thấy mã Lớp trong hệ thống!");
                }
            }
            catch (Exception ex) { MessageBox.Show("Lỗi sửa lớp: " + ex.Message); }
            finally { DongKetNoi(); }
        }
        private void btnXoaLop_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtMaLop.Text.Trim() == "") { MessageBox.Show("Vui lòng chọn lớp cần xóa!"); return; }

                if (MessageBox.Show("Bạn có chắc chắn muốn xóa lớp này khỏi hệ thống?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    MoKetNoi();
                    SqlCommand cmd = new SqlCommand("DELETE FROM Lop WHERE MaLop=@MaLop", sqlCon);
                    cmd.Parameters.AddWithValue("@MaLop", txtMaLop.Text.Trim());

                    if (cmd.ExecuteNonQuery() > 0)
                    {
                        MessageBox.Show("Xóa lớp thành công!");
                        LayDuLieuLop();
                        NapMaLopVaoComboBox();
                        ResetLop();
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Lỗi xóa lớp: " + ex.Message); }
            finally { DongKetNoi(); }
        }

        private void dgvLop_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dgvLop.Rows[e.RowIndex];
                txtMaLop.Text = row.Cells["MaLop"].Value.ToString().Trim();
                txtTenLop.Text = row.Cells["TenLop"].Value.ToString();
                txtKhoa.Text = row.Cells["Khoa"].Value.ToString();
                cboMaKhoa.SelectedValue = row.Cells["MaKhoa"].Value.ToString().Trim();
                txtMaLop.ReadOnly = true; // Khóa khóa chính MaLop lại
            }
        }
        // ==========================================
        // TAB 3: QUẢN LÝ SINH VIÊN (BẢNG CON)
        // ==========================================
        void LayDuLieuSinhVien()
        {
            try
            {
                MoKetNoi();
                SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM SinhVien", sqlCon);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                dgvSV.DataSource = dt;
                dgvSV.AllowUserToAddRows = false;
                dgvSV.RowTemplate.Height = 60; // Nâng độ cao lên 60px để có khoảng không gian hiện ảnh thẻ

                if (dgvSV.Columns["AnhSV"] != null)
                {
                    DataGridViewImageColumn imgCol = (DataGridViewImageColumn)dgvSV.Columns["AnhSV"];
                    imgCol.ImageLayout = DataGridViewImageCellLayout.Stretch; // Tự co giãn vừa ô ảnh
                }
            }
            catch (Exception ex) { MessageBox.Show("Lỗi tải danh sách sinh viên: " + ex.Message); }
            finally { DongKetNoi(); }
        }
        
        private void btnHienThiSV_Click(object sender, EventArgs e)
        {
            LayDuLieuSinhVien();
            ResetSinhVien();
        }

        private void btnThemSV_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtMaSV.Text.Trim() == "" || txtHoTen.Text.Trim() == "" || cboMaLop.SelectedIndex == -1)
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin!", "Thông báo");
                    return;
                }

                byte[] hinhAnh = ChuyenAnhThanhByte(picAnhSV.Image);

                MoKetNoi();
                SqlCommand cmd = new SqlCommand("INSERT INTO SinhVien(MaSV, HoTen, AnhSV, MaLop, NgaySinh) VALUES(@MaSV, @HoTen, @AnhSV, @MaLop, @NgaySinh)", sqlCon);
                cmd.Parameters.AddWithValue("@MaSV", txtMaSV.Text.Trim());
                cmd.Parameters.AddWithValue("@HoTen", txtHoTen.Text.Trim());
                cmd.Parameters.AddWithValue("@AnhSV", (object)hinhAnh ?? DBNull.Value); // Cho phép lưu trống ảnh nếu chưa chuẩn bị kịp
                cmd.Parameters.AddWithValue("@MaLop", cboMaLop.SelectedValue.ToString());
                cmd.Parameters.AddWithValue("@NgaySinh", dtpNgaySinh.Value);
                if (cmd.ExecuteNonQuery() > 0)
                {
                    MessageBox.Show("Thêm hồ sơ Sinh viên thành công!", "Thành công");
                    LayDuLieuSinhVien();
                    ResetSinhVien();
                }
            }
            catch (Exception ex) { MessageBox.Show("Lỗi mã Sinh viên(trùng mã khóa chính): " + ex.Message); }
            finally { DongKetNoi(); }
        }
        private void btnSuaSV_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtMaSV.Text.Trim() == "") { MessageBox.Show("Vui lòng nhập Mã Sinh Viên cần sửa!", "Thông báo"); txtMaSV.Focus(); return; }
                if (txtHoTen.Text.Trim() == "") { MessageBox.Show("Vui lòng nhập Họ tên Sinh viên cần sửa!", "Thông báo"); txtHoTen.Focus(); return; }
                if (cboMaLop.SelectedIndex == -1) { MessageBox.Show("Vui lòng chọn Mã Lớp!", "Thông báo"); cboMaLop.Focus(); return; }
                byte[] hinhAnh = ChuyenAnhThanhByte(picAnhSV.Image);

                MoKetNoi();
                SqlCommand cmd = new SqlCommand("UPDATE SinhVien SET HoTen=@Ten, AnhSV=@Anh, MaLop=@MaLop, NgaySinh=@NgaySinh WHERE MaSV=@MaSV", sqlCon);
                cmd.Parameters.AddWithValue("@MaSV", txtMaSV.Text.Trim());
                cmd.Parameters.AddWithValue("@Ten", txtHoTen.Text.Trim());
                cmd.Parameters.AddWithValue("@Anh", (object)hinhAnh ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@MaLop", cboMaLop.SelectedValue.ToString());
                cmd.Parameters.AddWithValue("@NgaySinh", dtpNgaySinh.Value);

                if (cmd.ExecuteNonQuery() > 0)
                {
                    MessageBox.Show("Cập nhật hồ sơ nhân viên thành công!");
                    LayDuLieuSinhVien();
                    ResetSinhVien();
                }
                else
                {
                    MessageBox.Show("Không tìm thấy mã nhân viên trong hệ thống!");
                }
            }
            catch (Exception ex) { MessageBox.Show("Lỗi sửa sinh viên: " + ex.Message); }
            finally { DongKetNoi(); }
        }
        private void btnXoaSV_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtMaSV.Text.Trim() == "") { MessageBox.Show("Vui lòng chọn Sinh viên cần xóa!"); return; }

                if (MessageBox.Show("Bạn có chắc chắn muốn xóa Sinh viên này khỏi hệ thống?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    MoKetNoi();
                    SqlCommand cmd = new SqlCommand("DELETE FROM SinhVien WHERE MaSV=@MaSV", sqlCon);
                    cmd.Parameters.AddWithValue("@MaSV", txtMaSV.Text.Trim());

                    if (cmd.ExecuteNonQuery() > 0)
                    {
                        MessageBox.Show("Xóa Sinh viên thành công!");
                        LayDuLieuSinhVien();
                        ResetSinhVien();
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Lỗi xóa Sinh viên: " + ex.Message); }
            finally { DongKetNoi(); }
        }
        private void dgvSV_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dgvSV.Rows[e.RowIndex];
                txtMaSV.Text = row.Cells["MaSV"].Value.ToString().Trim();
                txtHoTen.Text = row.Cells["HoTen"].Value.ToString();
                cboMaLop.SelectedValue = row.Cells["MaLop"].Value.ToString().Trim();
                dtpNgaySinh.Value = Convert.ToDateTime(row.Cells["NgaySinh"].Value);

                txtMaSV.ReadOnly = true; // Khóa khóa chính MaNV lại

                if (row.Cells["AnhSV"].Value != DBNull.Value && row.Cells["AnhSV"].Value != null)
                {
                    byte[] data = (byte[])row.Cells["AnhSV"].Value;
                    picAnhSV.Image = ByteThanhAnh(data); // Nạp ảnh thẻ bằng bản sao độc lập
                }
                else { picAnhSV.Image = null; }
            }
        }
        // CHỌN FILE ẢNH THẺ SINH VIÊN
        private void picAnhSV_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog { Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif" };
            if (open.ShowDialog() == DialogResult.OK)
            {
                picAnhSV.Image = new Bitmap(open.FileName);
            }
        }

        // =====================================
        // TAB 4: CHỨC NĂNG TÌM KIẾM SINH VIÊN 
        // =====================================
        void TimKiemSinhVienTheoMa(string maSV)
        {
            try
            {
                MoKetNoi();

                string query =
                    "SELECT sv.HoTen, sv.AnhSV, sv.NgaySinh, l.TenLop, l.Khoa, kh.TenKhoa " +
                    "FROM SinhVien sv " +
                    "INNER JOIN Lop l ON sv.MaLop = l.MaLop " +
                    "INNER JOIN Khoa kh ON l.MaKhoa = kh.MaKhoa " +
                    "WHERE sv.MaSV = @MaSV";

                SqlCommand cmd = new SqlCommand(query, sqlCon);
                cmd.Parameters.AddWithValue("@MaSV", maSV);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        // Hiển thị lên tab tìm kiếm
                        txtTimMaSV.Text = maSV;

                        txtResultTenSV.Text = reader["HoTen"].ToString();
                        txtResultNgaySinh.Text = string.Format("{0:dd/MM/yyyy}", reader["NgaySinh"]);
                        txtResultTenLop.Text = reader["TenLop"].ToString();
                        txtResultKhoa.Text = reader["Khoa"].ToString();
                        txtResultTenKhoa.Text = reader["TenKhoa"].ToString();

                        if (reader["AnhSV"] != DBNull.Value)
                        {
                            byte[] data = (byte[])reader["AnhSV"];
                            picResultAnhSV.Image = ByteThanhAnh(data);
                        }
                        else
                        {
                            picResultAnhSV.Image = null;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Không tìm thấy sinh viên có mã: " + maSV);

                        txtResultTenSV.Clear();
                        txtResultNgaySinh.Clear();
                        txtResultTenLop.Clear();
                        txtResultKhoa.Clear();
                        txtResultTenKhoa.Clear();
                        picResultAnhSV.Image = null;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tìm kiếm: " + ex.Message);
            }
            finally
            {
                DongKetNoi();
            }
        }
        // Tìm kiếm bằng Button Tìm Kiếm
        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            if (txtTimMaSV.Text.Trim() == "")
            {
                MessageBox.Show("Vui lòng nhập mã sinh viên!");
                return;
            }

            TimKiemSinhVienTheoMa(txtTimMaSV.Text.Trim());
        }
        // Tìm kiếm bằng Com
        // 1. Kết nối Com
        private void btnKetNoiCom_Click(object sender, EventArgs e)
        {
            try
            {
                // Kiểm tra kết nối
                if (mySerialPort == null || !mySerialPort.IsOpen)
                {
                    mySerialPort = new SerialPort();

                    mySerialPort.PortName = cboComPort.Text; // Lấy COM từ ComboBox
                    mySerialPort.BaudRate = 9600;
                    mySerialPort.DataBits = 8;
                    mySerialPort.StopBits = StopBits.One;
                    mySerialPort.Parity = Parity.None;

                    // Gắn sự kiện khi có dữ liệu gửi tới
                    mySerialPort.DataReceived += MySerialPort_DataReceived;

                    mySerialPort.Open();

                    MessageBox.Show("Kết nối COM thành công!");
                    btnKetNoiCom.Text = "Ngắt COM";
                }
                else
                {
                    // Nếu đang mở thì đóng
                    mySerialPort.Close();

                    MessageBox.Show("Đã ngắt kết nối COM!");
                    btnKetNoiCom.Text = "Kết nối COM";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi kết nối COM: " + ex.Message);
            }
        }
        // 2. Nhận dữ liệu từ Terminal
        private void MySerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                // Đọc dữ liệu gửi tới
                string maSV = mySerialPort.ReadExisting().Trim();

                // Ủy quyền (Invoke) truyền dữ liệu từ nguồn ẩn phần cứng về luồng giao diện chính tránh xung đột
                Invoke(new Action(() =>
                {
                    TimKiemSinhVienTheoMa(maSV);
                }));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}