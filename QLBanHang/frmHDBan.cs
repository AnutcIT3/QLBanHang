using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace QLBanHang
{
    public partial class frmHDBan : Form
    {
        Classes.DataProcesser dtBase = new Classes.DataProcesser();
        Classes.Function fns = new Classes.Function();
        private string maNVLogined = "";

        // Constructor nhận mã nhân viên đăng nhập từ Form Login
        public frmHDBan(string maNV)
        {
            InitializeComponent();
            this.maNVLogined = maNV;
        }

        public frmHDBan()
        {
            InitializeComponent();
        }

        private void frmHDBan_Load(object sender, EventArgs e)
        {
            // Yêu cầu 1: Khóa ComboBox Mã NV & Tên NV, tự động điền từ Login
            cboMaNhanVien.Enabled = false;
            txtTenNhanVien.ReadOnly = true;
            cboMaNhanVien.Text = maNVLogined;

            DataTable dtNV = dtBase.DocBang($"SELECT Tennhanvien FROM tblNhanVien WHERE Manhanvien = N'{maNVLogined}'");
            if (dtNV.Rows.Count > 0)
            {
                txtTenNhanVien.Text = dtNV.Rows[0]["Tennhanvien"].ToString();
            }

            // Nạp dữ liệu ban đầu
            LoadComboMaHDBan();
            LoadComboMaHang();
            LoadComboMaKhachHang();
            ResetForm();
        }

        private void LoadComboMaHDBan()
        {
            cboTimKiem.DropDownStyle = ComboBoxStyle.DropDown;
            DataTable dtHD = dtBase.DocBang("SELECT MaHDBan FROM tblHDBan");
            cboTimKiem.DataSource = dtHD;
            cboTimKiem.DisplayMember = "MaHDBan";
            cboTimKiem.ValueMember = "MaHDBan";
            cboTimKiem.SelectedIndex = -1;
        }

        private void LoadComboMaHang()
        {
            DataTable dtHang = dtBase.DocBang("SELECT Mahang, TenHang FROM tblHang");
            fns.FillComBox(cboMaHang, dtHang, "Mahang", "Mahang");

            txtTenHang.ReadOnly = true;
            txtDonGia.ReadOnly = true;
            txtThanhTien.ReadOnly = true;
        }

        private void LoadComboMaKhachHang()
        {
            cboMaKhachHang.DropDownStyle = ComboBoxStyle.DropDown;
            DataTable dtKhach = dtBase.DocBang("SELECT Makhach, TenKhach FROM tblKhach");
            cboMaKhachHang.DataSource = dtKhach;
            cboMaKhachHang.DisplayMember = "Makhach";
            cboMaKhachHang.ValueMember = "Makhach";
            cboMaKhachHang.SelectedIndex = -1;
        }

        private void txtSoLuong_TextChanged(object sender, EventArgs e)
        {
            TinhThanhTien();
        }

        private void TinhThanhTien()
        {
            try
            {
                double DonGia = 0, GiamGia = 0, soLuong = 0;
                double.TryParse(txtGiamGia.Text, out GiamGia);
                double.TryParse(txtSoLuong.Text, out soLuong);
                double.TryParse(txtDonGia.Text, out DonGia);

                double thanhTien = soLuong * DonGia - (soLuong * DonGia * GiamGia / 100);
                txtThanhTien.Text = thanhTien > 0 ? thanhTien.ToString() : "0";
            }
            catch
            {
                txtThanhTien.Text = "0";
            }
        }

        private void cboMaHang_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboMaHang.SelectedValue != null && cboMaHang.SelectedIndex != -1)
            {
                DataTable dt = dtBase.DocBang($"SELECT TenHang, DonGiaBan FROM tblHang WHERE Mahang = N'{cboMaHang.SelectedValue}'");
                if (dt.Rows.Count > 0)
                {
                    txtTenHang.Text = dt.Rows[0]["TenHang"].ToString();
                    txtDonGia.Text = dt.Rows[0]["DonGiaBan"].ToString();
                    TinhThanhTien();
                }
            }
        }

        private void txtDienThoai_Leave(object sender, EventArgs e)
        {
            string sdt = txtDienThoai.Text.Trim();
            if (string.IsNullOrEmpty(sdt)) return;

            DataTable dtKH = dtBase.DocBang($"SELECT * FROM tblKhach WHERE DienThoai = N'{sdt}'");
            if (dtKH.Rows.Count > 0)
            {
                cboMaKhachHang.Text = dtKH.Rows[0]["Makhach"].ToString();
                txtTenKhachHang.Text = dtKH.Rows[0]["TenKhach"].ToString();
                txtDiaChi.Text = dtKH.Rows[0]["DiaChi"].ToString();
            }
            else
            {
                if (MessageBox.Show("Số điện thoại này chưa có trong hệ thống! Thêm khách hàng mới?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    string maKhachMoi = "KH" + DateTime.Now.ToString("yyMMddHHmmss");
                    cboMaKhachHang.DataSource = null;
                    cboMaKhachHang.Text = maKhachMoi;
                    txtTenKhachHang.Clear();
                    txtDiaChi.Clear();
                    txtTenKhachHang.Focus();
                }
            }
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            string maHD = txtMaHoaDon.Text.Trim();
            string maHang = cboMaHang.SelectedValue?.ToString();

            if (string.IsNullOrEmpty(maHD))
            {
                MessageBox.Show("Vui lòng bấm Thêm hóa đơn trước khi lưu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(txtDienThoai.Text) || string.IsNullOrEmpty(txtTenKhachHang.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin Khách hàng!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(maHang) || string.IsNullOrEmpty(txtSoLuong.Text))
            {
                MessageBox.Show("Vui lòng chọn mặt hàng và nhập số lượng!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int slNhap = int.Parse(txtSoLuong.Text);

            // Kiểm tra tồn kho
            DataTable dtTonKho = dtBase.DocBang($"SELECT SoLuong FROM tblHang WHERE Mahang = N'{maHang}'");
            int slTon = Convert.ToInt32(dtTonKho.Rows[0]["SoLuong"]);
            if (slNhap > slTon)
            {
                MessageBox.Show($"Số lượng tồn kho không đủ! Trong kho chỉ còn {slTon} sản phẩm.", "Lỗi tồn kho", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 1. Lưu Khách hàng nếu chưa có
            string maKhach = cboMaKhachHang.Text.Trim();
            DataTable dtCheckKH = dtBase.DocBang($"SELECT * FROM tblKhach WHERE Makhach = N'{maKhach}'");
            if (dtCheckKH.Rows.Count == 0)
            {
                dtBase.CapNhat($"INSERT INTO tblKhach(Makhach, TenKhach, DiaChi, DienThoai) VALUES(N'{maKhach}', N'{txtTenKhachHang.Text}', N'{txtDiaChi.Text}', N'{txtDienThoai.Text}')");
            }

            // 2. Lưu Hóa đơn chính nếu chưa có
            DataTable dtCheckHD = dtBase.DocBang($"SELECT * FROM tblHDBan WHERE MaHDBan = N'{maHD}'");
            if (dtCheckHD.Rows.Count == 0)
            {
                string sqlHD = $"INSERT INTO tblHDBan(MaHDBan, Manhanvien, NgayBan, Makhach, TongTien) " +
                               $"VALUES(N'{maHD}', N'{maNVLogined}', '{dtpNgayBan.Value:yyyy-MM-dd}', N'{maKhach}', 0)";
                dtBase.CapNhat(sqlHD);
            }

            // 3. Thêm/Cập nhật sản phẩm vào tblChiTietHDBan (Tránh lỗi trùng khóa chính)
            double giamGia = double.Parse(string.IsNullOrEmpty(txtGiamGia.Text) ? "0" : txtGiamGia.Text);
            double thanhTien = double.Parse(txtThanhTien.Text);

            DataTable dtCheckCT = dtBase.DocBang($"SELECT * FROM tblChiTietHDBan WHERE MaHDBan = N'{maHD}' AND Mahang = N'{maHang}'");
            if (dtCheckCT.Rows.Count > 0)
            {
                int slCu = Convert.ToInt32(dtCheckCT.Rows[0]["SoLuong"]);
                int slMoi = slCu + slNhap;
                double donGia = double.Parse(txtDonGia.Text);
                double thanhTienMoi = (slMoi * donGia) - ((slMoi * donGia) * giamGia / 100);

                string sqlUpdateCT = $"UPDATE tblChiTietHDBan SET SoLuong = {slMoi}, GiamGia = {giamGia}, ThanhTien = {thanhTienMoi} " +
                                     $"WHERE MaHDBan = N'{maHD}' AND Mahang = N'{maHang}'";
                dtBase.CapNhat(sqlUpdateCT);
            }
            else
            {
                string sqlCT = $"INSERT INTO tblChiTietHDBan(MaHDBan, Mahang, SoLuong, GiamGia, ThanhTien) " +
                               $"VALUES(N'{maHD}', N'{maHang}', {slNhap}, {giamGia}, {thanhTien})";
                dtBase.CapNhat(sqlCT);
            }

            // 4. Trừ số lượng kho
            dtBase.CapNhat($"UPDATE tblHang SET SoLuong = SoLuong - {slNhap} WHERE Mahang = N'{maHang}'");

            // 5. Cập nhật Tổng tiền & nạp lại giao diện
            TinhTongTien(maHD);
            LoadGridChiTiet(maHD);
            LoadComboMaHDBan();

            // Mở lại các nút bấm
            btnInHoaDon.Enabled = true;
            btnHuyHoaDon.Enabled = true;
            btnThemHoaDon.Enabled = true; // KÍCH HOẠT LẠI NÚT THÊM HÓA ĐƠN

            MessageBox.Show("Thêm sản phẩm vào hóa đơn thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void TinhTongTien(string maHD)
        {
            DataTable dtSum = dtBase.DocBang($"SELECT SUM(CAST(ThanhTien AS FLOAT)) FROM tblChiTietHDBan WHERE MaHDBan = N'{maHD}'");
            double tongTien = 0;
            if (dtSum.Rows.Count > 0 && dtSum.Rows[0][0] != DBNull.Value)
            {
                tongTien = Convert.ToDouble(dtSum.Rows[0][0]);
            }
            txtTongTien.Text = tongTien.ToString();
            dtBase.CapNhat($"UPDATE tblHDBan SET TongTien = {tongTien} WHERE MaHDBan = N'{maHD}'");
        }

        private void LoadGridChiTiet(string maHD)
        {
            string sql = $"SELECT a.Mahang, b.TenHang, a.SoLuong, b.DonGiaBan, a.GiamGia, a.ThanhTien " +
                         $"FROM tblChiTietHDBan a JOIN tblHang b ON a.Mahang = b.Mahang WHERE a.MaHDBan = N'{maHD}'";
            dgvChiTietHDBan.DataSource = dtBase.DocBang(sql);
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            string maHD = cboTimKiem.Text.Trim();
            if (string.IsNullOrEmpty(maHD))
            {
                MessageBox.Show("Vui lòng nhập hoặc chọn Mã hóa đơn cần tìm!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataTable dtHD = dtBase.DocBang($"SELECT * FROM tblHDBan WHERE MaHDBan = N'{maHD}'");
            if (dtHD.Rows.Count > 0)
            {
                txtMaHoaDon.Text = maHD;
                dtpNgayBan.Value = Convert.ToDateTime(dtHD.Rows[0]["NgayBan"]);

                string maKhach = dtHD.Rows[0]["Makhach"].ToString();
                cboMaKhachHang.Text = maKhach;

                DataTable dtKH = dtBase.DocBang($"SELECT * FROM tblKhach WHERE Makhach = N'{maKhach}'");
                if (dtKH.Rows.Count > 0)
                {
                    txtTenKhachHang.Text = dtKH.Rows[0]["TenKhach"].ToString();
                    txtDiaChi.Text = dtKH.Rows[0]["DiaChi"].ToString();
                    txtDienThoai.Text = dtKH.Rows[0]["DienThoai"].ToString();
                }

                txtTongTien.Text = dtHD.Rows[0]["TongTien"].ToString();
                LoadGridChiTiet(maHD);
                btnInHoaDon.Enabled = true;
                btnHuyHoaDon.Enabled = true;
                btnThemHoaDon.Enabled = true;
            }
            else
            {
                MessageBox.Show("Không tìm thấy mã hóa đơn này!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ResetForm()
        {
            txtMaHoaDon.Clear();
            cboTimKiem.Text = "";
            txtTenKhachHang.Clear();
            txtDiaChi.Clear();
            txtDienThoai.Clear();
            cboMaHang.SelectedIndex = -1;
            txtTenHang.Clear();
            txtSoLuong.Clear();
            txtGiamGia.Text = "0";
            txtDonGia.Clear();
            txtThanhTien.Clear();
            txtTongTien.Text = "0";
            dgvChiTietHDBan.DataSource = null;

            // Giữ nguyên thông tin NV đăng nhập
            cboMaNhanVien.Text = maNVLogined;
        }

        private void btnThemHoaDon_Click(object sender, EventArgs e)
        {
            ResetForm();

            // Tiền tố mã hóa đơn theo ngày: HĐB_ddMMyyyy (Ví dụ: HĐB_13102026)
            string prefix = "HĐB_" + DateTime.Now.ToString("ddMMyyyy");

            // Lấy mã hóa đơn lớn nhất trong ngày hiện tại
            string sql = $"SELECT MAX(MaHDBan) FROM tblHDBan WHERE MaHDBan LIKE N'{prefix}%'";
            DataTable dtMax = dtBase.DocBang(sql);

            int stt = 1;
            if (dtMax.Rows.Count > 0 && dtMax.Rows[0][0] != DBNull.Value)
            {
                string maxMaHD = dtMax.Rows[0][0].ToString(); // Ví dụ: HĐB_131020260001

                // Cắt lấy 4 số cuối cùng và chuyển sang số nguyên
                string lastNum = maxMaHD.Substring(maxMaHD.Length - 4);
                stt = Convert.ToInt32(lastNum) + 1;
            }

            // Sinh mã mới định dạng 4 chữ số đuôi: HĐB_131020260001, HĐB_131020260002,...
            txtMaHoaDon.Text = prefix + stt.ToString("D4");

            dtpNgayBan.Value = DateTime.Now;
            btnLuu.Enabled = true;
            btnThemHoaDon.Enabled = false; // Tạm khóa cho đến khi ấn Lưu hoặc Reset
        }

        private void dgvChiTietHDBan_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvChiTietHDBan.CurrentRow == null) return;

            string maHD = txtMaHoaDon.Text.Trim();
            string maHang = dgvChiTietHDBan.CurrentRow.Cells["Mahang"].Value.ToString();
            int slTrả = Convert.ToInt32(dgvChiTietHDBan.CurrentRow.Cells["SoLuong"].Value);

            if (MessageBox.Show($"Bạn có chắc chắn muốn trả lại sản phẩm {maHang}?", "Xác nhận trả hàng", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                dtBase.CapNhat($"DELETE FROM tblChiTietHDBan WHERE MaHDBan = N'{maHD}' AND Mahang = N'{maHang}'");
                dtBase.CapNhat($"UPDATE tblHang SET SoLuong = SoLuong + {slTrả} WHERE Mahang = N'{maHang}'");
                TinhTongTien(maHD);
                LoadGridChiTiet(maHD);
                MessageBox.Show("Trả sản phẩm thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnHuyHoaDon_Click(object sender, EventArgs e)
        {
            string maHD = txtMaHoaDon.Text.Trim();
            if (string.IsNullOrEmpty(maHD)) return;

            if (MessageBox.Show($"Bạn có muốn HỦY TOÀN BỘ hóa đơn {maHD} không?", "Xác nhận hủy", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                DataTable dtCT = dtBase.DocBang($"SELECT Mahang, SoLuong FROM tblChiTietHDBan WHERE MaHDBan = N'{maHD}'");
                foreach (DataRow row in dtCT.Rows)
                {
                    string maHang = row["Mahang"].ToString();
                    int sl = Convert.ToInt32(row["SoLuong"]);
                    dtBase.CapNhat($"UPDATE tblHang SET SoLuong = SoLuong + {sl} WHERE Mahang = N'{maHang}'");
                }

                dtBase.CapNhat($"DELETE FROM tblChiTietHDBan WHERE MaHDBan = N'{maHD}'");
                dtBase.CapNhat($"DELETE FROM tblHDBan WHERE MaHDBan = N'{maHD}'");

                MessageBox.Show("Đã hủy hóa đơn và khôi phục số lượng kho thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ResetForm();
                LoadComboMaHDBan();
                btnThemHoaDon.Enabled = true;
            }
        }

        private void btnInHoaDon_Click(object sender, EventArgs e)
        {
            string maHD = txtMaHoaDon.Text.Trim();
            if (string.IsNullOrEmpty(maHD)) return;

            Excel.Application exApp = new Excel.Application();
            Excel.Workbook exBook = exApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet exSheet = (Excel.Worksheet)exBook.Worksheets[1];
            Excel.Range range = (Excel.Range)exSheet.Cells[1, 1];

            range.Range["A1:F1"].Merge();
            range.Range["A1"].Value = "HÓA ĐƠN BÁN HÀNG";
            range.Range["A1"].Font.Size = 18;
            range.Range["A1"].Font.Bold = true;
            range.Range["A1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            range.Range["A3"].Value = $"Mã Hóa Đơn: {maHD}";
            range.Range["A4"].Value = $"Ngày Bán: {dtpNgayBan.Value:dd/MM/yyyy}";
            range.Range["A5"].Value = $"Nhân Viên: {txtTenNhanVien.Text}";
            range.Range["D3"].Value = $"Khách Hàng: {txtTenKhachHang.Text}";
            range.Range["D4"].Value = $"Điện Thoại: {txtDienThoai.Text}";
            range.Range["D5"].Value = $"Địa Chỉ: {txtDiaChi.Text}";

            int hang = 7;
            range.Range[$"A{hang}"].Value = "STT";
            range.Range[$"B{hang}"].Value = "Tên Hàng";
            range.Range[$"C{hang}"].Value = "Số Lượng";
            range.Range[$"D{hang}"].Value = "Đơn Giá";
            range.Range[$"E{hang}"].Value = "Giảm Giá (%)";
            range.Range[$"F{hang}"].Value = "Thành Tiền";
            range.Range[$"A{hang}:F{hang}"].Font.Bold = true;

            DataTable dtCT = dtBase.DocBang($"SELECT b.TenHang, a.SoLuong, b.DonGiaBan, a.GiamGia, a.ThanhTien FROM tblChiTietHDBan a JOIN tblHang b ON a.Mahang = b.Mahang WHERE a.MaHDBan = N'{maHD}'");
            for (int i = 0; i < dtCT.Rows.Count; i++)
            {
                hang++;
                range.Range[$"A{hang}"].Value = (i + 1).ToString();
                range.Range[$"B{hang}"].Value = dtCT.Rows[i]["TenHang"].ToString();
                range.Range[$"C{hang}"].Value = dtCT.Rows[i]["SoLuong"].ToString();
                range.Range[$"D{hang}"].Value = dtCT.Rows[i]["DonGiaBan"].ToString();
                range.Range[$"E{hang}"].Value = dtCT.Rows[i]["GiamGia"].ToString();
                range.Range[$"F{hang}"].Value = dtCT.Rows[i]["ThanhTien"].ToString();
            }

            hang++;
            range.Range[$"E{hang}"].Value = "Tổng Tiền:";
            range.Range[$"F{hang}"].Value = txtTongTien.Text;
            range.Range[$"E{hang}:F{hang}"].Font.Bold = true;

            exSheet.Columns.AutoFit();
            exBook.Activate();

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Files (*.xlsx)|*.xlsx";
            sfd.FileName = $"HoaDon_{maHD}";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                exBook.SaveAs(sfd.FileName);
                MessageBox.Show("In/Xuất hóa đơn ra file Excel thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            exApp.Quit();
        }

        private void btnDong_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}