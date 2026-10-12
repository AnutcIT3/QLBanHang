using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QLBanHang
{
    public partial class Login : Form
    {
        Classes.DataProcesser dt = new Classes.DataProcesser();

        public Login()
        {
            InitializeComponent();
        }

        private void Login_Load(object sender, EventArgs e)
        {

        }

        private void btnDangNhap_Click(object sender, EventArgs e)
        {
            string user = txtMaNV.Text.Trim();
            string pass = txtMatKhau.Text.Trim();

            // Kiểm tra rỗng[cite: 10]
            if (string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ Tên đăng nhập và Mật khẩu!", "Thông báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Truy vấn kiểm tra tài khoản trong bảng tblUser[cite: 21]
            string sql = $"SELECT * FROM tblNhanVien WHERE Tennhanvien = N'{user}' AND CAST(Pass AS NVARCHAR(50)) = N'{pass}'";
            DataTable dtUser = dt.DocBang(sql);

            if (dtUser.Rows.Count > 0)
            {
                MessageBox.Show("Đăng nhập thành công!", "Thông báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Mở Form Cập nhật mặt hàng (frmSanPham)[cite: 21]
                frmSanPham mainForm = new frmSanPham(user);
                mainForm.Show();

                // Ẩn Form Đăng nhập hiện tại
                this.Hide();
            }
            else
            {
                MessageBox.Show("Tên đăng nhập hoặc mật khẩu không chính xác!", "Lỗi đăng nhập",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMatKhau.Clear();
                txtMatKhau.Focus();
            }
        }

        private void btnHuyBo_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Bạn có muốn thoát chương trình không?", "Xác nhận",
                                              MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                Application.Exit();
            }
        }
    }
}
