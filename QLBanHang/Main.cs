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
    public partial class Main : Form
    {
        private string maNVLogined = "";
        public Main()
        {
            InitializeComponent();
        }
        public Main(string maNV)
        {
            InitializeComponent();
            this.maNVLogined = maNV;
        }

        private void Main_Load(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(maNVLogined))
            {
                this.Text = "Hệ thống quản lý bán hàng - Xin chào: " + maNVLogined;
            }
        }

        private void sanPhamToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmSanPham formSP = new frmSanPham(maNVLogined);
            formSP.ShowDialog();
        }

        private void nhanVienToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void hoaDonToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmHDBan formHD = new frmHDBan(maNVLogined);
            formHD.ShowDialog();
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
