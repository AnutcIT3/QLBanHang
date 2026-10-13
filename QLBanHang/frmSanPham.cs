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
    public partial class frmSanPham : Form
    {
        Classes.DataProcesser dt = new Classes.DataProcesser();
        Classes.Function fn = new Classes.Function();
        string fileAnh = "";
        string userName = "";
        public frmSanPham(string user)
        {
            InitializeComponent();
            this.userName = user;
        }

        private void frmSanPham_Load(object sender, EventArgs e)
        {
            //lay dl vao combobox
            DataTable dtChatLieu = dt.DocBang("select * from tblChatLieu");
            fn.FillComBox(cboChatLieu, dtChatLieu, "TenChatLieu", "MaChatLieu");
            cboChatLieu.SelectedIndex = -1;
            //load du lieu vao dgvHang
            DataTable dtHang = dt.DocBang("select * from tblHang");
            dgvHang.DataSource = dtHang;
            //dinh dang datagrid
            dgvHang.Columns[0].HeaderText = "Mã hàng";
            dgvHang.Columns[1].HeaderText = "Tên hàng";
            dgvHang.Columns[2].HeaderText = "Mã CL";
            dgvHang.Columns[3].HeaderText = "Số lượng";
            dgvHang.Columns[4].HeaderText = "Đơn giá nhập";
            dgvHang.Columns[5].HeaderText = "Đơn giá bán";
            dgvHang.Columns[6].HeaderText = "File ảnh";

            dgvHang.Columns[0].Width = 150;
            dgvHang.Columns[1].Width = 250;
            dgvHang.Columns[2].Width = 150;
            dgvHang.Columns[3].Width = 150;
            dgvHang.Columns[4].Width = 150;
            dgvHang.Columns[5].Width = 150;
            dgvHang.Columns[6].Width = 150;
            dgvHang.BackgroundColor = Color.LightBlue;
            dtHang.Dispose();

            btnSua.Enabled = false;
            btnXoa.Enabled = false;
            btnLuu.Enabled = false;
            btnBoQua.Enabled = false;
            //load ten người dùng
            lbXinChao.Text = "Xin chào: " + userName;

        }

        private void btnAnh_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlgAnh = new OpenFileDialog();
            dlgAnh.Filter = "Bitmap(*.bmp)|*.bmp|Gif(*.gif)|*.gif|All files(*.*)|*.*";
            dlgAnh.InitialDirectory = Application.StartupPath;
            dlgAnh.FilterIndex = 3; // Quy định lọc mặc định là bộ lọc thứ 3 (All files)
            dlgAnh.Title = "Chọn ảnh để hiển thị";

            if (dlgAnh.ShowDialog() == DialogResult.OK)
            {
                // Lấy tên, đường dẫn ảnh khi người dùng chọn trong hộp hội thoại OpenDialog.
                // Sau đó gán cho thuộc tính Image của PictureBox
                picAnh.Image = Image.FromFile(dlgAnh.FileName);

                // Cắt lấy riêng tên file ảnh (không bao gồm đường dẫn thư mục) để lưu vào CSDL
                string[] str = dlgAnh.FileName.Split('\\');
                fileAnh = str[str.Length - 1].ToString();
            }
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            btnSua.Enabled = false;
            btnXoa.Enabled = false;
            btnThem.Enabled = true;
            btnLuu.Enabled = true;
            btnBoQua.Enabled = true;
            ResetValue();
            
        }
        void ResetValue()
        {
            txtMaHang.Text = "";
            txtTenHang.Text = "";
            cboChatLieu.Text = "";
            txtSoLuong.Text = "";
            txtDonGiaBan.Text = "";
            txtDonGiaNhap.Text = "";
            picAnh.Image = null;
            txtGhiChu.Text = "";
            txtMaHang.Focus();
        }

        private void btnBoQua_Click(object sender, EventArgs e)
        {
            btnSua.Enabled = false;
            btnXoa.Enabled = false;
            btnThem.Enabled = true;
            btnBoQua.Enabled = false; 
            btnLuu.Enabled = false;
            ResetValue();
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            //ktra tinh đầy đủ dữ liệu
            if(txtMaHang.Text.Trim() == "")
            {
                MessageBox.Show("Bạn phải nhập mã hàng!");
                txtMaHang.Focus();
                return;
            }
            if (txtTenHang.Text.Trim() == "")
            {
                MessageBox.Show("Bạn phải nhập tên hàng!");
                txtMaHang.Focus();
                return;
            }
            if (cboChatLieu.Text.Trim() == "")
            {
                MessageBox.Show("Bạn phải chọn chất liệu!");
                txtMaHang.Focus();
                return;
            }
            if (txtSoLuong.Text.Trim() == "")
            {
                MessageBox.Show("Bạn phải nhập số lượng!");
                txtMaHang.Focus();
                return;
            }
            if (txtDonGiaNhap.Text.Trim() == "")
            {
                MessageBox.Show("Bạn phải nhập đơn giá nhập!");
                txtMaHang.Focus();
                return;
            }
            if (txtDonGiaBan.Text.Trim() == "")
            {
                MessageBox.Show("Bạn phải nhập đơn giá bán!");
                txtMaHang.Focus();
                return;
            }
            // Thực hiện thêm mới và cập nhật lại dataGridView
            if (btnThem.Enabled == true)
            {
                // Kiểm tra trùng mã
                DataTable dtSP = dt.DocBang("Select * from tblHang where MaHang='" + txtMaHang.Text + "'");
                if (dtSP.Rows.Count > 0)
                {
                    MessageBox.Show("Mã sản phẩm đã tồn tại. Mời bạn nhập mã khác!!!");
                    txtMaHang.Focus();
                    return;
                }

                string strInsert = "Insert into tblHang values('" + txtMaHang.Text + "',N'" + txtTenHang.Text + "','" +
                    cboChatLieu.SelectedValue + "'," + int.Parse(txtSoLuong.Text) +
                    "," + float.Parse(txtDonGiaNhap.Text) + "," + float.Parse(txtDonGiaBan.Text) +
                    ",'" + fileAnh + "',N'" + txtGhiChu.Text + "')";
                dt.CapNhat(strInsert);
            }

            // Thực hiện Sửa dữ liệu
            if (btnSua.Enabled == true)
            {
                string sqlUpdate = "update tblHang set TenHang=N'" + txtTenHang.Text +
                    "',MaChatLieu='" + cboChatLieu.SelectedValue + "',SoLuong="
                    + Convert.ToInt16(txtSoLuong.Text) + ",DonGiaNhap=" + float.Parse(txtDonGiaNhap.Text) +
                    ",DonGiaBan=" + float.Parse(txtDonGiaBan.Text) + ",Anh='" + fileAnh +
                    "',GhiChu=N'" + txtGhiChu.Text + "' where MaHang='" + txtMaHang.Text + "'";
                 dt.CapNhat(sqlUpdate);
            }

            // Load dữ liệu lên dgvHang
            DataTable dtHang = dt.DocBang("select * from tblHang");
            dgvHang.DataSource = dtHang;

            // Đưa form về trạng thái ban đầu
            // Cấm click các nút Sửa, xóa, lưu và bỏ qua. Click được nút Thêm mới
            btnSua.Enabled = false;
            btnXoa.Enabled = false;
            btnLuu.Enabled = false;
            btnBoQua.Enabled = false;
            btnThem.Enabled = true;

            // Xóa trắng các dữ liệu trên các ô nhập liệu
            ResetValue();
        }

        private void dgvHang_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            btnXoa.Enabled = true;
            btnSua.Enabled = true;
            btnBoQua.Enabled = true;
            btnThem.Enabled = false;
            btnLuu.Enabled = false;

            // Hiển thị dữ liệu chi tiết
            txtMaHang.Text = dgvHang.CurrentRow.Cells[0].Value.ToString();
            txtTenHang.Text = dgvHang.CurrentRow.Cells[1].Value.ToString();
            txtSoLuong.Text = dgvHang.CurrentRow.Cells[3].Value.ToString();
            cboChatLieu.SelectedValue = dgvHang.CurrentRow.Cells[2].Value.ToString();
            txtDonGiaBan.Text = dgvHang.CurrentRow.Cells[5].Value.ToString();
            txtDonGiaNhap.Text = dgvHang.CurrentRow.Cells[4].Value.ToString();

            try
            {
                picAnh.Image = Image.FromFile("Images\\Hang\\" + dgvHang.CurrentRow.Cells[6].Value.ToString());
            }
            catch
            {
                // Xử lý trường hợp không tìm thấy file ảnh hoặc tên ảnh bị rỗng
                picAnh.Image = null;
            }
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            btnThem.Enabled = false;
            btnLuu.Enabled = true;
            btnXoa.Enabled = false;
            btnBoQua.Enabled = true;
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn xóa không?", "Xóa mặt hàng",
                 MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                dt.CapNhat("Delete tblHang where MaHang='" + txtMaHang.Text + "'");
                dgvHang.DataSource = dt.DocBang("select * from tblHang");
            }

            // Đưa form về trạng thái ban đầu
            // Cấm click các nút Sửa, xóa, lưu và bỏ qua. Click được nút Thêm mới
            btnSua.Enabled = false;
            btnXoa.Enabled = false;
            btnLuu.Enabled = false;
            btnBoQua.Enabled = false;
            btnThem.Enabled = true;
        }

        private void btnXuatfile_Click(object sender, EventArgs e)
        {
            //khai báo và khởi tạo các thành phần, đối tượng
            Excel.Application exApp = new Excel.Application();
            Excel.Workbook exBook = exApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet exSheet = (Excel.Worksheet)exBook.Worksheets[1];
            Excel.Range TenTruong = (Excel.Range)exSheet.Cells[1, 1];
            //định dạng file
            TenTruong.Range["B2"].Font.Size = 25;
            TenTruong.Range["B2"].Font.Name = "Times New Roman";
            TenTruong.Range["B2"].Font.Color = Color.Red;
            TenTruong.Range["B2"].Value = "Danh sách sản phẩm";

            TenTruong.Range["A4:F4"].Font.Size = 13;
            TenTruong.Range["A4:F4"].Font.Name = "Times New Roman";
            TenTruong.Range["A4:F4"].Font.Color = Color.Black;
            TenTruong.Range["A4:F4"].Font.Bold = true;

            TenTruong.Range["A4"].Value = "Mã hàng";
            TenTruong.Range["B4"].Value = "Tên hàng";
            TenTruong.Range["C4"].Value = "Chất liệu";
            TenTruong.Range["D4"].Value = "Số lượng";
            TenTruong.Range["E4"].Value = "Đơn giá nhập";
            TenTruong.Range["F4"].Value = "Đơn giá bán";

            int hang = 5;
            for(int i = 0; i <dgvHang.Rows.Count-1; i++)
            {
                TenTruong.Range["A" + hang.ToString()].Value = dgvHang.Rows[i].Cells[0].Value.ToString();
                TenTruong.Range["B" + hang.ToString()].Value = dgvHang.Rows[i].Cells[1].Value.ToString();
                TenTruong.Range["C" + hang.ToString()].Value = dgvHang.Rows[i].Cells[2].Value.ToString();
                TenTruong.Range["D" + hang.ToString()].Value = dgvHang.Rows[i].Cells[3].Value.ToString();
                TenTruong.Range["E" + hang.ToString()].Value = dgvHang.Rows[i].Cells[4].Value.ToString();
                TenTruong.Range["F" + hang.ToString()].Value = dgvHang.Rows[i].Cells[5].Value.ToString();
                hang++;
            }
            exSheet.Name = "DSSanPham";
            // kích hoạt file Excel
            exBook.Activate();
            SaveFileDialog dlFile = new SaveFileDialog();
            if (dlFile.ShowDialog() == DialogResult.OK)
                exBook.SaveAs(dlFile.FileName.ToString());
            exApp.Quit();
        }
    }
}
        

