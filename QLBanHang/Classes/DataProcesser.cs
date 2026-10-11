using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QLBanHang.Classes
{
    internal class DataProcesser
    {
        string strConnect = "Data Source=.\\SQLEXPRESS;Initial Catalog=QLBanHang;Integrated Security=True;";
        SqlConnection conn = null;

        //Hàm mở kết nối
        private void KetNoiCSDL()
        {
            conn = new SqlConnection(strConnect);
            if (conn.State != ConnectionState.Open)
                conn.Open();
        }

        private void DongKetNoi()
        {
            if(conn.State != ConnectionState.Closed)
                conn.Close();
            conn.Dispose();
        }

        //Thực thi câu lệnh select 
        public DataTable DocBang(string sql)
        {
            DataTable dt = new DataTable();
            KetNoiCSDL();
            SqlDataAdapter sqldataAdapter = new SqlDataAdapter(sql,conn);
            sqldataAdapter.Fill(dt);
            DongKetNoi();
            return dt;
        }

        public void CapNhat(string sql){
            KetNoiCSDL();
            SqlCommand sqlcomm = new SqlCommand();
            sqlcomm.Connection = conn;
            sqlcomm.CommandText = sql;
            sqlcomm.ExecuteNonQuery();
            DongKetNoi();
        }


    }
}
