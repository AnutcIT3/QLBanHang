using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QLBanHang.Classes
{
    internal class Function{
        public void FillComBox(ComboBox cb,DataTable dt, string display, string value){
            cb.DataSource = dt;
            cb.DisplayMember  = display;
            cb.ValueMember = value;
        }
    }
}
