using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TVMH.appCodes;

namespace TVMH
{
    public partial class frmDN : Form
    {
        hashCode hc = new hashCode();
       
        public frmDN()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("Bạn chắc muốn thoát không ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                Application.Exit();
            }
        }

        private void txtUserName_Enter(object sender, EventArgs e)
        {
            if(txtUserName.Text== "Tên tài khoản")
            {
                txtUserName.Text = "";
                txtUserName.ForeColor = Color.DimGray;
            }
        }

        private void txtUserName_Leave(object sender, EventArgs e)
        {
            if (txtUserName.Text == "")
            {
                txtUserName.Text = "Tên tài khoản";
                txtUserName.ForeColor = Color.DimGray;
            }
        }

        private void txtPassword_Enter(object sender, EventArgs e)
        {
            if(txtPassword.Text== "Mật khẩu")
            {
                txtPassword.Text = "";
                txtPassword.ForeColor = Color.DimGray;
                txtPassword.UseSystemPasswordChar = true;
            }
        }

        private void txtPassword_Leave(object sender, EventArgs e)
        {
            if (txtPassword.Text == "")
            {
                txtPassword.Text = "Mật khẩu";
                txtPassword.ForeColor = Color.DimGray;
                txtPassword.UseSystemPasswordChar = false;
            }
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            
            if (txtUserName.Text=="tranthanhsang1998" && (hc.PassHash(txtPassword.Text)== "118718213819283188213881491641864211117111316942133"))
            {
                Form1 frm = new Form1()
                {
                    StartPosition= FormStartPosition.CenterScreen
                };
                frm.Show();
                this.Hide();
            }
            else if (txtUserName.Text == "" || txtUserName.Text == "Tên tài khoản" || txtPassword.Text == "" || txtPassword.Text == "Mật khẩu")
            {
                this.lbError.Text = "Vui lòng nhập đủ các thông tin !";
            }
            else
            {
                this.lbError.Text = "Tài khoản hoặc mật khẩu không đúng !";
            }
            
        }

       
    }
}
