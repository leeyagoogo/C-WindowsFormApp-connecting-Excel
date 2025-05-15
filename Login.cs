using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.IO;

namespace Loop
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
         //   MyLogs log = new MyLogs();
         //   log.insertLogs("User","Message");
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            Dashboard d = new Dashboard();
            book.LoadFromFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx");
            Worksheet sh = book.Worksheets[0];
            int row = sh.Rows.Length;
            bool log = false;

            for (int i = 2; i <= row; i++)
            {
                if (sh.Range[i, 5].Value == txtUsername.Text && sh.Range[i, 6].Value == txtPassword.Text 
                    && sh.Range[i, 13].Value != "0")
                {
                    d.pictureBox1.Image = Image.FromFile(@"" + sh.Range[i, 12].Value);
                    d.pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                    User.CurrentUser = txtUsername.Text;
                    User.UserName = sh.Range[i, 1].Value;
                    User.Profile = sh.Range[i, 12].Value;
  
                    log = true;
                    break;
                }

                else
                {
                    log = false;
                }
            }

            if (log == true)
            {
                MyLogs myLog = new MyLogs();
                myLog.insertLogs(User.CurrentUser, "Logged in.");

                MessageBox.Show("Login Successfully!", "Login Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                d.Show();
                this.Hide();

            }

            else
            { 
                MessageBox.Show("Invalid Username and Password!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void chkShowPassword_CheckedChanged(object sender, EventArgs e)
        {
            if (chkShowPassword.Checked)
            {
                txtPassword.PasswordChar = '\0';
            }

            else
            {
                txtPassword.PasswordChar = '●';
            }
        }

        private void lblExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
