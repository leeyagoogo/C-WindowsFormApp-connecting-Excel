using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Loop
{
    public partial class Form1 : Form
    {
        string[] student = new string[5];
        int i = 0;
        Students_Active f2 = new Students_Active();
        Students_Active active = new Students_Active();
        private Dictionary<string, string> fieldLabels;

        public Form1()
        {
            InitializeComponent();
            fieldLabels = new Dictionary<string, string>
            {
                { "txtName", "Name" },
                { "txtEmail", "Email" },
                { "txtUUsername", "Username" },
                { "txtPPassword", "Password" },
                { "cmbColor", "Favorite Color" },
                { "cmbCourse", "Course" },
                { "txtS", "Saying" },
                { "txtPicture", "Profile Picture" }
            };

            dateTimePickerBirth.Format = DateTimePickerFormat.Custom;
            dateTimePickerBirth.CustomFormat = "MM/dd/yyyy";
        }

        public bool ValidateInputs(bool Update = false)
        {
            string error = "";

            foreach (Control c in Controls)
            {
                if (c is TextBox || c is ComboBox)
                {
                    if (string.IsNullOrWhiteSpace(c.Text))
                    {
                        if (fieldLabels.ContainsKey(c.Name))
                            error += fieldLabels[c.Name] + " is empty.\n";
                        else
                            error += c.Name + " is empty.\n";
                    }
                }
            }


            if (!IsValidEmail(txtEmail.Text))
            {
                error += "Invalid Email Format.\n";
                MessageBox.Show("Invalid Email Format.", "Error");
            }

            if (dateTimePickerBirth.Value.Date == DateTime.Now.Date)
            {
                error += "Birthdate is not properly set.\n";
            }

            if (!rdoM.Checked && !rdoF.Checked)
            {
                error += "Gender is not selected.\n";
            }

            if (!chkBasketball.Checked && !chkVolleyball.Checked && !chkSoccer.Checked)
            {
                error += "At least one hobby must be selected.\n";
            }

            if (!Update)
            {
                Workbook book = new Workbook();
                book.LoadFromFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx");
                Worksheet sh = book.Worksheets[0];

                bool usernameExists = false;
                bool passwordExists = false;

                for (int row = 2; row <= sh.Rows.Length; row++)
                {
                    string existingUsername = sh.Range[row, 5].Text;
                    string existingPassword = sh.Range[row, 6].Text;

                    if (existingUsername == txtUUsername.Text)
                        usernameExists = true;

                    if (existingPassword == txtPPassword.Text)
                        passwordExists = true;

                }

                if (usernameExists)
                {
                    error += "Username already exists.\n";
                    MessageBox.Show("Username already exists.", "Error");
                }

                if (passwordExists)
                {
                    error += "Password already exists.\n";
                    MessageBox.Show("Password already exists.", "Error");
                }
            }

            if (error != "")
            {
                lblErrors.Text = error;
                return false;
            }

            lblErrors.Text = "—";
            return true;
        }

        public bool IsValidEmail(string email)
        {
            string pattern = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$";
            Regex regex = new Regex(pattern);
            return regex.IsMatch(email);
        }


        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (!ValidateInputs())
            {
                return;
            }

            var data = "";
            string name = txtName.Text;
            DateTime bdate = dateTimePickerBirth.Value;
            string age = lblAge.Text;
            string email = txtEmail.Text;
            string username = txtUUsername.Text;
            string password = txtPPassword.Text;

                string gender = "";
                if (rdoM.Checked)
                {
                    gender = rdoM.Text;
                }

                else
                {
                    gender = rdoF.Text;
                }

                string hobby = "";
                if (chkBasketball.Checked)
                {
                    hobby += chkBasketball.Text + " ";

                }

                if (chkVolleyball.Checked)
                {
                    hobby += chkVolleyball.Text + " ";
                }

                if (chkSoccer.Checked)
                {
                    hobby += chkSoccer.Text + " ";

                }

                string color = cmbColor.Text;
                string course = cmbCourse.Text;
                string saying = txtS.Text;
                string profile = txtPicture.Text;

                data += name + ";";
                data += bdate.ToString("MM/dd/yyyy") + ";";
                data += age + ";";
                data += email + ";";
                data += username + ";";
                data += password + ";";
                data += gender + ";";
                data += hobby + ";";
                data += color + ";";
                data += course + ";";
                data += saying + ";";
                data += profile + ";";

                student[i] = data;
                i++;

                f2.insertData(name, bdate, age, email, username, password, gender, hobby, color, course, saying, profile);
                Workbook book = new Workbook();
                book.LoadFromFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx");
                Worksheet sh = book.Worksheets[0];

                 //int r = sh.Rows.Length + 1;
                 int r = 2; 
                 while (!string.IsNullOrEmpty(sh.Range[r, 1].Value))
                 {
                     r++;
                 }

                sh.Range[r, 1].Value = name;
                sh.Range[r, 2].Value = bdate.ToString("MM/dd/yyyy");
                sh.Range[r, 3].Value = age;
                sh.Range[r, 4].Value = email;
                sh.Range[r, 5].Value = username;
                sh.Range[r, 6].Value = password;
                sh.Range[r, 7].Value = gender;
                sh.Range[r, 8].Value = hobby;
                sh.Range[r, 9].Value = color;
                sh.Range[r, 10].Value = course;
                sh.Range[r, 11].Value = saying;
                sh.Range[r, 12].Value = profile;
                sh.Range[r, 13].Value = "1";

                book.SaveToFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx", ExcelVersion.Version2016);
                MessageBox.Show("Information Added!", "Information Added Successfully", MessageBoxButtons.OK, MessageBoxIcon.Information);
             
                DataTable dt = sh.ExportDataTable();

                f2.dataGridViewActive.DataSource = dt;

                MyLogs myLog = new MyLogs();
                myLog.insertLogs(User.CurrentUser, "Adding New Student Information.");

                txtName.Clear();
                dateTimePickerBirth.Value = DateTime.Now;
                lblAge.Text = "—";
                txtEmail.Clear();
                txtUUsername.Clear();
                txtPPassword.Clear();
                rdoM.Checked = false;
                rdoF.Checked = false;
                chkBasketball.Checked = false;
                chkVolleyball.Checked = false;
                chkSoccer.Checked = false;
                cmbColor.SelectedIndex = -1;
                cmbCourse.SelectedIndex = -1;
                txtS.Clear();
                txtPicture.Clear();
            
        }

        private void btnDisplay_Click(object sender, EventArgs e)
        {
            string val = "";
            for (int j = 0; j < student.Length; j++)
            {
                val += "[" + j + "] = " + student[j] + "\n";
                active = new Students_Active();

            }

            MessageBox.Show(val);
            active.Show();
            this.Hide();

        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (!ValidateInputs(true))
            {
                return;
            }

            var data = "";
            string name = txtName.Text;
            DateTime bdate = dateTimePickerBirth.Value;
            string age = lblAge.Text;
            string email = txtEmail.Text;
            string username = txtUUsername.Text;
            string password = txtPPassword.Text;

            string gender = "";
            if (rdoM.Checked)
            {
                gender = rdoM.Text;
            }

            else
            {
                gender = rdoF.Text;
            }

            string hobby = "";
            if (chkBasketball.Checked)
            {
                hobby += chkBasketball.Text + " ";

            }

            if (chkVolleyball.Checked)
            {
                hobby += chkVolleyball.Text + " ";
            }

            if (chkSoccer.Checked)
            {
                hobby += chkSoccer.Text + " ";

            }

            string color = cmbColor.Text;
            string course = cmbCourse.Text;
            string saying = txtS.Text;
            string profile = txtPicture.Text;

            data += name + ";";
            data += bdate.ToString("MM/dd/yyyy") + ";";
            data += age + ";";
            data += email + ";";
            data += username + ";";
            data += password + ";";
            data += gender + ";";
            data += hobby + ";";
            data += color + ";";
            data += course + ";";
            data += saying + ";";
            data += profile + ";";

            student[i] = data;
            i++;


            f2.update(Convert.ToInt32(lblID.Text), name, bdate, age, email, username, password, gender, hobby, color, course, saying, profile);
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx");
            Worksheet sh = book.Worksheets[0];

            int r = Convert.ToInt32(lblID.Text) + 2;
            if (sh.Range[r, 13].Value != "1")
            {
                MessageBox.Show("Cannot update inactive record.");
                return;
            }

            sh.Range[r, 1].Value = name;
            sh.Range[r, 2].Value = bdate.ToString("MM/dd/yyyy");
            sh.Range[r, 3].Value = age;
            sh.Range[r, 4].Value = email;
            sh.Range[r, 5].Value = username;
            sh.Range[r, 6].Value = password;
            sh.Range[r, 7].Value = gender;
            sh.Range[r, 8].Value = hobby;
            sh.Range[r, 9].Value = color;
            sh.Range[r, 10].Value = course;
            sh.Range[r, 11].Value = saying;
            sh.Range[r, 12].Value = profile;
            sh.Range[r, 13].Value = "1";

            book.SaveToFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx", ExcelVersion.Version2016);
            MessageBox.Show("Updated Successfully!", "Information Updated Successfully", MessageBoxButtons.OK, MessageBoxIcon.Information);

            DataTable dt = sh.ExportDataTable();

            f2.dataGridViewActive.DataSource = dt;

            MyLogs myLog = new MyLogs();
            myLog.insertLogs(User.CurrentUser, "Updating Student Information.");
            //active.update(Convert.ToInt32(lblID.Text), name, bdate, age, email, username, password, gender, hobby, color, course, saying, profile);

            txtName.Clear();
            dateTimePickerBirth.Value = DateTime.Now;
            lblAge.Text = "—";
            txtEmail.Clear();
            txtUUsername.Clear();
            txtPPassword.Clear();
            rdoM.Checked = false;
            rdoF.Checked = false;
            chkBasketball.Checked = false;
            chkVolleyball.Checked = false;
            chkSoccer.Checked = false;
            cmbColor.SelectedIndex = -1;
            cmbCourse.SelectedIndex = -1;
            txtS.Clear();
            txtPicture.Clear();
            btnAdd.Visible = true;
            btnUpdate.Visible = false;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();

            if (d.ShowDialog() == DialogResult.OK)
            {
                txtPicture.Text = d.FileName;
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            Dashboard dashboard = new Dashboard();
            dashboard.Show();
            this.Hide();
        }

        private void dateTimePickerBirth_ValueChanged(object sender, EventArgs e)
        {
            int age;

            DateTime birthdate = dateTimePickerBirth.Value;
            DateTime today = DateTime.Today;

            age = today.Year - birthdate.Year;

            if (birthdate.Date > today.AddYears(-age))
            {
                age--;
            }

            lblAge.Text = age.ToString();
        }

        private void chkShowPassword_CheckedChanged(object sender, EventArgs e)
        {
            if (chkShowPassword.Checked)
            {
                txtPPassword.PasswordChar = '\0';
            }

            else
            {
                txtPPassword.PasswordChar = '●';
            }
        }
    }
}
