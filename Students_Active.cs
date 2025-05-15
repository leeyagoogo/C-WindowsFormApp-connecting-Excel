using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.AxHost;

namespace Loop
{
    public partial class Students_Active : Form
    {
        public Students_Active()
        {
            InitializeComponent();
            //LoadExcelFile();
            showActive(dataGridViewActive);

        }

        public void showActive(DataGridView v)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx");
            Worksheet sh = book.Worksheets[0];
            int row = sh.Rows.Length;
            DataTable dt = sh.ExportDataTable();

            dt.Columns.Add("ExcelRow", typeof(int));
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dt.Rows[i]["ExcelRow"] = i + 2; 
            }

            LinkedList<DataRow> rowsToDelete = new LinkedList<DataRow>();

            for (int i = 2; i <= row; i++)
            {
                if (sh.Range[i, 13].Value != "1") // Column 13 = Status
                {
                    int rowIndex = i - 2;

                    if (rowIndex >= 0 && rowIndex < dt.Rows.Count)
                    {
                        rowsToDelete.AddFirst(dt.Rows[rowIndex]);
                    }
                }
            }

            foreach (DataRow r in rowsToDelete)
            {
                dt.Rows.Remove(r);
            }

            dt.AcceptChanges(); 

            v.DataSource = dt;
            v.Columns["ExcelRow"].Visible = false; // Hide tracking column

        }

        //public void LoadExcelFile()
        //{
        //    Workbook book = new Workbook();
        //    book.LoadFromFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx");
        //    Worksheet sh = book.Worksheets[0];
        //    DataTable dt = sh.ExportDataTable();
        //    dataGridView.DataSource = dt;
        //}

        public void insertData(string n, DateTime bdate, string a, string e, string u, string pa, string g, string h, string col, string cou, string s, string p)
        { 
            //int i = dataGridView.Rows.Add();
            
            //dataGridView.Rows[i].Cells[0].Value = n;
            //dataGridView.Rows[i].Cells[1].Value = g;
            //dataGridView.Rows[i].Cells[2].Value = h;
            //dataGridView.Rows[i].Cells[3].Value = c;
            //dataGridView.Rows[i].Cells[4].Value = s;

        }

        public void update(int id, string n, DateTime bdate, string a, string e, string u, string pa, string g, string h, string col, string cou, string s, string p)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx");
            Worksheet sh = book.Worksheets[0];

            sh.Range[id + 2, 1].Value = n;
            sh.Range[id + 2, 2].Value = bdate.ToString("MM/dd/yyyy");
            sh.Range[id + 2, 3].Value = a;
            sh.Range[id + 2, 4].Value = e;
            sh.Range[id + 2, 5].Value = u;
            sh.Range[id + 2, 6].Value = pa;
            sh.Range[id + 2, 7].Value = g;
            sh.Range[id + 2, 8].Value = h;
            sh.Range[id + 2, 9].Value = col;
            sh.Range[id + 2, 10].Value = cou;
            sh.Range[id + 2, 11].Value = s;
            sh.Range[id + 2, 12].Value = p;


            //book.SaveToFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx", ExcelVersion.Version2016);
            //showActive(dataGridViewActive);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx");
            Worksheet sh = book.Worksheets[0];
            int row = sh.Rows.Length;
            bool deleted = false;

            if (dataGridViewActive.SelectedRows.Count > 0)
            {
                string selectedName = dataGridViewActive.SelectedRows[0].Cells[0].Value.ToString();

                var delete = MessageBox.Show("Are you sure you want to move this information?", "Active to Inactive", MessageBoxButtons.YesNo);
                if (delete == DialogResult.Yes)
                {
                    for (int i = 2; i <= row; i++)
                    {

                        if (sh.Range[i, 1].Value == selectedName && sh.Range[i, 13].Value == "1")
                        {
                            sh.Range[i, 13].Text = "0";
                            deleted = true;
                            break;
                        }
                    }

                }

                if (deleted)
                {
                    book.SaveToFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx", ExcelVersion.Version2016);

                    MyLogs myLog = new MyLogs();
                    myLog.insertLogs(User.CurrentUser, "Deleting Student Information (0).");

                    showActive(dataGridViewActive);


                    foreach (Form openForm in Application.OpenForms)
                    {
                        if (openForm is Students_Inactive inactiveForm)
                        {
                            inactiveForm.showInactive(inactiveForm.dataGridViewInactive);
                            break;
                        }
                    }
                }

                else
                {
                    MessageBox.Show("Action cancelled.", "Cancelled");
                }
            }

            else
            {
                MessageBox.Show("Please select a student first!", "Warning");
            }
        }
           

        private void dataGridView_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //int r = dataGridViewActive.CurrentCell.RowIndex;

            
            //Form1 f1 = (Form1)Application.OpenForms["Form1"];
            Form1 f1 = new Form1();

            int r = dataGridViewActive.CurrentCell.RowIndex;
            int excelRow = Convert.ToInt32(dataGridViewActive.Rows[r].Cells["ExcelRow"].Value);
            f1.lblID.Text = (excelRow - 2).ToString();

            //f1.lblID.Text = r.ToString();
            f1.txtName.Text = dataGridViewActive.Rows[r].Cells[0].Value.ToString();

            string dateStr = dataGridViewActive.Rows[r].Cells[1].Value.ToString();
            DateTime parsedDate;

            if (DateTime.TryParse(dateStr, out parsedDate))
            {
                f1.dateTimePickerBirth.Value = parsedDate;
            }

            else
            {
            }

            f1.lblAge.Text = dataGridViewActive.Rows[r].Cells[2].Value.ToString();
            f1.txtEmail.Text = dataGridViewActive.Rows[r].Cells[3].Value.ToString();
            f1.txtUUsername.Text = dataGridViewActive.Rows[r].Cells[4].Value.ToString();
            f1.txtPPassword.Text = dataGridViewActive.Rows[r].Cells[5].Value.ToString();

            string gender = dataGridViewActive.Rows[r].Cells[6].Value.ToString();
            if (gender == "Male")
            {
                f1.rdoM.Checked = true;
            }

            else
            {
                f1.rdoF.Checked = true;
            }

            string hobbies = dataGridViewActive.Rows[r].Cells[7].Value.ToString();
            string[]h = hobbies.Split(' ');

            foreach(string value in h)
            {
                if (value == "Basketball")
                {
                    f1.chkBasketball.Checked = true;
                }

                else if(value == "Volleyball")
                {
                    f1.chkVolleyball.Checked = true;
                }

                else if(value == "Soccer")
                {
                    f1.chkSoccer.Checked = true;
                }
            }

            f1.cmbColor.Text = dataGridViewActive.Rows[r].Cells[8].Value.ToString();
            f1.cmbCourse.Text = dataGridViewActive.Rows[r].Cells[9].Value.ToString();
            f1.txtS.Text = dataGridViewActive.Rows[r].Cells[10].Value.ToString();
            f1.txtPicture.Text = dataGridViewActive.Rows[r].Cells[11].Value.ToString();
            
            f1.btnAdd.Visible = false;
            f1.btnUpdate.Visible = true;
            f1.Show();
            this.Hide();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            dataGridViewActive.ClearSelection();

            try
            {
                foreach (DataGridViewRow row in dataGridViewActive.Rows)
                {
                    if (row.Cells[0].Value.ToString().Equals(txtSearch.Text))
                    {
                        row.Selected = true;
                        break;
                    }
                    else
                    { 
                    }
                }

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnAddInfo_Click(object sender, EventArgs e)
        {
            Form1 f1 = new Form1();
            f1.Show();
            this.Hide();
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            Dashboard dash = new Dashboard();
            dash.Show();
            this.Hide();
        }
    }
}
