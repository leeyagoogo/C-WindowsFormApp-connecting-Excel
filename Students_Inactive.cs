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

namespace Loop
{
    public partial class Students_Inactive : Form
    {
        public Students_Inactive()
        {
            InitializeComponent();
            showInactive(dataGridViewInactive);

        }

        public void showInactive(DataGridView v)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx");
            Worksheet sh = book.Worksheets[0];
            int row = sh.Rows.Length;
            DataTable dt = sh.ExportDataTable();

            List<DataRow> rowsToDelete = new List<DataRow>();

            for (int i = 2; i <= row; i++)
            {
                if (sh.Range[i, 13].Value != "0")
                {
                    int rowindex = i - 2;

                    if (rowindex >= 0 && rowindex < dt.Rows.Count)
                    {
                        rowsToDelete.Add(dt.Rows[rowindex]);
                    }
                }
            }

            foreach (DataRow Rows in rowsToDelete)
            {
                dt.Rows.Remove(Rows);
            }

            v.DataSource = dt;
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            dataGridViewInactive.ClearSelection();

            try
            {
                foreach (DataGridViewRow row in dataGridViewInactive.Rows)
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

        private void btnDelete_Click(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx");
            Worksheet sh = book.Worksheets[0];
            int row = sh.Rows.Length;
            bool deleted = false;

            if (dataGridViewInactive.SelectedRows.Count > 0)
            {
                string selectedName = dataGridViewInactive.SelectedRows[0].Cells[0].Value.ToString();

                var delete = MessageBox.Show("Are you sure you want to move this information?", "Inactive to Active", MessageBoxButtons.YesNo);
                if (delete == DialogResult.Yes)
                {
                    for (int i = 2; i <= row; i++)
                    {
                        if (sh.Range[i, 1].Value == selectedName && sh.Range[i, 13].Value == "0")
                        {
                            sh.Range[i, 13].Text = "1";
                            deleted = true;

                            break;

                        }
                    }
                }

                if (deleted)
                {
                    book.SaveToFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx", ExcelVersion.Version2016);

                    MyLogs myLog = new MyLogs();
                    myLog.insertLogs(User.CurrentUser, "Deleting Student Information (1).");

                    showInactive(dataGridViewInactive);


                    foreach (Form openForm in Application.OpenForms)
                    {
                        if (openForm is Students_Active activeForm)
                        {
                            activeForm.showActive(activeForm.dataGridViewActive);

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

        private void btnBack_Click(object sender, EventArgs e)
        {
            Dashboard dash = new Dashboard();
            dash.Show();
            this.Hide();
        }
    }
}
