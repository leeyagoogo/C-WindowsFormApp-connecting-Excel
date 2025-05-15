using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Loop
{
    public partial class Logs: Form
    {
        public Logs()
        {
            InitializeComponent();
            showLogs(dataGridViewlog);
        }

        public void showLogs(DataGridView v)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx");
            Worksheet sh = book.Worksheets[1];
            DataTable dt = sh.ExportDataTable();
            v.DataSource = dt;
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            Dashboard dash = new Dashboard();
            dash.Show();
            this.Hide();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string searchText = txtSearch.Text.Trim();
            dataGridViewlog.ClearSelection();

            try
            {
                foreach (DataGridViewRow row in dataGridViewlog.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[0].Value.ToString().Equals(searchText, StringComparison.OrdinalIgnoreCase))
                    {
                        row.Selected = true;
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
    }
}

//MGA PATH NI ANTE KAY CGEG ILIS 

// in school path (@"C:\Users\ACT-STUDENT\Desktop\Lerio_ForLoop\Book1.xlsx");
// in home path (@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx");

//for dp
//(if nasa balay): C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\dp.jpg
//(if naa sa skol): C:\Users\ACT-STUDENT\Desktop\Lerio_ForLoop\dp.jpg