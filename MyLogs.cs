using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Loop
{
    class MyLogs
    {
        public void insertLogs (string u, string m)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx");
            Worksheet sh = book.Worksheets[1];
            int row = sh.Rows.Length + 1;

            sh.Range[row, 1].Value = u;
            sh.Range[row, 2].Value = m;
            sh.Range[row, 3].Value = DateTime.Now.ToString("MM/dd/yyyy");
            sh.Range[row, 4].Value = DateTime.Now.ToString("hh : mm : ss : tt");


            book.SaveToFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx");
        }
    }
}
