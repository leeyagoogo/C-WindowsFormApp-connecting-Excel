using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Xml.Linq;

namespace Loop
{
    public partial class Dashboard: Form
    {
        public Dashboard()
        {
            InitializeComponent();
            //pictureBox1.Image = Image.FromFile("C:\\Users\\ACT-STUDENT\\Desktop\\Lerio_ForLoop\\dp.jpg");
            if (!string.IsNullOrEmpty(User.Profile) && File.Exists(User.Profile))
            {
                pictureBox1.Image = Image.FromFile(User.Profile);
            }

            else
            {
                pictureBox1.Image = Image.FromFile(@"C:\\Users\\Arlah\\OneDrive\\Desktop\\Lerio_ForLoop\\nodp.jpg");
            }
        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            var lout = MessageBox.Show("Are you sure you want to log out?", "Sign out", MessageBoxButtons.YesNo);
            if (lout == DialogResult.Yes)
            {
                Workbook book = new Workbook();
                book.LoadFromFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx");
                Worksheet sh = book.Worksheets[0];
                int row = sh.Rows.Length;

                Login log = new Login();
                log.Show();

                MyLogs myLog = new MyLogs();
                myLog.insertLogs(User.CurrentUser, "Logged out.");

                this.Hide();
            }
        }

        private void btnLogs_Click(object sender, EventArgs e)
        {
            Logs logs = new Logs();
            logs.Show();
            this.Hide();
        }

        private void btnActiveStudents_Click(object sender, EventArgs e)
        {
            Students_Active active = new Students_Active();
            active.Show();
            this.Hide();
        
        }

        private void btnInactiveStudents_Click(object sender, EventArgs e)
        {
            Students_Inactive inactive = new Students_Inactive();
            inactive.Show();
            this.Hide();
        
        }

        public int showCount(int c, string val)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\Arlah\OneDrive\Desktop\Lerio_ForLoop\Book1.xlsx");
            Worksheet sh = book.Worksheets[0];
            int row = sh.Rows.Length;
            int counter = 0;

            for (int i = 2; i <= row; i++)
            {
                if (sh.Range[i, c].Value.ToString().Contains(val))
                {
                    counter++;
                }

            }
            return counter;

        }

        private void ChartFavoriteColor()
        {
            int redCount = showCount(9, "Red");
            int blueCount = showCount(9, "Blue");
            int yellowCount = showCount(9, "Yellow");

            chart1.Series.Clear();

            var series = new System.Windows.Forms.DataVisualization.Charting.Series("s1");
            series.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;

            series.Points.AddXY("Red", redCount);
            series.Points.AddXY("Blue", blueCount);
            series.Points.AddXY("Yellow", yellowCount);

            series.Points[0].Color = System.Drawing.Color.Red;  
            series.Points[1].Color = System.Drawing.Color.Blue;  
            series.Points[2].Color = System.Drawing.Color.Yellow;

            series.Points[0].LegendText = "Red (" + redCount + ")";
            series.Points[1].LegendText = "Blue (" + blueCount + ")";
            series.Points[2].LegendText = "Yellow (" + yellowCount + ")";

            chart1.ChartAreas[0].Position.Width = 80;
            chart1.ChartAreas[0].Position.Height = 80;
            chart1.ChartAreas[0].Position.X = 1;
            chart1.ChartAreas[0].Position.Y = 10;

            chart1.Series.Add(series);
        }

        private void Programs()
        {
            int bsitCount = showCount(10, "BSIT");
            int bscpeCount = showCount(10, "BSCpE");
            int bscsCount = showCount(10, "BSCS");
            

            chart2.Series.Clear();

            var series = new System.Windows.Forms.DataVisualization.Charting.Series("Programs");
            series.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Bar;

            series.Points.AddXY("BSIT", bsitCount);
            series.Points.AddXY("BSCpE", bscpeCount);
            series.Points.AddXY("BSCS", bscsCount);

            series.Points[0].Color = System.Drawing.Color.Blue;
            series.Points[1].Color = System.Drawing.Color.Green;
            series.Points[2].Color = System.Drawing.Color.Orange;

            series.Points[0].LegendText = "BSIT (" + bsitCount + ")";
            series.Points[1].LegendText = "BSCpE (" + bscpeCount + ")";
            series.Points[2].LegendText = "BSCS (" + bscsCount + ")";

            chart2.Series.Add(series);

            //chart2.Titles.Clear();
            //chart2.Titles.Add("Program Enrollment Distribution");
            chart2.Legends.Clear();
        }

        private void Dashboard_Load(object sender, EventArgs e)
        {
            
            lblName.Text = User.UserName;
            ChartFavoriteColor();
            Programs();

            lblGenderCountM.Text = showCount(7, "Male").ToString();
            lblGenderCountF.Text = showCount(7, "Female").ToString();

            lblHCountV.Text = showCount(8, "Volleyball").ToString();
            lblHCountB.Text = showCount(8, "Basketball").ToString();
            lblHCountS.Text = showCount(8, "Soccer").ToString();

            //lblColorCountRed.Text = showCount(9, "Red").ToString();
            //lblColorCountBlue.Text = showCount(9, "Blue").ToString();
            //lblColorCountYellow.Text = showCount(9, "Yellow").ToString();

            //lblBSITCount.Text = showCount(10, "BSIT").ToString();
            //lblBSCpECount.Text = showCount(10, "BSCpE").ToString();
            //lblBSCSCount.Text = showCount(10, "BSCS").ToString();

            lblActiveCount.Text = showCount(13, "1").ToString();
            lblInactiveCount.Text = showCount(13, "0").ToString();

            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            TimeZoneInfo phTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Singapore Standard Time");
            DateTime phTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, phTimeZone);

            lblTime.Text = phTime.ToString("MMMM dd, yyyy\nhh:mm:ss tt");
        }
    }
}
