using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ComponentFactory.Krypton.Toolkit;
using Stream_25percent;

namespace Stream
{
    public partial class Beam_Schedule : KryptonForm
    {
        public Beam_Schedule()
        {
            InitializeComponent();
        }

        private void Beam_Schedule_Load(object sender, EventArgs e)
        {

        }

        private void Data_GRD_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void filenamepdf_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Files Only | *.xlsx; *.xls";
                ofd.Title = "Choose the file";
                if (ofd.ShowDialog() == DialogResult.OK)
                    filenameexcel.Text = ofd.FileName;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlapp;
            Microsoft.Office.Interop.Excel.Workbook xlworkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlworksheet;
            Microsoft.Office.Interop.Excel.Range xlrange;
            try
            {
                xlapp = new Microsoft.Office.Interop.Excel.Application();
                xlworkbook = xlapp.Workbooks.Open(filenameexcel.Text);
                xlworksheet = xlworkbook.Worksheets["Sheet1"];
                xlrange = xlworksheet.UsedRange;

                Data_GRD.ColumnCount = xlrange.Columns.Count;

                for (int xlrow = 1; xlrow <= xlrange.Rows.Count; xlrow++)
                {
                    Data_GRD.Rows.Add(xlrange.Cells[xlrow, 1].Text, xlrange.Cells[xlrow, 2].Text, xlrange.Cells[xlrow, 3].Text, xlrange.Cells[xlrow, 4].Text, xlrange.Cells[xlrow, 5].Text, xlrange.Cells[xlrow, 6].Text,
                        xlrange.Cells[xlrow, 7].Text, xlrange.Cells[xlrow, 8].Text, xlrange.Cells[xlrow, 9].Text, xlrange.Cells[xlrow, 10].Text, xlrange.Cells[xlrow, 11].Text, xlrange.Cells[xlrow, 12].Text, xlrange.Cells[xlrow, 13].Text);
                }

                xlworkbook.Close();
                xlapp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button1_MouseHover(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromArgb(30, 27, 46);
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            lblText.Text = "The Beam Schedule from STAAD RCDC";
            Microsoft.Office.Interop.Excel.Application xlapp;
            Microsoft.Office.Interop.Excel.Workbook xlworkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlworksheet;
            Microsoft.Office.Interop.Excel.Range xlrange;
            try
            {
                xlapp = new Microsoft.Office.Interop.Excel.Application();
                xlworkbook = xlapp.Workbooks.Open(filenameexcel.Text);
                xlworksheet = xlworkbook.Worksheets["Sheet1"];
                xlrange = xlworksheet.UsedRange;

                Data_GRD.ColumnCount = xlrange.Columns.Count;

                for (int xlrow = 1; xlrow <= xlrange.Rows.Count; xlrow++)
                {
                    Data_GRD.Rows.Add(xlrange.Cells[xlrow, 1].Text, xlrange.Cells[xlrow, 2].Text, xlrange.Cells[xlrow, 3].Text, xlrange.Cells[xlrow, 4].Text, xlrange.Cells[xlrow, 5].Text, xlrange.Cells[xlrow, 6].Text,
                        xlrange.Cells[xlrow, 7].Text, xlrange.Cells[xlrow, 8].Text, xlrange.Cells[xlrow, 9].Text, xlrange.Cells[xlrow, 10].Text, xlrange.Cells[xlrow, 11].Text, xlrange.Cells[xlrow, 12].Text, xlrange.Cells[xlrow, 13].Text);
                }



                xlworkbook.Close();
                xlapp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            lblText.Text = "Groupings of continuous beams";
            Microsoft.Office.Interop.Excel.Application xlapp;
            Microsoft.Office.Interop.Excel.Workbook xlworkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlworksheet;
            Microsoft.Office.Interop.Excel.Range xlrange;
            try
            {
                xlapp = new Microsoft.Office.Interop.Excel.Application();
                xlworkbook = xlapp.Workbooks.Open(filenameexcel.Text);
                xlworksheet = xlworkbook.Worksheets["Sheet1"];
                xlrange = xlworksheet.UsedRange;

                xlrange.Range["A3:A6"].Cells.Merge();
                xlrange.Range["A3:A6"].Cells.Value = "2";
                xlrange.Range["B3:B6"].Cells.Merge();
                xlrange.Range["C3:C6"].Cells.Merge();
                xlrange.Range["D3:D6"].Cells.Merge();
                xlrange.Range["E3:E6"].Cells.Merge();
                xlrange.Range["F3:F6"].Cells.Merge();
                xlrange.Range["G3:G6"].Cells.Merge();
                xlrange.Range["H3:H6"].Cells.Merge();
                xlrange.Range["I3:I6"].Cells.Merge();
                xlrange.Range["J3:J6"].Cells.Merge();
                xlrange.Range["K3:K6"].Cells.Merge();
                xlrange.Range["L3:L6"].Cells.Merge();


                xlrange.Range["A7:A10"].Cells.Merge();
                xlrange.Range["A7:A10"].Cells.Value = "2";
                xlrange.Range["B7:B10"].Cells.Merge();
                xlrange.Range["B7:B10"].Cells.Value = "B2";
                xlrange.Range["C7:C10"].Cells.Merge();
                xlrange.Range["D7:D10"].Cells.Merge();
                xlrange.Range["E7:E10"].Cells.Merge();
                xlrange.Range["F7:F10"].Cells.Merge();
                xlrange.Range["G7:G10"].Cells.Merge();
                xlrange.Range["H7:H10"].Cells.Merge();
                xlrange.Range["I7:I10"].Cells.Merge();
                xlrange.Range["J7:J10"].Cells.Merge();
                xlrange.Range["K7:K10"].Cells.Merge();
                xlrange.Range["L7:L10"].Cells.Merge();

                xlrange.Range["A11:A14"].Cells.Merge();
                xlrange.Range["A11:A14"].Cells.Value = "2";
                xlrange.Range["B11:B14"].Cells.Merge();
                xlrange.Range["B11:B14"].Cells.Value = "B3";
                xlrange.Range["C11:C14"].Cells.Merge();
                xlrange.Range["D11:D14"].Cells.Merge();
                xlrange.Range["E11:E14"].Cells.Merge();
                xlrange.Range["F11:F14"].Cells.Merge();
                xlrange.Range["G11:G14"].Cells.Merge();
                xlrange.Range["H11:H14"].Cells.Merge();
                xlrange.Range["I11:I14"].Cells.Merge();
                xlrange.Range["J11:J14"].Cells.Merge();
                xlrange.Range["K11:K14"].Cells.Merge();
                xlrange.Range["L11:L14"].Cells.Merge();

                xlrange.Range["A15:A18"].Cells.Merge();
                xlrange.Range["A15:A18"].Cells.Value = "2";
                xlrange.Range["B15:B18"].Cells.Merge();
                xlrange.Range["B15:B18"].Cells.Value = "B4";
                xlrange.Range["C15:C18"].Cells.Merge();
                xlrange.Range["D15:D18"].Cells.Merge();
                xlrange.Range["E15:E18"].Cells.Merge();
                xlrange.Range["F15:F18"].Cells.Merge();
                xlrange.Range["G15:G18"].Cells.Merge();
                xlrange.Range["H15:H18"].Cells.Merge();
                xlrange.Range["I15:I18"].Cells.Merge();
                xlrange.Range["J15:J18"].Cells.Merge();
                xlrange.Range["K15:K18"].Cells.Merge();
                xlrange.Range["L15:L18"].Cells.Merge();

                xlrange.Range["A19:A22"].Cells.Merge();
                xlrange.Range["A19:A22"].Cells.Value = "2";
                xlrange.Range["B19:B22"].Cells.Merge();
                xlrange.Range["B19:B22"].Cells.Value = "B5";
                xlrange.Range["C19:C22"].Cells.Merge();
                xlrange.Range["D19:D22"].Cells.Merge();
                xlrange.Range["E19:E22"].Cells.Merge();
                xlrange.Range["F19:F22"].Cells.Merge();
                xlrange.Range["G19:G22"].Cells.Merge();
                xlrange.Range["H19:H22"].Cells.Merge();
                xlrange.Range["I19:I22"].Cells.Merge();
                xlrange.Range["J19:J22"].Cells.Merge();
                xlrange.Range["K19:K22"].Cells.Merge();
                xlrange.Range["L19:L22"].Cells.Merge();

                xlrange.Range["A23:A26"].Cells.Merge();
                xlrange.Range["A23:A26"].Cells.Value = "2";
                xlrange.Range["B23:B26"].Cells.Merge();
                xlrange.Range["B23:B26"].Cells.Value = "B6";
                xlrange.Range["C23:C26"].Cells.Merge();
                xlrange.Range["D23:D26"].Cells.Merge();
                xlrange.Range["E23:E26"].Cells.Merge();
                xlrange.Range["F23:F26"].Cells.Merge();
                xlrange.Range["G23:G26"].Cells.Merge();
                xlrange.Range["H23:H26"].Cells.Merge();
                xlrange.Range["I23:I26"].Cells.Merge();
                xlrange.Range["J23:J26"].Cells.Merge();
                xlrange.Range["K23:K26"].Cells.Merge();
                xlrange.Range["L23:L26"].Cells.Merge();

                xlrange.Range["A27:A30"].Cells.Merge();
                xlrange.Range["A27:A30"].Cells.Value = "2";
                xlrange.Range["B27:B30"].Cells.Merge();
                xlrange.Range["B27:B30"].Cells.Value = "B7";
                xlrange.Range["C27:C30"].Cells.Merge();
                xlrange.Range["D27:D30"].Cells.Merge();
                xlrange.Range["E27:E30"].Cells.Merge();
                xlrange.Range["F27:F30"].Cells.Merge();
                xlrange.Range["G27:G30"].Cells.Merge();
                xlrange.Range["H27:H30"].Cells.Merge();
                xlrange.Range["I27:I30"].Cells.Merge();
                xlrange.Range["J27:J30"].Cells.Merge();
                xlrange.Range["K27:K30"].Cells.Merge();
                xlrange.Range["L27:L30"].Cells.Merge();

                xlrange.Range["A31:A34"].Cells.Merge();
                xlrange.Range["A31:A34"].Cells.Value = "2";
                xlrange.Range["B31:B34"].Cells.Merge();
                xlrange.Range["B31:B34"].Cells.Value = "B8";
                xlrange.Range["C31:C34"].Cells.Merge();
                xlrange.Range["D31:D34"].Cells.Merge();
                xlrange.Range["E31:E34"].Cells.Merge();
                xlrange.Range["F31:F34"].Cells.Merge();
                xlrange.Range["G31:G34"].Cells.Merge();
                xlrange.Range["H31:H34"].Cells.Merge();
                xlrange.Range["I31:I34"].Cells.Merge();
                xlrange.Range["J31:J34"].Cells.Merge();
                xlrange.Range["K31:K34"].Cells.Merge();
                xlrange.Range["L31:L34"].Cells.Merge();

                xlrange.Range["A35:A38"].Cells.Merge();
                xlrange.Range["A35:A38"].Cells.Value = "2";
                xlrange.Range["B35:B38"].Cells.Merge();
                xlrange.Range["B35:B38"].Cells.Value = "B9";
                xlrange.Range["C35:C38"].Cells.Merge();
                xlrange.Range["D35:D38"].Cells.Merge();
                xlrange.Range["E35:E38"].Cells.Merge();
                xlrange.Range["F35:F38"].Cells.Merge();
                xlrange.Range["G35:G38"].Cells.Merge();
                xlrange.Range["H35:H38"].Cells.Merge();
                xlrange.Range["I35:I38"].Cells.Merge();
                xlrange.Range["J35:J38"].Cells.Merge();
                xlrange.Range["K35:K38"].Cells.Merge();
                xlrange.Range["L35:L38"].Cells.Merge();

                xlrange.Range["A39:A42"].Cells.Merge();
                xlrange.Range["A39:A42"].Cells.Value = "2";
                xlrange.Range["B39:B42"].Cells.Merge();
                xlrange.Range["B39:B42"].Cells.Value = "B10";
                xlrange.Range["C39:C42"].Cells.Merge();
                xlrange.Range["D39:D42"].Cells.Merge();
                xlrange.Range["E39:E42"].Cells.Merge();
                xlrange.Range["F39:F42"].Cells.Merge();
                xlrange.Range["G39:G42"].Cells.Merge();
                xlrange.Range["H39:H42"].Cells.Merge();
                xlrange.Range["I39:I42"].Cells.Merge();
                xlrange.Range["J39:J42"].Cells.Merge();
                xlrange.Range["K39:K42"].Cells.Merge();
                xlrange.Range["L39:L42"].Cells.Merge();




                xlworkbook.Close();
                xlapp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            lblText.Text = "Simplified Beam Schedule";
            Microsoft.Office.Interop.Excel.Application xlapp;
            Microsoft.Office.Interop.Excel.Workbook xlworkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlworksheet;
            Microsoft.Office.Interop.Excel.Range xlrange;
            try
            {
                xlapp = new Microsoft.Office.Interop.Excel.Application();
                xlworkbook = xlapp.Workbooks.Open(filenameexcel.Text);
                xlworksheet = xlworkbook.Worksheets["Sheet1"];
                xlrange = xlworksheet.UsedRange;

                xlrange.Range["A3:A18"].Cells.Merge();
                xlrange.Range["A3:A18"].Cells.Value = "3";
                xlrange.Range["B3:B18"].Cells.Merge();
                xlrange.Range["B3:B18"].Cells.Value = "B1";
                xlrange.Range["C3:C18"].Cells.Merge();
                xlrange.Range["D3:D18"].Cells.Merge();
                xlrange.Range["E3:E18"].Cells.Merge();
                xlrange.Range["F3:F18"].Cells.Merge();
                xlrange.Range["G3:G18"].Cells.Merge();
                xlrange.Range["H3:H18"].Cells.Merge();
                xlrange.Range["I3:I18"].Cells.Merge();
                xlrange.Range["J3:J18"].Cells.Merge();
                xlrange.Range["K3:K18"].Cells.Merge();
                xlrange.Range["L3:L18"].Cells.Merge();

                xlrange.Range["A19:A42"].Cells.Merge();
                xlrange.Range["A19:A42"].Cells.Value = "3";
                xlrange.Range["B19:B42"].Cells.Merge();
                xlrange.Range["B19:B42"].Cells.Value = "B2";
                xlrange.Range["C19:C42"].Cells.Merge();
                xlrange.Range["D19:D42"].Cells.Merge();
                xlrange.Range["E19:E42"].Cells.Merge();
                xlrange.Range["F19:F42"].Cells.Merge();
                xlrange.Range["G19:G42"].Cells.Merge();
                xlrange.Range["H19:H42"].Cells.Merge();
                xlrange.Range["I19:I42"].Cells.Merge();
                xlrange.Range["J19:J42"].Cells.Merge();
                xlrange.Range["K19:K42"].Cells.Merge();
                xlrange.Range["L19:L42"].Cells.Merge();






                xlworkbook.Close();
                xlapp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            lblText.Text = "Determination of the maximum number of bars for checking";
            Microsoft.Office.Interop.Excel.Application xlapp;
            Microsoft.Office.Interop.Excel.Workbook xlworkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlworksheet;
            Microsoft.Office.Interop.Excel.Range xlrange;
            try
            {
                xlapp = new Microsoft.Office.Interop.Excel.Application();
                xlworkbook = xlapp.Workbooks.Open(filenameexcel.Text);
                xlworksheet = xlworkbook.Worksheets["Sheet1"];
                xlrange = xlworksheet.UsedRange;

                xlrange.Range["A3:A42"].Cells.Merge();
                xlrange.Range["A3:A42"].Cells.Value = "4";
                xlrange.Range["B3:B42"].Cells.Merge();
                xlrange.Range["B3:B42"].Cells.Value = "B1";
                xlrange.Range["C3:C42"].Cells.Merge();
                xlrange.Range["D3:D42"].Cells.Merge();
                xlrange.Range["E3:E42"].Cells.Merge();
                xlrange.Range["F3:F42"].Cells.Merge();
                xlrange.Range["G3:G42"].Cells.Merge();
                xlrange.Range["H3:H42"].Cells.Merge();
                xlrange.Range["I3:I42"].Cells.Merge();
                xlrange.Range["J3:J42"].Cells.Merge();
                xlrange.Range["K3:K42"].Cells.Merge();
                xlrange.Range["L3:L42"].Cells.Merge();


                xlrange.Range["A3:A42"].Cells.UnMerge();
                xlrange.Range["B3:B42"].Cells.UnMerge();
                xlrange.Range["C3:C42"].Cells.UnMerge();
                xlrange.Range["D3:D42"].Cells.UnMerge();
                xlrange.Range["E3:E42"].Cells.UnMerge();
                xlrange.Range["F3:F42"].Cells.UnMerge();
                xlrange.Range["G3:G42"].Cells.UnMerge();
                xlrange.Range["H3:H42"].Cells.UnMerge();
                xlrange.Range["I3:I42"].Cells.UnMerge();
                xlrange.Range["J3:J42"].Cells.UnMerge();
                xlrange.Range["K3:K42"].Cells.UnMerge();
                xlrange.Range["L3:L42"].Cells.UnMerge();

                xlworkbook.Close();
                xlapp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void iconButton2_Click_2(object sender, EventArgs e)
        {
            lblText.Text = "Calculations of the simplification process";
            Reports r1 = new Reports();
            r1.Show();
        }

        private void iconButton6_Click_1(object sender, EventArgs e)
        {
            Second_Page page = new Second_Page();
            page.Show();
            this.Hide();
        }

        private void button1_MouseDown(object sender, MouseEventArgs e)
        {
            button1.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void button2_MouseDown(object sender, MouseEventArgs e)
        {
            button2.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void button2_MouseEnter(object sender, EventArgs e)
        {
            button2.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void button2_MouseHover(object sender, EventArgs e)
        {
            button2.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void button2_MouseLeave(object sender, EventArgs e)
        {
            button2.BackColor = Color.FromArgb(30, 27, 46);
        }

        private void button3_MouseDown(object sender, MouseEventArgs e)
        {
            button3.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void button3_MouseEnter(object sender, EventArgs e)
        {
            button3.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void button3_MouseHover(object sender, EventArgs e)
        {
            button3.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void button3_MouseLeave(object sender, EventArgs e)
        {
            button3.BackColor = Color.FromArgb(30, 27, 46);
        }

        private void button4_MouseDown(object sender, MouseEventArgs e)
        {
            button4.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void button4_MouseEnter(object sender, EventArgs e)
        {
            button4.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void button4_MouseHover(object sender, EventArgs e)
        {
            button4.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void button4_MouseLeave(object sender, EventArgs e)
        {
            button4.BackColor = Color.FromArgb(30, 27, 46);
        }

        private void iconButton2_MouseDown(object sender, MouseEventArgs e)
        {
            iconButton2.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void iconButton2_MouseEnter(object sender, EventArgs e)
        {
            iconButton2.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void iconButton2_MouseHover(object sender, EventArgs e)
        {
            iconButton2.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void iconButton2_MouseLeave(object sender, EventArgs e)
        {
            iconButton2.BackColor = Color.FromArgb(30, 27, 46);
        }

        private void iconButton6_MouseDown(object sender, MouseEventArgs e)
        {
            iconButton6.BackColor = Color.FromArgb(80, 80, 80);

        }

        private void iconButton6_MouseEnter(object sender, EventArgs e)
        {
            iconButton6.BackColor = Color.FromArgb(80, 80, 80);

        }

        private void iconButton6_MouseHover(object sender, EventArgs e)
        {
            iconButton6.BackColor = Color.FromArgb(80, 80, 80);

        }

        private void iconButton6_MouseLeave(object sender, EventArgs e)
        {
            iconButton6.BackColor = Color.FromArgb(30, 27, 46);
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
