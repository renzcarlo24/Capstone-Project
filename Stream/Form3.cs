using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FontAwesome.Sharp;
using Point = System.Drawing.Point;
using System.Runtime.InteropServices;
using Stream;
using Microsoft.Office.Interop.Word;

namespace Stream_25percent
{
    public partial class Form3 : Form
    {
        private IconButton currentBtn;
        private Panel leftBorderBtn;
        private int borderSize = 2;
        public Form3()
        {
            InitializeComponent();
            leftBorderBtn = new Panel();
            leftBorderBtn.Size = new Size(7, 60);
            panelMenu.Controls.Add(leftBorderBtn);
            this.Padding = new Padding(borderSize);

            this.Text = string.Empty;
            this.ControlBox = false;
            this.DoubleBuffered = true;
            this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
        }
        private struct RGBColors
        {
            public static Color color1 = Color.FromArgb(24,161, 251);
           
        }
        private void ActivateButton(object senderBtn, Color color)
        {
            if (senderBtn != null)
            {
                DisableButton();    
                currentBtn = (IconButton)senderBtn;
                currentBtn.BackColor = Color.FromArgb(37,36,81);
                currentBtn.BackColor = color;
                currentBtn.TextAlign = ContentAlignment.MiddleCenter;
                currentBtn.IconColor = color;
                currentBtn.TextImageRelation = TextImageRelation.ImageBeforeText;
                currentBtn.ImageAlign = ContentAlignment.MiddleRight;

                leftBorderBtn.BackColor = color;
                leftBorderBtn.Visible = true;

                iconCurrentChildForm.IconChar = currentBtn.IconChar;
                iconCurrentChildForm.IconColor = color;
                

                
            }

        }
        private void DisableButton()
        {
            if (currentBtn != null)
            {
                currentBtn.BackColor = Color.FromArgb(31, 30, 68);
                currentBtn.ForeColor = Color.Gainsboro;
                currentBtn.TextAlign = ContentAlignment.MiddleCenter;
                currentBtn.IconColor = Color.Gainsboro;
                currentBtn.TextImageRelation = TextImageRelation.ImageBeforeText;   
                currentBtn.ImageAlign= ContentAlignment.MiddleRight;
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlapp;
            Microsoft.Office.Interop.Excel.Workbook xlworkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlworksheet;
            Microsoft.Office.Interop.Excel.Range xlrange;
            try
            {
                xlapp = new Microsoft.Office.Interop.Excel.Application();
                xlworkbook = xlapp.Workbooks.Open(FileName_LBL.Text);
                xlworksheet = xlworkbook.Worksheets["Sheet1"];
                xlrange = xlworksheet.UsedRange;

                Data_GRD.ColumnCount = xlrange.Columns.Count;

                for (int xlrow = 1; xlrow <= xlrange.Rows.Count; xlrow++)
                {
                    Data_GRD.Rows.Add(xlrange.Cells[xlrow, 1].Text, xlrange.Cells[xlrow, 2].Text, xlrange.Cells[xlrow, 3].Text, xlrange.Cells[xlrow, 4].Text, xlrange.Cells[xlrow, 5].Text, xlrange.Cells[xlrow, 6].Text,
                        xlrange.Cells[xlrow, 7].Text, xlrange.Cells[xlrow, 8].Text, xlrange.Cells[xlrow, 9].Text, xlrange.Cells[xlrow, 10].Text, xlrange.Cells[xlrow, 11].Text, xlrange.Cells[xlrow, 12].Text);
                }



                xlworkbook.Close();
                xlapp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Files Only | *.xlsx; *.xls";
                ofd.Title = "Choose the file";
                if (ofd.ShowDialog() == DialogResult.OK)
                    FileName_LBL.Text = ofd.FileName;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlapp;
            Microsoft.Office.Interop.Excel.Workbook xlworkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlworksheet;
            Microsoft.Office.Interop.Excel.Range xlrange;
            try
            {
                xlapp = new Microsoft.Office.Interop.Excel.Application();
                xlworkbook = xlapp.Workbooks.Open(FileName_LBL.Text);
                xlworksheet = xlworkbook.Worksheets["Sheet1"];
                xlrange = xlworksheet.UsedRange;

                Data_GRD.ColumnCount = xlrange.Columns.Count;

                for (int xlrow = 1; xlrow <= xlrange.Rows.Count; xlrow++)
                {
                    Data_GRD.Rows.Add(xlrange.Cells[xlrow, 1].Text, xlrange.Cells[xlrow, 2].Text, xlrange.Cells[xlrow, 3].Text, xlrange.Cells[xlrow, 4].Text, xlrange.Cells[xlrow, 5].Text, xlrange.Cells[xlrow, 6].Text,
                        xlrange.Cells[xlrow, 7].Text, xlrange.Cells[xlrow, 8].Text, xlrange.Cells[xlrow, 9].Text, xlrange.Cells[xlrow, 10].Text, xlrange.Cells[xlrow, 11].Text, xlrange.Cells[xlrow, 12].Text);
                }

                

                xlworkbook.Close();
                xlapp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlapp;
            Microsoft.Office.Interop.Excel.Workbook xlworkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlworksheet;
            Microsoft.Office.Interop.Excel.Range xlrange;
            try
            {
                xlapp = new Microsoft.Office.Interop.Excel.Application();
                xlworkbook = xlapp.Workbooks.Open(FileName_LBL.Text);
                xlworksheet = xlworkbook.Worksheets["Sheet1"];
                xlrange = xlworksheet.UsedRange;

                xlrange.Range["A3"].Value = "2";
                xlrange.Range[xlrange.Range["B3"], xlrange.Range["B5"]].Merge();
                xlrange.Range[xlrange.Range["C3"], xlrange.Range["C5"]].Merge();
                xlrange.Range[xlrange.Range["D3"], xlrange.Range["D5"]].Merge();
                xlrange.Range[xlrange.Range["E3"], xlrange.Range["E5"]].Merge();
                xlrange.Range[xlrange.Range["F3"], xlrange.Range["F5"]].Merge();
                xlrange.Range[xlrange.Range["G3"], xlrange.Range["G5"]].Merge();
                xlrange.Range["G3"].Value = "3-#16";
                xlrange.Range[xlrange.Range["H3"], xlrange.Range["H5"]].Merge();
                xlrange.Range[xlrange.Range["I5"], xlrange.Range["I6"]].Merge();
                xlrange.Range[xlrange.Range["J5"], xlrange.Range["J3"]].Merge();
                xlrange.Range[xlrange.Range["K3"], xlrange.Range["K4"]].Merge();


                xlrange.Range["A8"].Value = "2";
                xlrange.Range[xlrange.Range["B8"], xlrange.Range["B10"]].Merge();
                xlrange.Range["B8"].Value = "B2";
                xlrange.Range[xlrange.Range["C8"], xlrange.Range["C10"]].Merge();
                xlrange.Range[xlrange.Range["D8"], xlrange.Range["D10"]].Merge();
                xlrange.Range[xlrange.Range["E8"], xlrange.Range["E10"]].Merge();
                xlrange.Range[xlrange.Range["F8"], xlrange.Range["F10"]].Merge();
                xlrange.Range[xlrange.Range["G8"], xlrange.Range["G10"]].Merge();
                xlrange.Range[xlrange.Range["H8"], xlrange.Range["H10"]].Merge();
                xlrange.Range[xlrange.Range["I8"], xlrange.Range["I10"]].Merge();
                xlrange.Range[xlrange.Range["J8"], xlrange.Range["J10"]].Merge();
                xlrange.Range[xlrange.Range["K8"], xlrange.Range["K10"]].Merge();

                xlrange.Range["A12"].Value = "2";
                xlrange.Range[xlrange.Range["B12"], xlrange.Range["B14"]].Merge();
                xlrange.Range["B12"].Value = "B3";
                xlrange.Range[xlrange.Range["C12"], xlrange.Range["C14"]].Merge();
                xlrange.Range[xlrange.Range["D12"], xlrange.Range["D14"]].Merge();
                xlrange.Range[xlrange.Range["E12"], xlrange.Range["E14"]].Merge();
                xlrange.Range[xlrange.Range["F12"], xlrange.Range["F14"]].Merge();
                xlrange.Range[xlrange.Range["G12"], xlrange.Range["G14"]].Merge();
                xlrange.Range[xlrange.Range["H12"], xlrange.Range["H14"]].Merge();
                xlrange.Range[xlrange.Range["I12"], xlrange.Range["I14"]].Merge();
                xlrange.Range[xlrange.Range["I13"], xlrange.Range["I15"]].Merge();
                xlrange.Range[xlrange.Range["J12"], xlrange.Range["J14"]].Merge();
                xlrange.Range[xlrange.Range["K13"], xlrange.Range["K14"]].Merge();

                xlrange.Range["A17"].Value = "2";
                xlrange.Range[xlrange.Range["B17"], xlrange.Range["B19", "B21"]].Merge();
                xlrange.Range["B17"].Value = "B4";
                xlrange.Range[xlrange.Range["C17"], xlrange.Range["C19", "C21"]].Merge();
                xlrange.Range[xlrange.Range["D17"], xlrange.Range["D19", "D21"]].Merge();
                xlrange.Range[xlrange.Range["E17"], xlrange.Range["E19", "E21"]].Merge();
                xlrange.Range[xlrange.Range["F17"], xlrange.Range["F19", "F21"]].Merge();
                xlrange.Range[xlrange.Range["G17"], xlrange.Range["G19", "G21"]].Merge();
                xlrange.Range["G17"].Value = "3-#16";
                xlrange.Range[xlrange.Range["H17"], xlrange.Range["H19", "H21"]].Merge();
                xlrange.Range[xlrange.Range["I20"], xlrange.Range["I21"]].Merge();
                xlrange.Range[xlrange.Range["I19"], xlrange.Range["I20"]].Merge();
                xlrange.Range[xlrange.Range["I18"], xlrange.Range["I19"]].Merge();
                xlrange.Range[xlrange.Range["J18"], xlrange.Range["J19", "J21"]].Merge();
                xlrange.Range[xlrange.Range["K19"], xlrange.Range["K21"]].Merge();
                xlrange.Range[xlrange.Range["K18"], xlrange.Range["K19"]].Merge();

                xlrange.Range["A23"].Value = "2";
                xlrange.Range["B23"].Value = "B5";

                xlrange.Range["A25"].Value = "2";
                xlrange.Range[xlrange.Range["B25"], xlrange.Range["B26", "B28"]].Merge();
                xlrange.Range["B25"].Value = "B6";
                xlrange.Range[xlrange.Range["C25"], xlrange.Range["C26", "C28"]].Merge();
                xlrange.Range[xlrange.Range["D25"], xlrange.Range["D26", "D28"]].Merge();
                xlrange.Range[xlrange.Range["E25"], xlrange.Range["E26", "E28"]].Merge();
                xlrange.Range[xlrange.Range["F25"], xlrange.Range["F26", "F28"]].Merge();
                xlrange.Range[xlrange.Range["G25"], xlrange.Range["G26", "G28"]].Merge();
                xlrange.Range["G25"].Value = "3-#16";
                xlrange.Range[xlrange.Range["H25"], xlrange.Range["H26", "H28"]].Merge();
                xlrange.Range[xlrange.Range["I25"], xlrange.Range["I26", "I28"]].Merge();
                xlrange.Range[xlrange.Range["J25"], xlrange.Range["J26", "J28"]].Merge();
                xlrange.Range[xlrange.Range["K28"], xlrange.Range["K27", "K26"]].Merge();
                xlrange.Range[xlrange.Range["K28"], xlrange.Range["K25"]].Merge();

                xlrange.Range["A31"].Value = "2";
                xlrange.Range[xlrange.Range["B31"], xlrange.Range["B32", "B33"]].Merge();
                xlrange.Range["B31"].Value = "B7";
                xlrange.Range[xlrange.Range["C31"], xlrange.Range["C32", "C33"]].Merge();
                xlrange.Range[xlrange.Range["D31"], xlrange.Range["D32", "D33"]].Merge();
                xlrange.Range[xlrange.Range["E31"], xlrange.Range["E32", "E33"]].Merge();
                xlrange.Range[xlrange.Range["F31"], xlrange.Range["F32", "F33"]].Merge();
                xlrange.Range[xlrange.Range["G31"], xlrange.Range["G32", "G33"]].Merge();
                xlrange.Range[xlrange.Range["H31"], xlrange.Range["H32", "H33"]].Merge();
                xlrange.Range[xlrange.Range["I31"], xlrange.Range["I32", "I33"]].Merge();
                xlrange.Range[xlrange.Range["J31"], xlrange.Range["J32", "J33"]].Merge();
                xlrange.Range[xlrange.Range["K31"], xlrange.Range["K32", "K33"]].Merge();

                xlrange.Range["A35"].Value = "2";
                xlrange.Range["B35"].Value = "B8";

                xlrange.Range["A37"].Value = "2";
                xlrange.Range[xlrange.Range["B37"], xlrange.Range["B38", "B40"]].Merge();
                xlrange.Range["B37"].Value = "B9";
                xlrange.Range[xlrange.Range["C37"], xlrange.Range["C38", "C40"]].Merge();
                xlrange.Range[xlrange.Range["D37"], xlrange.Range["D38", "D40"]].Merge();
                xlrange.Range[xlrange.Range["E37"], xlrange.Range["E38", "E40"]].Merge();
                xlrange.Range[xlrange.Range["F37"], xlrange.Range["F38", "F40"]].Merge();
                xlrange.Range[xlrange.Range["G37"], xlrange.Range["G38", "G40"]].Merge();
                xlrange.Range["G37"].Value = "3-#16";
                xlrange.Range[xlrange.Range["H37"], xlrange.Range["H38", "H40"]].Merge();
                xlrange.Range[xlrange.Range["I40"], xlrange.Range["I38", "I37"]].Merge();
                xlrange.Range[xlrange.Range["J37"], xlrange.Range["J38", "J40"]].Merge();
                xlrange.Range[xlrange.Range["K37"], xlrange.Range["K38"]].Merge();
                xlrange.Range[xlrange.Range["K39"], xlrange.Range["K40"]].Merge();

                xlrange.Range["A43"].Value = "2";
                xlrange.Range[xlrange.Range["B43"], xlrange.Range["B45"]].Merge();
                xlrange.Range["B43"].Value = "B10";
                xlrange.Range[xlrange.Range["C43"], xlrange.Range["C45"]].Merge();
                xlrange.Range[xlrange.Range["D43"], xlrange.Range["D45"]].Merge();
                xlrange.Range[xlrange.Range["E43"], xlrange.Range["E45"]].Merge();
                xlrange.Range[xlrange.Range["F43"], xlrange.Range["F45"]].Merge();
                xlrange.Range[xlrange.Range["G43"], xlrange.Range["G45"]].Merge();
                xlrange.Range[xlrange.Range["H43"], xlrange.Range["H45"]].Merge();
                xlrange.Range[xlrange.Range["I44"], xlrange.Range["I45", "I46"]].Merge();
                xlrange.Range[xlrange.Range["J43"], xlrange.Range["J45"]].Merge();
                xlrange.Range[xlrange.Range["K44"], xlrange.Range["K45", "K46"]].Merge();


                xlworkbook.Close();
                xlapp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlapp;
            Microsoft.Office.Interop.Excel.Workbook xlworkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlworksheet;
            Microsoft.Office.Interop.Excel.Range xlrange;
            try
            {
                xlapp = new Microsoft.Office.Interop.Excel.Application();
                xlworkbook = xlapp.Workbooks.Open(FileName_LBL.Text);
                xlworksheet = xlworkbook.Worksheets["Sheet1"];
                xlrange = xlworksheet.UsedRange;

                xlrange.Range["A3"].Value = "3";
                xlrange.Range["B3"].Value = "B1";
                xlrange.Range["A17"].Value = "3";
                xlrange.Range["B17"].Value = "B2";
                xlrange.Range["A8"].Value = "3";
                xlrange.Range["B8"].Value = "B3";
                xlrange.Range["A23"].Value = "3";
                xlrange.Range["B23"].Value = "B4";

                xlrange.Range[xlrange.Range["A12", "A13"], xlrange.Range["A14", "A15"]].Value = "";
                xlrange.Range[xlrange.Range["B12", "B13"], xlrange.Range["B14", "B15"]].Value = "";
                xlrange.Range[xlrange.Range["C12", "C13"], xlrange.Range["C14", "C15"]].Value = "";
                xlrange.Range[xlrange.Range["D12", "D13"], xlrange.Range["D14", "D15"]].Value = "";
                xlrange.Range[xlrange.Range["E12", "E13"], xlrange.Range["E14", "E15"]].Value = "";
                xlrange.Range[xlrange.Range["F12", "F13"], xlrange.Range["F14", "F15"]].Value = "";
                xlrange.Range[xlrange.Range["G12", "G13"], xlrange.Range["G14", "F15"]].Value = "";
                xlrange.Range[xlrange.Range["G12", "G13"], xlrange.Range["G14", "G15"]].Value = "";
                xlrange.Range[xlrange.Range["H12", "H13"], xlrange.Range["H14", "H15"]].Value = "";
                xlrange.Range[xlrange.Range["I12", "I13"], xlrange.Range["I14", "I15"]].Value = "";
                xlrange.Range[xlrange.Range["J12", "J13"], xlrange.Range["J14", "J15"]].Value = "";
                xlrange.Range[xlrange.Range["K12", "K13"], xlrange.Range["K14", "K15"]].Value = "";



                
                xlrange.Range[xlrange.Range["A25", "A26"], xlrange.Range["A27", "A28"]].Value = "";
                xlrange.Range[xlrange.Range["A29", "B25"], xlrange.Range["B26", "B27"]].Value = "";
                xlrange.Range[xlrange.Range["B28", "B29"], xlrange.Range["C25", "C26"]].Value = "";
                xlrange.Range[xlrange.Range["C27", "C28"], xlrange.Range["C29", "D25"]].Value = "";
                xlrange.Range[xlrange.Range["D26", "D27"], xlrange.Range["D28", "D29"]].Value = "";
                xlrange.Range[xlrange.Range["E25", "E26"], xlrange.Range["E27", "E28"]].Value = "";
                xlrange.Range[xlrange.Range["E29", "F25"], xlrange.Range["F26", "F27"]].Value = "";
                xlrange.Range[xlrange.Range["F28", "F29"], xlrange.Range["G25", "G26"]].Value = "";
                xlrange.Range[xlrange.Range["G27", "G28"], xlrange.Range["G29", "H25"]].Value = "";
                xlrange.Range[xlrange.Range["H26", "H27"], xlrange.Range["H28", "H29"]].Value = "";
                xlrange.Range[xlrange.Range["I25", "I26"], xlrange.Range["I27", "I28"]].Value = "";
                xlrange.Range[xlrange.Range["I29", "J25"], xlrange.Range["J26", "J27"]].Value = "";
                xlrange.Range[xlrange.Range["J28", "J29"], xlrange.Range["K25", "K26"]].Value = "";
                xlrange.Range[xlrange.Range["K27", "K28"], xlrange.Range["K29"]].Value = "";

                
                xlrange.Range[xlrange.Range["A37", "A38"], xlrange.Range["A39", "A40"]].Value = "";
                xlrange.Range[xlrange.Range["A41", "B37"], xlrange.Range["B38", "B39"]].Value = "";
                xlrange.Range[xlrange.Range["B40", "B41"], xlrange.Range["C37", "C38"]].Value = "";
                xlrange.Range[xlrange.Range["C39", "C40"], xlrange.Range["C41", "D37"]].Value = "";
                xlrange.Range[xlrange.Range["D38", "D39"], xlrange.Range["D40", "D41"]].Value = "";
                xlrange.Range[xlrange.Range["E37", "E38"], xlrange.Range["E39", "E40"]].Value = "";
                xlrange.Range[xlrange.Range["E41", "F37"], xlrange.Range["F38", "F39"]].Value = "";
                xlrange.Range[xlrange.Range["F40", "F41"], xlrange.Range["G37", "G38"]].Value = "";
                xlrange.Range[xlrange.Range["G39", "G40"], xlrange.Range["G41", "H37"]].Value = "";
                xlrange.Range[xlrange.Range["H38", "H39"], xlrange.Range["H40", "H41"]].Value = "";
                xlrange.Range[xlrange.Range["I37", "I38"], xlrange.Range["I39"]].Value = "";
                xlrange.Range[xlrange.Range["I40", "I41"], xlrange.Range["J37", "J38"]].Value = "";
                xlrange.Range[xlrange.Range["J39", "J40"], xlrange.Range["J41", "K37"]].Value = "";
                xlrange.Range[xlrange.Range["K38", "K39"], xlrange.Range["K40", "K41"]].Value = "";


                
                xlrange.Range[xlrange.Range["A43", "A44"], xlrange.Range["A45", "A46"]].Value = "";
                xlrange.Range[xlrange.Range["B43", "B44"], xlrange.Range["B45", "B46"]].Value = "";
                xlrange.Range[xlrange.Range["C43", "C44"], xlrange.Range["C45", "C46"]].Value = "";
                xlrange.Range[xlrange.Range["D43", "D44"], xlrange.Range["D45", "D46"]].Value = "";
                xlrange.Range[xlrange.Range["E43", "E44"], xlrange.Range["E45", "E46"]].Value = "";
                xlrange.Range[xlrange.Range["F43", "F44"], xlrange.Range["F45", "F46"]].Value = "";
                xlrange.Range[xlrange.Range["G43", "G44"], xlrange.Range["G45", "F46"]].Value = "";
                xlrange.Range[xlrange.Range["G43", "G44"], xlrange.Range["G45", "G46"]].Value = "";
                xlrange.Range[xlrange.Range["H43", "H44"], xlrange.Range["H45", "H46"]].Value = "";
                xlrange.Range[xlrange.Range["I43", "I44"], xlrange.Range["I45", "I46"]].Value = "";
                xlrange.Range[xlrange.Range["J43", "J44"], xlrange.Range["J45", "J46"]].Value = "";
                xlrange.Range[xlrange.Range["K43", "K44"], xlrange.Range["K45", "K46"]].Value = "";

                
                xlrange.Range[xlrange.Range["A31", "A32"], xlrange.Range["A33", "B31"]].Value = "";
                xlrange.Range[xlrange.Range["B32", "B33"], xlrange.Range["C31", "C32"]].Value = "";
                xlrange.Range[xlrange.Range["C33", "D31"], xlrange.Range["D32", "D33"]].Value = "";
                xlrange.Range[xlrange.Range["E31", "E32"], xlrange.Range["E33", "F31"]].Value = "";
                xlrange.Range[xlrange.Range["F32", "F33"], xlrange.Range["G31", "G32"]].Value = "";
                xlrange.Range[xlrange.Range["G33", "H31"], xlrange.Range["H32", "H33"]].Value = "";
                xlrange.Range[xlrange.Range["I31", "I32"], xlrange.Range["I33", "J31"]].Value = "";
                xlrange.Range[xlrange.Range["J32", "J33"], xlrange.Range["K31", "K32"]].Value = "";
                xlrange.Range["K33"].Value = "";



                
                xlrange.Range[xlrange.Range["A35", "B35"], xlrange.Range["C35", "D35"]].Value = "";
                xlrange.Range[xlrange.Range["E35", "F35"], xlrange.Range["G35", "H35"]].Value = "";
                xlrange.Range[xlrange.Range["I35"], xlrange.Range["J35"]].Value = "";
                xlrange.Range["K35"].Value = "";


                xlworkbook.Close();
                xlapp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlapp;
            Microsoft.Office.Interop.Excel.Workbook xlworkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlworksheet;
            Microsoft.Office.Interop.Excel.Range xlrange;
            try
            {
                xlapp = new Microsoft.Office.Interop.Excel.Application();
                xlworkbook = xlapp.Workbooks.Open(FileName_LBL.Text);
                xlworksheet = xlworkbook.Worksheets["Sheet1"];
                xlrange = xlworksheet.UsedRange;

                xlrange.Range["A3"].Value = "4";

                
                xlrange.Range[xlrange.Range["A8", "A9"], xlrange.Range["A10", "B8"]].Value = "";
                xlrange.Range[xlrange.Range["B9", "B10"], xlrange.Range["C8", "C9"]].Value = "";
                xlrange.Range[xlrange.Range["C10", "D8"], xlrange.Range["D9", "D10"]].Value = "";
                xlrange.Range[xlrange.Range["E8", "E9"], xlrange.Range["E10", "F8"]].Value = "";
                xlrange.Range[xlrange.Range["F9", "F10"], xlrange.Range["G8", "G9"]].Value = "";
                xlrange.Range[xlrange.Range["G10", "H8"], xlrange.Range["H9", "H10"]].Value = "";
                xlrange.Range[xlrange.Range["I8", "I9"], xlrange.Range["I10", "J8"]].Value = "";
                xlrange.Range[xlrange.Range["J9", "J10"], xlrange.Range["K8", "K9"]].Value = "";
                xlrange.Range["K10"].Value = "";

                
                xlrange.Range[xlrange.Range["A17", "A18"], xlrange.Range["A19", "A20"]].Value = "";
                xlrange.Range[xlrange.Range["A21", "B17"], xlrange.Range["B18", "B19"]].Value = "";
                xlrange.Range[xlrange.Range["B20", "B21"], xlrange.Range["C17", "C18"]].Value = "";
                xlrange.Range[xlrange.Range["C19", "C20"], xlrange.Range["C21", "D17"]].Value = "";
                xlrange.Range[xlrange.Range["D18", "D19"], xlrange.Range["D20", "D21"]].Value = "";
                xlrange.Range[xlrange.Range["E17", "E18"], xlrange.Range["E19", "E20"]].Value = "";
                xlrange.Range[xlrange.Range["E21", "F17"], xlrange.Range["F18", "F19"]].Value = "";
                xlrange.Range[xlrange.Range["F20", "F21"], xlrange.Range["G17", "G18"]].Value = "";
                xlrange.Range[xlrange.Range["G19", "G20"], xlrange.Range["G21", "H17"]].Value = "";
                xlrange.Range[xlrange.Range["H18", "H19"], xlrange.Range["H20", "H21"]].Value = "";
                xlrange.Range[xlrange.Range["I17", "I18"], xlrange.Range["I19"]].Value = "";
                xlrange.Range[xlrange.Range["I20", "I21"], xlrange.Range["J17", "J18"]].Value = "";
                xlrange.Range[xlrange.Range["J19", "J20"], xlrange.Range["J21", "K17"]].Value = "";
                xlrange.Range[xlrange.Range["K18", "K19"], xlrange.Range["K20", "K21"]].Value = "";

                
                xlrange.Range[xlrange.Range["A23", "B23"], xlrange.Range["C23", "D23"]].Value = "";
                xlrange.Range[xlrange.Range["E23", "F23"], xlrange.Range["G23", "H23"]].Value = "";
                xlrange.Range[xlrange.Range["I23"], xlrange.Range["J23"]].Value = "";
                xlrange.Range["K23"].Value = "";





                xlworkbook.Close();
                xlapp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void back_Click(object sender, EventArgs e)
        {
            Second_Page page = new Second_Page();
            page.Show();
            this.Hide();
        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }

        private void iconButton2_Click(object sender, EventArgs e)
        {
            ActivateButton(sender, RGBColors.color1);
            Microsoft.Office.Interop.Excel.Application xlapp;
            Microsoft.Office.Interop.Excel.Workbook xlworkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlworksheet;
            Microsoft.Office.Interop.Excel.Range xlrange;
            try
            {
                xlapp = new Microsoft.Office.Interop.Excel.Application();
                xlworkbook = xlapp.Workbooks.Open(FileName_LBL.Text);
                xlworksheet = xlworkbook.Worksheets["Sheet1"];
                xlrange = xlworksheet.UsedRange;

                Data_GRD.ColumnCount = xlrange.Columns.Count;

                for (int xlrow = 1; xlrow <= xlrange.Rows.Count; xlrow++)
                {
                    Data_GRD.Rows.Add(xlrange.Cells[xlrow, 1].Text, xlrange.Cells[xlrow, 2].Text, xlrange.Cells[xlrow, 3].Text, xlrange.Cells[xlrow, 4].Text, xlrange.Cells[xlrow, 5].Text, xlrange.Cells[xlrow, 6].Text,
                        xlrange.Cells[xlrow, 7].Text, xlrange.Cells[xlrow, 8].Text, xlrange.Cells[xlrow, 9].Text, xlrange.Cells[xlrow, 10].Text, xlrange.Cells[xlrow, 11].Text, xlrange.Cells[xlrow, 12].Text);
                }



                xlworkbook.Close();
                xlapp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void iconButton1_Click(object sender, EventArgs e)
        {
            ActivateButton(sender, RGBColors.color1);
            Second_Page page = new Second_Page();
            page.Show();
            this.Hide();
        }
        

        private void iconButton3_Click(object sender, EventArgs e)
        {
            ActivateButton(sender, RGBColors.color1);
            Microsoft.Office.Interop.Excel.Application xlapp;
            Microsoft.Office.Interop.Excel.Workbook xlworkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlworksheet;
            Microsoft.Office.Interop.Excel.Range xlrange;
            try
            {
                xlapp = new Microsoft.Office.Interop.Excel.Application();
                xlworkbook = xlapp.Workbooks.Open(FileName_LBL.Text);
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

                xlrange.Range["A43:A46"].Cells.Merge();
                xlrange.Range["A43:A46"].Cells.Value = "2";
                xlrange.Range["B43:B46"].Cells.Merge();
                xlrange.Range["B43:B46"].Cells.Value = "B11";
                xlrange.Range["C43:C46"].Cells.Merge();
                xlrange.Range["D43:D46"].Cells.Merge();
                xlrange.Range["E43:E46"].Cells.Merge();
                xlrange.Range["F43:F46"].Cells.Merge();
                xlrange.Range["G43:G46"].Cells.Merge();
                xlrange.Range["H43:H46"].Cells.Merge();
                xlrange.Range["I43:I46"].Cells.Merge();
                xlrange.Range["J43:J46"].Cells.Merge();
                xlrange.Range["K43:K46"].Cells.Merge();

                xlrange.Range["A47:A51"].Cells.Merge();
                xlrange.Range["A47:A51"].Cells.Value = "2";
                xlrange.Range["B47:B51"].Cells.Merge();
                xlrange.Range["B47:B51"].Cells.Value = "B12";
                xlrange.Range["C47:C51"].Cells.Merge();
                xlrange.Range["D47:D51"].Cells.Merge();
                xlrange.Range["E47:E51"].Cells.Merge();
                xlrange.Range["F47:F51"].Cells.Merge();
                xlrange.Range["G47:G51"].Cells.Merge();
                xlrange.Range["H47:H51"].Cells.Merge();
                xlrange.Range["I47:I51"].Cells.Merge();
                xlrange.Range["J47:J51"].Cells.Merge();
                xlrange.Range["K47:K51"].Cells.Merge();




                xlworkbook.Close();
                xlapp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void iconButton4_Click(object sender, EventArgs e)
        {
            ActivateButton(sender, RGBColors.color1);
            Microsoft.Office.Interop.Excel.Application xlapp;
            Microsoft.Office.Interop.Excel.Workbook xlworkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlworksheet;
            Microsoft.Office.Interop.Excel.Range xlrange;
            try
            {
                xlapp = new Microsoft.Office.Interop.Excel.Application();
                xlworkbook = xlapp.Workbooks.Open(FileName_LBL.Text);
                xlworksheet = xlworkbook.Worksheets["Sheet1"];
                xlrange = xlworksheet.UsedRange;

                xlrange.Range["A3:A10"].Cells.Merge();
                xlrange.Range["A3:A10"].Cells.Value = "3";
                xlrange.Range["B3:B10"].Cells.Merge();
                xlrange.Range["B3:B10"].Cells.Value = "B1";
                xlrange.Range["C3:C10"].Cells.Merge();
                xlrange.Range["D3:D10"].Cells.Merge();
                xlrange.Range["E3:E10"].Cells.Merge();
                xlrange.Range["F3:F10"].Cells.Merge();
                xlrange.Range["G3:G10"].Cells.Merge();
                xlrange.Range["H3:H10"].Cells.Merge();
                xlrange.Range["I3:I10"].Cells.Merge();
                xlrange.Range["J3:J10"].Cells.Merge();
                xlrange.Range["K3:K10"].Cells.Merge();

                xlrange.Range["A11:A18"].Cells.Merge();
                xlrange.Range["A11:A18"].Cells.Value = "3";
                xlrange.Range["B11:B18"].Cells.Merge();
                xlrange.Range["B11:B18"].Cells.Value = "B2";
                xlrange.Range["C11:C18"].Cells.Merge();
                xlrange.Range["D11:D18"].Cells.Merge();
                xlrange.Range["E11:E18"].Cells.Merge();
                xlrange.Range["F11:F18"].Cells.Merge();
                xlrange.Range["G11:G18"].Cells.Merge();
                xlrange.Range["H11:H18"].Cells.Merge();
                xlrange.Range["I11:I18"].Cells.Merge();
                xlrange.Range["J11:J18"].Cells.Merge();
                xlrange.Range["K11:K18"].Cells.Merge();

                xlrange.Range["A19:A26"].Cells.Merge();
                xlrange.Range["A19:A26"].Cells.Value = "3";
                xlrange.Range["B19:B26"].Cells.Merge();
                xlrange.Range["B19:B26"].Cells.Value = "B3";
                xlrange.Range["C19:C26"].Cells.Merge();
                xlrange.Range["D19:D26"].Cells.Merge();
                xlrange.Range["E19:E26"].Cells.Merge();
                xlrange.Range["F19:F26"].Cells.Merge();
                xlrange.Range["G19:G26"].Cells.Merge();
                xlrange.Range["H19:H26"].Cells.Merge();
                xlrange.Range["I19:I26"].Cells.Merge();
                xlrange.Range["J19:J26"].Cells.Merge();
                xlrange.Range["K19:K26"].Cells.Merge();

                xlrange.Range["A27:A34"].Cells.Merge();
                xlrange.Range["A27:A34"].Cells.Value = "3";
                xlrange.Range["B27:B34"].Cells.Merge();
                xlrange.Range["B27:B34"].Cells.Value = "B4";
                xlrange.Range["C27:C34"].Cells.Merge();
                xlrange.Range["D27:D34"].Cells.Merge();
                xlrange.Range["E27:E34"].Cells.Merge();
                xlrange.Range["F27:F34"].Cells.Merge();
                xlrange.Range["G27:G34"].Cells.Merge();
                xlrange.Range["H27:H34"].Cells.Merge();
                xlrange.Range["I27:I34"].Cells.Merge();
                xlrange.Range["J27:J34"].Cells.Merge();
                xlrange.Range["K27:K34"].Cells.Merge();


                xlrange.Range["A35:A42"].Cells.Merge();
                xlrange.Range["A35:A42"].Cells.Value = "3";
                xlrange.Range["B35:B42"].Cells.Merge();
                xlrange.Range["B35:B42"].Cells.Value = "B5";
                xlrange.Range["C35:C42"].Cells.Merge();
                xlrange.Range["D35:D42"].Cells.Merge();
                xlrange.Range["E35:E42"].Cells.Merge();
                xlrange.Range["F35:F42"].Cells.Merge();
                xlrange.Range["G35:G42"].Cells.Merge();
                xlrange.Range["H35:H42"].Cells.Merge();
                xlrange.Range["I35:I42"].Cells.Merge();
                xlrange.Range["J35:J42"].Cells.Merge();
                xlrange.Range["K35:K42"].Cells.Merge();

                xlrange.Range["A43:A51"].Cells.Merge();
                xlrange.Range["A43:A51"].Cells.Value = "3";
                xlrange.Range["B43:B51"].Cells.Merge();
                xlrange.Range["B43:B51"].Cells.Value = "B6";
                xlrange.Range["C43:C51"].Cells.Merge();
                xlrange.Range["D43:D51"].Cells.Merge();
                xlrange.Range["E43:E51"].Cells.Merge();
                xlrange.Range["F43:F51"].Cells.Merge();
                xlrange.Range["G43:G51"].Cells.Merge();
                xlrange.Range["H43:H51"].Cells.Merge();
                xlrange.Range["I43:I51"].Cells.Merge();
                xlrange.Range["J43:J51"].Cells.Merge();
                xlrange.Range["K43:K51"].Cells.Merge();


                xlworkbook.Close();
                xlapp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void iconButton5_Click(object sender, EventArgs e)
        {
            ActivateButton(sender, RGBColors.color1);
            Microsoft.Office.Interop.Excel.Application xlapp;
            Microsoft.Office.Interop.Excel.Workbook xlworkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlworksheet;
            Microsoft.Office.Interop.Excel.Range xlrange;
            try
            {
                xlapp = new Microsoft.Office.Interop.Excel.Application();
                xlworkbook = xlapp.Workbooks.Open(FileName_LBL.Text);
                xlworksheet = xlworkbook.Worksheets["Sheet1"];
                xlrange = xlworksheet.UsedRange;

                xlrange.Range["A3:A51"].Cells.Merge();
                xlrange.Range["A3:A51"].Cells.Value = "4";
                xlrange.Range["B3:B51"].Cells.Merge();
                xlrange.Range["B3:B51"].Cells.Value = "B1";
                xlrange.Range["C3:C51"].Cells.Merge();
                xlrange.Range["D3:D51"].Cells.Merge();
                xlrange.Range["E3:E51"].Cells.Merge();
                xlrange.Range["F3:F51"].Cells.Merge();
                xlrange.Range["G3:G51"].Cells.Merge();
                xlrange.Range["H3:H51"].Cells.Merge();
                xlrange.Range["I3:I51"].Cells.Merge();
                xlrange.Range["J3:J51"].Cells.Merge();
                xlrange.Range["K3:K51"].Cells.Merge();

                xlrange.Range["A3:A51"].Cells.UnMerge(); 
                xlrange.Range["B3:B51"].Cells.UnMerge();
                xlrange.Range["C3:C51"].Cells.UnMerge();
                xlrange.Range["D3:D51"].Cells.UnMerge();
                xlrange.Range["E3:E51"].Cells.UnMerge();
                xlrange.Range["F3:F51"].Cells.UnMerge();
                xlrange.Range["G3:G51"].Cells.UnMerge();
                xlrange.Range["H3:H51"].Cells.UnMerge(); 
                xlrange.Range["I3:I51"].Cells.UnMerge();
                xlrange.Range["J3:J51"].Cells.UnMerge();
                xlrange.Range["K3:K51"].Cells.UnMerge();








                xlworkbook.Close();
                xlapp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

       

        private void panel1_Paint(object sender, PaintEventArgs e)
        {
            
        }
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }
        protected override void WndProc(ref Message m)
        {
            const int WM_NCCALCSIZE = 0X0083;
            const int WM_NCHITTEST = 0x0084;//Win32, Mouse Input Notification: Determine what part of the window corresponds to a point, allows to resize the form.
            const int resizeAreaSize = 10;

            const int HTCLIENT = 1; //Represents the client area of the window
            const int HTLEFT = 10;  //Left border of a window, allows resize horizontally to the left
            const int HTRIGHT = 11; //Right border of a window, allows resize horizontally to the right
            const int HTTOP = 12;   //Upper-horizontal border of a window, allows resize vertically up
            const int HTTOPLEFT = 13;//Upper-left corner of a window border, allows resize diagonally to the left
            const int HTTOPRIGHT = 14;//Upper-right corner of a window border, allows resize diagonally to the right
            const int HTBOTTOM = 15; //Lower-horizontal border of a window, allows resize vertically down
            const int HTBOTTOMLEFT = 16;//Lower-left corner of a window border, allows resize diagonally to the left
            const int HTBOTTOMRIGHT = 17;//Lower-right corner of a window border, allows resize diagonally to the right
            if (m.Msg == WM_NCHITTEST)
            { //If the windows m is WM_NCHITTEST
                base.WndProc(ref m);
                if (this.WindowState == FormWindowState.Normal)//Resize the form if it is in normal state
                {
                    if ((int)m.Result == HTCLIENT)//If the result of the m (mouse pointer) is in the client area of the window
                    {
                        Point screenPoint = new Point(m.LParam.ToInt32()); //Gets screen point coordinates(X and Y coordinate of the pointer)                           
                        Point clientPoint = this.PointToClient(screenPoint); //Computes the location of the screen point into client coordinates                          
                        if (clientPoint.Y <= resizeAreaSize)//If the pointer is at the top of the form (within the resize area- X coordinate)
                        {
                            if (clientPoint.X <= resizeAreaSize) //If the pointer is at the coordinate X=0 or less than the resizing area(X=10) in 
                                m.Result = (IntPtr)HTTOPLEFT; //Resize diagonally to the left
                            else if (clientPoint.X < (this.Size.Width - resizeAreaSize))//If the pointer is at the coordinate X=11 or less than the width of the form(X=Form.Width-resizeArea)
                                m.Result = (IntPtr)HTTOP; //Resize vertically up
                            else //Resize diagonally to the right
                                m.Result = (IntPtr)HTTOPRIGHT;
                        }
                        else if (clientPoint.Y <= (this.Size.Height - resizeAreaSize)) //If the pointer is inside the form at the Y coordinate(discounting the resize area size)
                        {
                            if (clientPoint.X <= resizeAreaSize)//Resize horizontally to the left
                                m.Result = (IntPtr)HTLEFT;
                            else if (clientPoint.X > (this.Width - resizeAreaSize))//Resize horizontally to the right
                                m.Result = (IntPtr)HTRIGHT;
                        }
                        else
                        {
                            if (clientPoint.X <= resizeAreaSize)//Resize diagonally to the left
                                m.Result = (IntPtr)HTBOTTOMLEFT;
                            else if (clientPoint.X < (this.Size.Width - resizeAreaSize)) //Resize vertically down
                                m.Result = (IntPtr)HTBOTTOM;
                            else //Resize diagonally to the right
                                m.Result = (IntPtr)HTBOTTOMRIGHT;
                        }
                    }
                }
                return;
            }
            if (m.Msg == WM_NCCALCSIZE && m.WParam.ToInt32() == 1)
            {
                return;
            }
            base.WndProc(ref m);
        }

        private void iconButton7_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void iconButton9_Click(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Normal)
                WindowState = FormWindowState.Maximized;
            else
                WindowState = FormWindowState.Normal;
        }

        private void iconButton8_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void iconButton10_Click(object sender, EventArgs e)
        {
            ActivateButton(sender, RGBColors.color1);
            Report r1 = new Report();
            r1.Show();

        }

        private void Form3_Resize(object sender, EventArgs e)
        {
            AdjustForm();
        }

        private void AdjustForm()
        {
            switch (this.WindowState)
            {
                case FormWindowState.Maximized:
                    this.Padding = new Padding(8, 8, 8, 0);
                    break;
                case FormWindowState.Normal: 
                    if (this.Padding.Top != borderSize)
                        this.Padding = new Padding(borderSize);
                    break;
            }
        }

        private void FileName_LBL_Click(object sender, EventArgs e)
        {

        }

        private void iconCurrentChildForm_Click(object sender, EventArgs e)
        {

        }

        private void lblTitleChildForm_Click(object sender, EventArgs e)
        {

        }
    }
    }
    

