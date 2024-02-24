using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using ComponentFactory.Krypton.Toolkit;
using Spire.Pdf;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.IO;
using Point = System.Drawing.Point;
using Stream_25percent;

namespace Stream
{
    public partial class Beam_Layout : KryptonForm
    {
        public Beam_Layout()
        {
            InitializeComponent();
        }
        string path = "";
        Microsoft.Office.Interop.Word.Application app;
        Microsoft.Office.Interop.Word.Document doc;
        object objMiss = Missing.Value;
        object TmpFile = System.IO.Path.GetTempFileName() + "ConvertW.pdf";
        object Filelocation = @"C:\Users\renzc\source\repos\Stream 25percent - Copy\bin\Debug\ConvertW.docx";

        private void Beam_Layout_Load(object sender, EventArgs e)
        {

        }

        private void axAcroPDF1_Enter(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            lblText.Text = "The Beam Layout from STAAD RCDC";
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "PDF |*.pdf";


            if (ofd.ShowDialog() == DialogResult.OK)
            {
                axAcroPDF1.src = ofd.FileName;
                path = ofd.FileName;
                filenamepdf.Text = path;
            }
        }

       

      

        

        private void lblText_Click(object sender, EventArgs e)
        {

        }

        

        

        private void button1_Click_1(object sender, EventArgs e)
        {
            lblText.Text = "Conversion of PDF file to Word file";
            try
            {
                PdfDocument obj = new PdfDocument();
                obj.LoadFromFile(path);
                obj.SaveToFile("ConvertW.docx", FileFormat.DOCX);

                if (File.Exists("ConvertW.docx") == true)
                {
                    Process.Start("ConvertW.docx");


                }
            }
            catch (Exception ext)
            {
                System.Console.WriteLine(ext.Message);
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            lblText.Text = "Simplified Beam Layout";
            try
            {
                app = new Microsoft.Office.Interop.Word.Application();
                doc = app.Documents.Open(ref Filelocation, ref objMiss);

                
                doc.ExportAsFixedFormat(TmpFile.ToString(), Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
                this.axAcroPDF1.src = TmpFile.ToString();
                this.axAcroPDF1.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
            finally
            {
                doc.Close(WdSaveOptions.wdDoNotSaveChanges, WdOriginalFormat.wdOriginalDocumentFormat, true);
                app.Quit(WdSaveOptions.wdDoNotSaveChanges);
            }
        }

        private void button1_MouseDown(object sender, MouseEventArgs e)
        {
            button1.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void button1_MouseHover(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromArgb(80, 80, 80);
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromArgb(30, 27, 46);
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

        private void iconButton6_Click_1(object sender, EventArgs e)
        {
            Second_Page page = new Second_Page();
            page.Show();
            this.Hide();
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
