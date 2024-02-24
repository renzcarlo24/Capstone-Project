using FontAwesome.Sharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Pdf;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.IO;
using Point = System.Drawing.Point;

namespace Stream_25percent
{
    public partial class Form2 : Form
    {
       

        public Form2()
        {
            InitializeComponent();
            
        }
        string path = "";
        Microsoft.Office.Interop.Word.Application app;
        Microsoft.Office.Interop.Word.Document doc;
        object objMiss = Missing.Value;
        object TmpFile = System.IO.Path.GetTempFileName() + "ConvertW.pdf";
        object Filelocation = @"C:\Users\renzc\source\repos\Stream 25percent - Copy\bin\Debug\ConvertW.docx";
       
       


        private void filenamepdf_Click(object sender, EventArgs e)
        {
           
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "PDF |*.pdf";


            if (ofd.ShowDialog() == DialogResult.OK)
            {
                axAcroPDF1.src = ofd.FileName;
                path = ofd.FileName;
                Filename.Text = path;
            }
        }

        private void back_Click(object sender, EventArgs e)
        {
            Second_Page page = new Second_Page();
            page.Show();
            this.Hide();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void iconButton1_Click(object sender, EventArgs e)
        {
            Second_Page page = new Second_Page();
            page.Show();
            this.Hide();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        

        private void panelMenu_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void Filename_Click(object sender, EventArgs e)
        {

        }

        private void iconButton8_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void iconButton9_Click(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Normal)
                WindowState = FormWindowState.Maximized;
            else
                WindowState = FormWindowState.Normal;
        }

        private void iconButton7_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void iconCurrentChildForm_Click(object sender, EventArgs e)
        {

        }

        private void iconButton2_Click(object sender, EventArgs e)
        {
            
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

        private void iconButton3_Click(object sender, EventArgs e)
        {
            
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
        private void FindAndReplace(object Findtext, object ReplaceText)
        {
            this.app.Selection.Find.Execute(ref Findtext, true, true, false, false, false, true, false, 1,
                ref ReplaceText, 2, false, false, false, false);
        }

        private void iconButton4_Click(object sender, EventArgs e)
        {
            
        }

        private void iconButton5_Click(object sender, EventArgs e)
        {
            
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            
        }
       

        private void Form2_Resize(object sender, EventArgs e)
        {
            
        }
       

        private void axAcroPDF1_Enter(object sender, EventArgs e)
        {

        }
    }
}
