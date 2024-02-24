using Stream;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Stream_25percent
{
    public partial class Second_Page : Form
    {
        public Second_Page()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Second_Page_Load(object sender, EventArgs e)
        {

        }  

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Beam_Schedule BS = new Beam_Schedule();
            BS.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Beam_Layout BL = new Beam_Layout();
            BL.Show();
        }
    }
}
