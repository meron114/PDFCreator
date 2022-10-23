using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PDFCreator
{
    public partial class Form2 : Form
    {
        public Form2(string jpg, string r)
        {
            InitializeComponent();
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            if (r == "")
            {
                PictureBox1.ImageLocation = jpg + ".jpg";
            }
            else
            {
                PictureBox1.ImageLocation = jpg + "_" + r + ".jpg";
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
