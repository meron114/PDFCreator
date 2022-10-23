using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PDFCreator
{
    public partial class Form3 : Form
    {
        public Form1 f1; //親フォーム

        public Form3()
        {
            InitializeComponent();
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            Random r = new System.Random();
            int result = r.Next(0, 3);
            if (result == 0) { PictureBox1.Image = Properties.Resources.loading_blue; }
            else if (result == 1) { PictureBox1.Image = Properties.Resources.loading_red; }
            else if (result == 2) { PictureBox1.Image = Properties.Resources.loading_green; }
            PictureBox1.Paint += new PaintEventHandler(Form3_Paint);
            ImageAnimator.Animate(PictureBox1.Image, new EventHandler(this.Image_FrameChanged));
        }
        private void Image_FrameChanged(object o, EventArgs e)
        {
            PictureBox1.Invalidate();
        }

        private void Form3_Paint(object sender, PaintEventArgs e)
        {
            ImageAnimator.UpdateFrames(PictureBox1.Image);
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (f1.cancelTokensource != null)
            {
                f1.cancelTokensource.Cancel();//キャンセルを発行
            }
        }
    }
}