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
using ClosedXML.Excel;

namespace FileBrowser
{
    public partial class Form2 : Form
    {
        public Form2(string path)
        {
            InitializeComponent();
            this.pictureBox1.Image = Image.FromFile(path);
        }

        private void pictureBox1_Paint(object sender, PaintEventArgs e)
        {
            Image image = this.pictureBox1.Image;
            if (image != null)
            {
/*                // 画像を(0, 0)に描画する
                e.Graphics.DrawImage(image, 0, 0, image.Width, image.Height);*/

                // 画像サイズをコントロールのサイズに合うよう自動で調整する
                this.pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
                
            }
        }
    }
}
