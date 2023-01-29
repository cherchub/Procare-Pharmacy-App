using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MINIPROJECT_1
{
    public partial class Form9 : Form
    {
        public Form9()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form1 f1 = new Form1();
            this.Hide();
            f1.ShowDialog();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form10 f10 = new Form10();
            this.Hide();
            f10.ShowDialog();
            this.Close();
        
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form11 f11 = new Form11();
            this.Hide();
            f11.ShowDialog();
            this.Close();
        }

        private void Form9_Load(object sender, EventArgs e)
        {
            int w = Screen.PrimaryScreen.Bounds.Width;
            int h = Screen.PrimaryScreen.Bounds.Height;
            this.Location = new Point(0, 0);
            this.Size = new Size(w, h);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
