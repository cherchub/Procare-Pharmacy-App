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
    public partial class Form7 : Form
    {
        public Form7()
        {
            InitializeComponent();
        }

        private void Form7_Load(object sender, EventArgs e)
        {
            int w = Screen.PrimaryScreen.Bounds.Width;
            int h = Screen.PrimaryScreen.Bounds.Height;
            this.Location = new Point(0, 0);
            this.Size = new Size(w, h);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String username, password;
            username = textBox1.Text;
            password = textBox2.Text;
            if ((username == "procare_owner") && (password == "owner123"))
            {
                Form8 f8 = new Form8();
                this.Hide();
                f8.ShowDialog();
                this.Close();
                
            }
            else
            {
                MessageBox.Show("*Wrong username or password !!* Please try again");
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form1 f1 = new Form1();
            f1.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
