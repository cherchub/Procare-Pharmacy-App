using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace MINIPROJECT_1
{
    public partial class Form11 : Form
    {
        public Form11()
        {
            InitializeComponent();
        }

        private void Form11_Load(object sender, EventArgs e)
        {
            int w = Screen.PrimaryScreen.Bounds.Width;
            int h = Screen.PrimaryScreen.Bounds.Height;
            this.Location = new Point(0, 0);
            this.Size = new Size(w, h);
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form9 f9 = new Form9();
            this.Hide();
            f9.ShowDialog();
            this.Close();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            String name = textBox1.Text;
            long mob = long.Parse(textBox2.Text);
            String mail = textBox3.Text;
            String item = textBox4.Text;
            String cat=" ";

            if (radioButton1.Checked == true)
            {

                cat = "Donate";

            }
            else if (radioButton2.Checked == true)
            {

                cat = "Recycle";

            }
            else if (radioButton3.Checked == true)
            {

                cat = "Sell";

            }
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Ace.OLEDB.12.0;Data Source='D:\\DONATION.xlsx';Extended Properties=Excel 8.0;");
            con.Open();
            OleDbCommand cmd = new OleDbCommand();
            String sql = "insert into[sheet1$] values('" + name + "'," + mob + ",'" + mail + "','" + cat + "','" + item + "')";
            cmd.Connection = con;
            cmd.CommandText = sql;
            cmd.ExecuteNonQuery();
            con.Close();

            Form12 f12 = new Form12();
            this.Hide();
            f12.ShowDialog();
            this.Close();
        
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form13 f13 = new Form13();
            this.Hide();
            f13.ShowDialog();
            this.Close();
        }
    }
}
