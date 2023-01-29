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
    public partial class Form10 : Form
    {
        public Form10()
        {
            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void Form10_Load(object sender, EventArgs e)
        {
            int w = Screen.PrimaryScreen.Bounds.Width;
            int h = Screen.PrimaryScreen.Bounds.Height;
            this.Location = new Point(0, 0);
            this.Size = new Size(w, h);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form9 f9 = new Form9();
            this.Hide();
            f9.ShowDialog();
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                String name = textBox1.Text;
                int age = int.Parse(textBox2.Text);
                long n = long.Parse(textBox3.Text);
                String address = textBox4.Text;
                string dr = listBox1.SelectedItem.ToString();

                OleDbConnection con = new OleDbConnection("Provider=Microsoft.Ace.OLEDB.12.0;Data Source='D:\\APPOINTMENT.xlsx';Extended Properties=Excel 8.0;");
                con.Open();
                OleDbCommand cmd = new OleDbCommand();
                String sql = "insert into[sheet1$] values('" + name + "'," + age + "," + n + ",'" + address + "','" + this.dateTimePicker1.Text + "','" + dr + "');";
                cmd.Connection = con;
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Form14 f14 = new Form14();
            this.Hide();
            f14.ShowDialog();
            this.Close();
        
        }
    }
}
