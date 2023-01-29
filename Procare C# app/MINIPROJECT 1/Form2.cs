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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form4_Load(object sender, EventArgs e)
        {
            int w = Screen.PrimaryScreen.Bounds.Width;
            int h = Screen.PrimaryScreen.Bounds.Height;
            this.Location = new Point(0, 0);
            this.Size = new Size(w, h);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
                String name = textBox1.Text;
                int age = int.Parse(textBox2.Text);
                long n = long.Parse(textBox3.Text);
                String address = textBox4.Text;


                OleDbConnection con = new OleDbConnection("Provider=Microsoft.Ace.OLEDB.12.0;Data Source='D:\\PHARMACY.xlsx';Extended Properties=Excel 8.0;");
                con.Open();
                OleDbCommand cmd = new OleDbCommand();
                String sql = "insert into[sheet1$](CUSTOMER_NAME,AGE,PH_NO,ADDRESS)values('" + name + "'," + age + "," + n + ",'" + address + "')";
                cmd.Connection = con;
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();
                con.Close();

                Form3 f3 = new Form3();
                f3.passingvalue = textBox1.Text;
                f3.ShowDialog();


                
                this.Hide();
                f3.ShowDialog();
                this.Close();
                
                
               
         
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
