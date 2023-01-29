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
    public partial class Form4 : Form
    {
        public string name;
        public string passingvalue
        {

            get { return name; }
            set { name = value; }
        }  
       

        public Form4()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           try
           {

               OleDbConnection con = new OleDbConnection("Provider=Microsoft.Ace.OLEDB.12.0;Data Source='D:\\PHARMACY.xlsx';Extended Properties=Excel 8.0;");
            con.Open();
            OleDbCommand cmd = new OleDbCommand();
            String mode=" ";
            
            if (radioButton1.Checked == true)
            {
                
                mode = "Credit/Debit Card";
                
            }
            else if (radioButton2.Checked == true)
            {

                mode = "Cash On Delivery";
                
            }

            

            String sql = "update [sheet1$] set MODE_OF_PAYMENT='" + mode + "' where CUSTOMER_NAME='" + name + "'";

            cmd.Connection = con;
            cmd.CommandText = sql;
            cmd.ExecuteNonQuery();
            con.Close();
            if (mode == "Credit/Debit Card")
            {
                Form5 f5 = new Form5();
                this.Hide();
                f5.ShowDialog();
                this.Close();
            }
            else if (mode == "Cash On Delivery")
            {
                Form6 f6 = new Form6();
                this.Hide();
                f6.ShowDialog();
                this.Close();
            }

           }

               catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void Form4_Load(object sender, EventArgs e)
        {
            int w = Screen.PrimaryScreen.Bounds.Width;
            int h = Screen.PrimaryScreen.Bounds.Height;
            this.Location = new Point(0, 0);
            this.Size = new Size(w, h);
        }
    }
}
