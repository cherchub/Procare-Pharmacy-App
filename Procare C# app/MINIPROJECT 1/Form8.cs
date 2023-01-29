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
    public partial class Form8 : Form
    {
        public Form8()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Ace.OLEDB.12.0;Data Source='D:\\PHARMACY.xlsx';Extended Properties=Excel 8.0;");
            con.Open();
            OleDbCommand cmd = new OleDbCommand();
            String qry = "Select * from [sheet1$]";
            OleDbDataAdapter ad = new OleDbDataAdapter(qry, con);
            DataSet ds = new DataSet();
            ad.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            con.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form8_Load(object sender, EventArgs e)
        {
            int w = Screen.PrimaryScreen.Bounds.Width;
            int h = Screen.PrimaryScreen.Bounds.Height;
            this.Location = new Point(0, 0);
            this.Size = new Size(w, h);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form9 f9 = new Form9();
            this.Hide();
            f9.ShowDialog();
            this.Close();
        }
    }
}
