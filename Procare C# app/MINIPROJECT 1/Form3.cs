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
    public partial class Form3 : Form
    {

        public string name;
        public string passingvalue
        {

            get { return name; }
            set { name = value; }
        }

        public Form3()
        {
            InitializeComponent();
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {



            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Ace.OLEDB.12.0;Data Source='D:\\PHARMACY.xlsx';Extended Properties=Excel 8.0;");
            con.Open();
            OleDbCommand cmd = new OleDbCommand();

            listBox1.Items.Clear();
            foreach (string s in checkedListBox1.CheckedItems)
                listBox1.Items.Add(s);

            int sum = 0;
            String meds = " ";


            if (checkedListBox1.GetItemCheckState(0) == CheckState.Checked)
            {
                sum = sum + 30;
                textBox1.Text = sum.ToString();
                meds = meds + "Restyl 0.25 - 30/- \n ";
            }

            if (checkedListBox1.GetItemCheckState(1) == CheckState.Checked)
            {
                sum = sum + 20;
                textBox1.Text = sum.ToString();
                meds = meds + "Crocin - 20/- \n ";
            }
            if (checkedListBox1.GetItemCheckState(2) == CheckState.Checked)
            {
                sum = sum + 30;
                textBox1.Text = sum.ToString();
                meds = meds + "Dolo 650 - 30/- \n ";

            }
            if (checkedListBox1.GetItemCheckState(3) == CheckState.Checked)
            {
                sum = sum + 69;
                textBox1.Text = sum.ToString();
                meds = meds + "Roseday 5 - 69/- \n ";
            }
            if (checkedListBox1.GetItemCheckState(4) == CheckState.Checked)
            {
                sum = sum + 30;
                textBox1.Text = sum.ToString();
                meds = meds + "Evion 400 - 30/- \n ";
            }
            if (checkedListBox1.GetItemCheckState(5) == CheckState.Checked)
            {
                sum = sum + 30;
                textBox1.Text = sum.ToString();
                meds = meds + "Combiflam - 30/- \n ";
            }
            if (checkedListBox1.GetItemCheckState(6) == CheckState.Checked)
            {
                sum = sum + 99;
                textBox1.Text = sum.ToString();
                meds = meds + "TeszlocAM - 99/- \n ";
            }
            if (checkedListBox1.GetItemCheckState(7) == CheckState.Checked)
            {
                sum = sum + 100;
                textBox1.Text = sum.ToString();
                meds = meds + "Shelcal 500 - 100/- \n ";
            }
            if (checkedListBox1.GetItemCheckState(8) == CheckState.Checked)
            {
                sum = sum + 40;
                textBox1.Text = sum.ToString();
                meds = meds + "Eldoper - 40/- \n ";
            }
            if (checkedListBox1.GetItemCheckState(9) == CheckState.Checked)
            {
                sum = sum + 80;
                textBox1.Text = sum.ToString();
                meds = meds + "Limcee - 80/- \n ";
            }
            if (checkedListBox1.GetItemCheckState(10) == CheckState.Checked)
            {
                sum = sum + 34;
                textBox1.Text = sum.ToString();
                meds = meds + "Becosules - 34/- \n ";
            }
            if (checkedListBox1.GetItemCheckState(11) == CheckState.Checked)
            {
                sum = sum + 150;
                textBox1.Text = sum.ToString();
                meds = meds + "CNX dox - 150/- \n ";
            }
            if (checkedListBox1.GetItemCheckState(12) == CheckState.Checked)
            {
                sum = sum + 350;
                textBox1.Text = sum.ToString();
                meds = meds + "Ivernew 12 - 350/- \n ";
            }
            if (checkedListBox1.GetItemCheckState(13) == CheckState.Checked)
            {
                sum = sum + 200;
                textBox1.Text = sum.ToString();
                meds = meds + "Augmentin 625 - 200/- \n ";
            }
            if (checkedListBox1.GetItemCheckState(14) == CheckState.Checked)
            {
                sum = sum + 10;
                textBox1.Text = sum.ToString();
                meds = meds + "Zentel - 10/- \n ";
            }
            if (checkedListBox1.GetItemCheckState(15) == CheckState.Checked)
            {
                sum = sum + 138;
                textBox1.Text = sum.ToString();
                meds = meds + "Pan 40 - 138/- \n ";
            }
            if (checkedListBox1.GetItemCheckState(16) == CheckState.Checked)
            {
                sum = sum + 118;
                textBox1.Text = sum.ToString();
                meds = meds + "Azee 500 - 118/- \n ";
            }
            if (checkedListBox1.GetItemCheckState(17) == CheckState.Checked)
            {
                sum = sum + 100;
                textBox1.Text = sum.ToString();
                meds = meds + "Tenlimac 20 - 100/- \n ";
            }
            if (checkedListBox1.GetItemCheckState(18) == CheckState.Checked)
            {
                sum = sum + 120;
                textBox1.Text = sum.ToString();
                meds = meds + "CTD 12.5 - 120/- \n ";
            }

            
            String sql = "update [sheet1$] set MEDICINES='" + meds + "',TOTAL_AMT='" + sum + "' where CUSTOMER_NAME='" + name + "'";

            cmd.Connection = con;
            cmd.CommandText = sql;
            cmd.ExecuteNonQuery();
            con.Close();
            
                
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form4 f4 = new Form4();
           
            f4.passingvalue = name;
            f4.ShowDialog();
            this.Hide();
            this.Close();
           
           
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            int w = Screen.PrimaryScreen.Bounds.Width;
            int h = Screen.PrimaryScreen.Bounds.Height;
            this.Location = new Point(0, 0);
            this.Size = new Size(w, h);
        }
    }
}
