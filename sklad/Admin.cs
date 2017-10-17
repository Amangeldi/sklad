using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace sklad
{
    public partial class Admin : Form
    {
        string FIO = null;
        int Uid;
        public Admin(int _id)
        {
            InitializeComponent();
            this.Uid = _id;
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void Admin_Load(object sender, EventArgs e)
        {
            ConnOpen admLoad = new ConnOpen();
            admLoad.connection.Open();
            SqlCommand sqlCom = new SqlCommand("SELECT * FROM dbo.users WHERE user_id LIKE '%" + Uid + "'", admLoad.connection);
            SqlDataReader dr = sqlCom.ExecuteReader();
            dr.Read();
            FIO = dr["user_familija"].ToString() + " " + dr["user_imja"].ToString() + " " + dr["user_otchestvo"].ToString();
            label1.Text = "Здравствуйте " + FIO;
            admLoad.connection.Close();
            //-------
            ConnOpen waybLoad = new ConnOpen();
            waybLoad.connection.Open();
            SqlDataAdapter adapter = new SqlDataAdapter("SELECT DISTINCT waybill FROM dbo.Responsibility WHERE traffic = '0'", waybLoad.connection);
            DataTable tbl = new DataTable();
            adapter.Fill(tbl);

            comboBox1.DataSource = tbl;
            comboBox1.DisplayMember = "waybill";
            comboBox1.ValueMember = "waybill";
            waybLoad.connection.Close();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Users f = new Users();
            f.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Products f = new Products();
            f.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Responsibility f = new Responsibility();
            f.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Report f1 = new Report();
            f1.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Waybill f;
            if(radioButton1.Checked==true)
            {
                f = new Waybill(comboBox1.SelectedValue.ToString(), "0");
            }
            else
            {
                f = new Waybill(comboBox1.SelectedValue.ToString(), "1");
            }
            f.ShowDialog();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if(radioButton1.Checked == true)
            {
                ConnOpen waybLoad = new ConnOpen();
                waybLoad.connection.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("SELECT DISTINCT waybill FROM dbo.Responsibility WHERE traffic = '0'", waybLoad.connection);
                DataTable tbl = new DataTable();
                adapter.Fill(tbl);

                comboBox1.DataSource = tbl;
                comboBox1.DisplayMember = "waybill";
                comboBox1.ValueMember = "waybill";
                waybLoad.connection.Close();
            }
            else
            {
                ConnOpen waybLoad = new ConnOpen();
                waybLoad.connection.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("SELECT DISTINCT waybill FROM dbo.Responsibility WHERE traffic = '1'", waybLoad.connection);
                DataTable tbl = new DataTable();
                adapter.Fill(tbl);

                comboBox1.DataSource = tbl;
                comboBox1.DisplayMember = "waybill";
                comboBox1.ValueMember = "waybill";
                waybLoad.connection.Close();
            }
        }
    }
}
