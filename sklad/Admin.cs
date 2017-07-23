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
    }
}
