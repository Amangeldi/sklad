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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            int role = 5;
            if (radioButton1.Checked == true)
            {
                role = 1;
            }
            else if (radioButton2.Checked == true)
            {
                role = 2;
            }
            else if (radioButton3.Checked == true)
            {
                role = 3;
            }
            else
            {
                role = 5;
            }
            string login = textBox1.Text;
            string password = textBox2.Text;
            ConnOpen loginConnection = new ConnOpen();
            loginConnection.connection.Open();
            SqlCommand sqlCom = new SqlCommand("SELECT * FROM dbo.users WHERE role LIKE '%" + role + "' and user_login LIKE '%" + login + "'and user_password LIKE '%" + password + "'", loginConnection.connection);
            SqlDataReader dr = sqlCom.ExecuteReader();
            int id;
            dr.Read();
            if (dr.HasRows == true)
            {
                id = Convert.ToInt32(dr["user_id"]);
                if (role == 1)
                {
                    Admin f1 = new Admin(id);
                    f1.ShowDialog();
                }
                else if (role == 2)
                {
                    Inspector f2 = new Inspector(id);
                    f2.ShowDialog();
                }
                else if (role == 3)
                {
                    Responsible f3 = new Responsible(id);
                    f3.ShowDialog();
                }
            }
            else
            {
                this.Text = "Не войдете";
            }
            
        }
    }
}
