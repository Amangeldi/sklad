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
        public ConnOpen conn = new ConnOpen();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            conn.connection.Open();
            SqlCommand comm = new SqlCommand("SELECT * FROM dbo.Unit",conn.connection);
            SqlDataReader reader = comm.ExecuteReader();
            while(reader.Read())
            {
                this.Text = reader["unit_name"].ToString();
            }
            conn.connection.Close();
        }
    }
}
