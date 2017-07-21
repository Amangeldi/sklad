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
    public partial class Users : Form
    {
        public Users()
        {
            InitializeComponent();
        }

        private void Users_Load(object sender, EventArgs e)
        {
            ConnOpen usersLoad = new ConnOpen();
            SqlDataAdapter adapter = new SqlDataAdapter();
            usersLoad.connection.Open();
            SqlCommand sqlCom = new SqlCommand("SELECT * FROM dbo.users", usersLoad.connection);
            usersLoad.connection.Close();
            adapter.SelectCommand = sqlCom;
            DataSet dataset = new DataSet();
            adapter.Fill(dataset);
            dataGridView1.DataSource = dataset.Tables[0];
            adapter.Update(dataset);
        }
    }
}
