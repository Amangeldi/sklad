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
    public partial class Unit : Form
    {
        public Unit()
        {
            InitializeComponent();
        }

        private void Unit_Load(object sender, EventArgs e)
        {
            ConnOpen unitsLoad = new ConnOpen();
            SqlDataAdapter adapter = new SqlDataAdapter();
            unitsLoad.connection.Open();
            SqlCommand sqlCom = new SqlCommand("SELECT * FROM dbo.unit", unitsLoad.connection);
            unitsLoad.connection.Close();
            adapter.SelectCommand = sqlCom;
            DataSet dataset = new DataSet();
            adapter.Fill(dataset);
            dataGridView1.DataSource = dataset.Tables[0];
            adapter.Update(dataset);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ConnOpen add_unit = new ConnOpen();
            add_unit.connection.Open();
            string sql = string.Format("Insert Into Unit" +
                       "(unit_name, unit_description) Values(@unit_name, @unit_description)");
            using (SqlCommand cmd = new SqlCommand(sql, add_unit.connection))
            {
                cmd.Parameters.AddWithValue("@unit_name", textBox1.Text);
                cmd.Parameters.AddWithValue("@unit_description", textBox2.Text);
                cmd.ExecuteNonQuery();
            }
            add_unit.connection.Close();
        }
    }
}
