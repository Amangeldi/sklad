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
    public partial class Add_product : Form
    {
        public Add_product()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int balans = 0;
            if(checkBox1.Checked == true)
            {
                balans = 1;
            }
            Product p = new Product();
            p.add(textBox1.Text, textBox2.Text,float.Parse(textBox3.Text), float.Parse(textBox4.Text), Convert.ToInt32(comboBox1.SelectedValue), textBox5.Text, "", dateTimePicker1.Value.ToString(), dateTimePicker2.Value.ToString(), textBox6.Text, balans);
            MessageBox.Show("Материал добавлен", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Unit f = new Unit();
            f.ShowDialog();
        }

        private void Add_product_Load(object sender, EventArgs e)
        {
            ConnOpen productLoad = new ConnOpen();
            productLoad.connection.Open();
            SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM dbo.Unit", productLoad.connection);
            DataTable tbl = new DataTable();
            adapter.Fill(tbl);

            comboBox1.DataSource = tbl;
            comboBox1.DisplayMember = "unit_name";
            comboBox1.ValueMember = "unit_id";
            productLoad.connection.Close();
        }
    }
}
