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
    public partial class Edit_product : Form
    {
        int Uid;
        public Edit_product(int _id)
        {
            InitializeComponent();
            this.Uid = _id;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int balans = 0;
            if (checkBox1.Checked == true)
            {
                balans = 1;
            }
            Product p = new Product();
            p.update(Uid, textBox1.Text, textBox2.Text, float.Parse(textBox3.Text), float.Parse(textBox4.Text),Convert.ToInt32(comboBox1.SelectedValue), textBox5.Text, "", dateTimePicker1.Value.ToString(), dateTimePicker2.Value.ToString(), textBox6.Text, balans);
            MessageBox.Show("Материал изменен", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Edit_product_Load(object sender, EventArgs e)
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
            //---------------------
            ConnOpen EPL = new ConnOpen();
            EPL.connection.Open();
            SqlCommand command = new SqlCommand("SELECT * FROM dbo.Product WHERE product_id LIKE '"+Uid+"'", EPL.connection);
            SqlDataReader reader = command.ExecuteReader();
            while(reader.Read())
            {
                textBox1.Text = reader["product_kod"].ToString();
                textBox2.Text = reader["product_name"].ToString();
                textBox3.Text = reader["product_price"].ToString();
                textBox4.Text = reader["product_quantity"].ToString();
                textBox5.Text = reader["product_description"].ToString();
                dateTimePicker1.Value = Convert.ToDateTime(reader["receipt_date"]);
                textBox6.Text = reader["location"].ToString();
                if(reader["balans"].ToString()=="1")
                {
                    checkBox1.Checked = true;
                }
                comboBox1.SelectedValue = reader["product_unit"].ToString();
            }
            EPL.connection.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Unit f = new Unit();
            f.ShowDialog();
        }
    }
}
