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
    public partial class Responsibility : Form
    {
        public Responsibility()
        {
            InitializeComponent();
        }

        private void Responsibility_Load(object sender, EventArgs e)
        {
            ConnOpen respLoad = new ConnOpen();
            respLoad.connection.Open();
            //SqlCommand command = new SqlCommand("SELECT * FROM dbo.Users", respLoad.connection);
            SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM dbo.Users WHERE role='2'", respLoad.connection);
            DataTable tbl = new DataTable();
            adapter.Fill(tbl);

            comboBox1.DataSource = tbl;
            comboBox1.DisplayMember = "fio";
            comboBox1.ValueMember = "user_id";
            respLoad.connection.Close();
            //----------
            respLoad.connection.Open();
            SqlDataAdapter adapter_p = new SqlDataAdapter("SELECT * FROM dbo.Product", respLoad.connection);
            DataTable tbl_p = new DataTable();
            adapter_p.Fill(tbl_p);

            comboBox2.DataSource = tbl_p;
            comboBox2.DisplayMember = "product_name";
            comboBox2.ValueMember = "product_id";
            respLoad.connection.Close();
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            ConnOpen add_responsible = new ConnOpen();
            ConnOpen update_product = new ConnOpen();
            ConnOpen update_user = new ConnOpen();
            add_responsible.connection.Open();
            string sql = string.Format("Insert Into Responsibility" +
                       "(responsible, location, product, product_quantity, date, waybill, traffic) Values(@responsible, @location, @product, @product_quantity, @date, @waybill, @traffic)");
            using (SqlCommand cmd = new SqlCommand(sql, add_responsible.connection))
            {
                cmd.Parameters.AddWithValue("@responsible", comboBox1.SelectedValue.ToString());
                cmd.Parameters.AddWithValue("@location", textBox1.Text);
                cmd.Parameters.AddWithValue("@product", comboBox2.SelectedValue.ToString());
                cmd.Parameters.AddWithValue("@product_quantity", textBox2.Text);
                cmd.Parameters.AddWithValue("@date", dateTimePicker1.Value.ToString());
                cmd.Parameters.AddWithValue("@waybill", textBox3.Text);
                cmd.Parameters.AddWithValue("@traffic", comboBox3.SelectedIndex.ToString());
                cmd.ExecuteNonQuery();
            }
            add_responsible.connection.Close();
            update_product.connection.Open();
            string sqlProd = string.Format("Update Product Set product_flag = '1', [last_date] = '" + dateTimePicker1.Value.ToString() + "' WHERE product_id = '" + comboBox2.SelectedValue.ToString()+"' " );
            using (SqlCommand cmd = new SqlCommand(sqlProd, update_product.connection))
            {
                cmd.ExecuteNonQuery();
            }
            update_product.connection.Close();
            
            update_user.connection.Open();
            string sqlUser = "";

            if (comboBox3.SelectedIndex.ToString() == "0")
            {
                sqlUser = string.Format("Update Users Set prih = '1' WHERE user_id = '" + comboBox1.SelectedValue.ToString()+"'");
                using (SqlCommand cmd = new SqlCommand(sqlUser, update_user.connection))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            else if(comboBox3.SelectedIndex.ToString() == "1")
            {
                sqlUser = string.Format("Update Users Set rash = '1' WHERE user_id = '" + comboBox1.SelectedValue.ToString() + "'");
                using (SqlCommand cmd = new SqlCommand(sqlUser, update_user.connection))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            update_user.connection.Close();
            MessageBox.Show("Ok", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }
    }
}
