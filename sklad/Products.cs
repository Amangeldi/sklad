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
    public partial class Products : Form
    {
        public Products()
        {
            InitializeComponent();
        }

        private void Products_Load(object sender, EventArgs e)
        {
            ConnOpen productsLoad = new ConnOpen();
            SqlDataAdapter adapter = new SqlDataAdapter();
            productsLoad.connection.Open();
            SqlCommand sqlCom = new SqlCommand("SELECT * FROM dbo.product", productsLoad.connection);
            productsLoad.connection.Close();
            adapter.SelectCommand = sqlCom;
            DataSet dataset = new DataSet();
            adapter.Fill(dataset);
            dataGridView1.DataSource = dataset.Tables[0];
            adapter.Update(dataset);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Add_product f = new Add_product();
            f.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Product p = new Product();
            bool test = p.test_id(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            if (test == true)
            {
                MessageBox.Show("Материал с таким id не существует или был удален", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                
            }
            else
            {
                p.delete(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                MessageBox.Show("Удален продукт " + dataGridView1.CurrentRow.Cells[0].Value.ToString(), "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Product p = new Product();
            bool test = p.test_id(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            if (test == true)
            {
                MessageBox.Show("Материал с таким id не существует или был удален", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            else
            {
                Edit_product f = new Edit_product(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value));
                f.ShowDialog();
            }
        }
    }
}
