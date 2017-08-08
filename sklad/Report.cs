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
    public partial class Report : Form
    {
        public Report()
        {
            InitializeComponent();
        }
        int responsible, product, traffic, unit;
        string sResponsible, sProduct, sUnit, location, date, waybill;
        float product_quantity, price, rProduct_quantity;
        private void Report_Load(object sender, EventArgs e)
        {
            ConnOpen reportLoad = new ConnOpen();
            ConnOpen productLoad = new ConnOpen();
            ConnOpen unitLoad = new ConnOpen();
            ConnOpen userLoad = new ConnOpen();
            reportLoad.connection.Open();
            productLoad.connection.Open();
            unitLoad.connection.Open();
            userLoad.connection.Open();
            //Открыли все коннекты
            SqlCommand commandResponsible = new SqlCommand("SELECT * FROM dbo.Responsibility", reportLoad.connection);
            SqlDataReader readerResponsible = commandResponsible.ExecuteReader();
            SqlCommand commandProduct;
            SqlDataReader readerProduct;
            SqlCommand commandUnit;
            SqlDataReader readerUnit;
            SqlCommand commandUser;
            SqlDataReader readerUser;
            //Создали команды и датаридеры
            var columnPName = new DataGridViewColumn();
            columnPName.HeaderText = "Название";
            columnPName.Name = "productName";
            columnPName.CellTemplate = new DataGridViewTextBoxCell();
            
            var columnPUnit = new DataGridViewColumn();
            columnPUnit.HeaderText = "Ед. Изм.";
            columnPUnit.Name = "productUnit";
            columnPUnit.CellTemplate = new DataGridViewTextBoxCell();

            var columnPPrice = new DataGridViewColumn();
            columnPPrice.HeaderText = "Цена";
            columnPPrice.Name = "productPrice";
            columnPPrice.CellTemplate = new DataGridViewTextBoxCell();

            dataGridView1.Columns.Add(columnPName);
            dataGridView1.Columns.Add(columnPUnit);
            dataGridView1.Columns.Add(columnPPrice);
            //Добавили постоянные колонки

            while (readerResponsible.Read())
            {
                product = Convert.ToInt32(readerResponsible["product"]);
                responsible = Convert.ToInt32(readerResponsible["responsible"]);
                traffic = Convert.ToInt32(readerResponsible["traffic"]);
                waybill = readerResponsible["waybill"].ToString();
                rProduct_quantity = Convert.ToSingle(readerResponsible["product_quantity"]);
                //-----
                commandProduct = new SqlCommand("SELECT * FROM dbo.Product WHERE product_id LIKE '%"+product+"'", productLoad.connection);
                readerProduct = commandProduct.ExecuteReader();
                readerProduct.Read();
                sProduct = readerProduct["product_name"].ToString();
                unit =Convert.ToInt32(readerProduct["product_unit"]);
                price = Convert.ToSingle(readerProduct["product_price"]);
                product_quantity = Convert.ToSingle(readerProduct["product_quantity"]);
                readerProduct.Close();
                //------
                commandUnit = new SqlCommand("SELECT * FROM dbo.Unit WHERE unit_id LIKE '%" + unit+"'", unitLoad.connection);
                readerUnit = commandUnit.ExecuteReader();
                readerUnit.Read();
                sUnit = readerUnit["unit_name"].ToString();
                readerUnit.Close();
                //------
                commandUser = new SqlCommand("SELECT * FROM dbo.Users WHERE user_id LIKE '%" + responsible + "'", userLoad.connection);
                readerUser = commandUser.ExecuteReader();
                readerUser.Read();
                sResponsible = readerUser["user_familija"].ToString() + " " + readerUser["user_imja"].ToString() + " " + readerUser["user_otchestvo"].ToString();
                readerUser.Close();
                this.Text += sProduct + sUnit;

                dataGridView1.Rows.Add(sProduct, sUnit, price);
            }

            reportLoad.connection.Close();
            productLoad.connection.Close();
            unitLoad.connection.Close();
            userLoad.connection.Close();
            //Закрыли коннекты
            
        }
    }
}
