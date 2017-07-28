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
        int responsible, product, traffic, price, unit;
        string sResponsible, sProduct, sUnit, location, date, waybill;
        float product_quantity;
        private void Report_Load(object sender, EventArgs e)
        {
            ConnOpen reportLoad = new ConnOpen();
            ConnOpen productLoad = new ConnOpen();
            ConnOpen unitLoad = new ConnOpen();
            reportLoad.connection.Open();
            productLoad.connection.Open();
            unitLoad.connection.Open();
            SqlCommand commandResponsible = new SqlCommand("SELECT * FROM dbo.Responsibility", reportLoad.connection);
            SqlDataReader readerResponsible = commandResponsible.ExecuteReader();
            SqlCommand commandProduct;
            SqlDataReader readerProduct;
            SqlCommand commandUnit;
            SqlDataReader readerUnit;

            var columnPName = new DataGridViewColumn();
            columnPName.HeaderText = "";
            columnPName.Name = "productName";
            while(readerResponsible.Read())
            {
                product = Convert.ToInt32(readerResponsible["product"]);
                //-----
                commandProduct = new SqlCommand("SELECT * FROM dbo.Product WHERE product_id LIKE '%"+product+"'", productLoad.connection);
                readerProduct = commandProduct.ExecuteReader();
                readerProduct.Read();
                sProduct = readerProduct["product_name"].ToString();
                unit = Convert.ToInt32(readerProduct["product_unit"]);
                //------
                commandUnit = new SqlCommand("", unitLoad.connection);
                readerUnit = commandUnit.ExecuteReader();
                sUnit = readerUnit["unit_name"].ToString();

            }
            reportLoad.connection.Close();
            productLoad.connection.Close();
            unitLoad.connection.Close();
        }
    }
}
