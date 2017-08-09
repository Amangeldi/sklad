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
        float product_quantity, price, rProduct_quantity, pValue = 0, rValue = 0, sum, sPrice;
        private void Report_Load(object sender, EventArgs e)
        {
            ConnOpen reportLoad = new ConnOpen();
            ConnOpen productLoad = new ConnOpen();
            ConnOpen unitLoad = new ConnOpen();
            ConnOpen userLoad = new ConnOpen();
            ConnOpen tLoad = new ConnOpen();
            reportLoad.connection.Open();
            productLoad.connection.Open();
            unitLoad.connection.Open();
            userLoad.connection.Open();
            tLoad.connection.Open();
            //Открыли все коннекты
            SqlCommand commandResponsible = new SqlCommand("SELECT * FROM dbo.Responsibility", reportLoad.connection);
            SqlDataReader readerResponsible = commandResponsible.ExecuteReader();
            SqlCommand commandProduct = new SqlCommand("SELECT * FROM dbo.Product WHERE product_flag = '" + 1 + "'", productLoad.connection);
            SqlDataReader readerProduct = commandProduct.ExecuteReader();
            SqlCommand commandUnit;
            SqlDataReader readerUnit;
            SqlCommand commandUser;
            SqlDataReader readerUser;
            SqlCommand commandT;
            SqlDataReader readerT;
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

            var prih = new DataGridViewColumn();
            prih.HeaderText = "Приход";
            prih.Name = "prihod";
            prih.CellTemplate = new DataGridViewTextBoxCell();

            var rash = new DataGridViewColumn();
            rash.HeaderText = "Расход";
            rash.Name = "rashod";
            rash.CellTemplate = new DataGridViewTextBoxCell();

            var ostatok = new DataGridViewColumn();
            ostatok.HeaderText = "Остаток на ";
            ostatok.Name = "ostatok";
            ostatok.CellTemplate = new DataGridViewTextBoxCell();

            var summa = new DataGridViewColumn();
            summa.HeaderText = "Остаток";
            summa.Name = "summa";
            summa.CellTemplate = new DataGridViewTextBoxCell();

            var priceSumma = new DataGridViewColumn();
            priceSumma.HeaderText = "Итого цена";
            priceSumma.Name = "sumPrice";
            priceSumma.CellTemplate = new DataGridViewTextBoxCell();

            dataGridView1.Columns.Add(columnPName);
            dataGridView1.Columns.Add(columnPUnit);
            dataGridView1.Columns.Add(ostatok);
            dataGridView1.Columns.Add(columnPPrice);
            dataGridView1.Columns.Add(prih);
            dataGridView1.Columns.Add(rash);
            dataGridView1.Columns.Add(summa);
            dataGridView1.Columns.Add(priceSumma);
            //Добавили постоянные колонки
            while (readerProduct.Read())
            {
                product = Convert.ToInt32(readerProduct["product_id"]);
                sProduct = readerProduct["product_name"].ToString();
                unit = Convert.ToInt32(readerProduct["product_unit"]);
                price = Convert.ToSingle(readerProduct["product_price"]);
                product_quantity = Convert.ToSingle(readerProduct["product_quantity"]);
                //-------
                commandUnit = new SqlCommand("SELECT * FROM dbo.Unit WHERE unit_id = '" + unit + "'", unitLoad.connection);
                readerUnit = commandUnit.ExecuteReader();
                readerUnit.Read();
                sUnit = readerUnit["unit_name"].ToString();
                readerUnit.Close();
                //-------
                commandT = new SqlCommand("SELECT * FROM dbo.Responsibility WHERE product = '"+product+"'", tLoad.connection);
                readerT = commandT.ExecuteReader();
                while(readerT.Read())
                {
                    if(readerT["traffic"].ToString() == "0")
                    {
                        pValue += Convert.ToInt32(readerT["product_quantity"]);
                    }
                    if (readerT["traffic"].ToString() == "1")
                    {
                        rValue += Convert.ToInt32(readerT["product_quantity"]);
                    }
                }
                readerT.Close();
                sum = product_quantity + pValue - rValue;
                sPrice = sum * price;
                dataGridView1.Rows.Add(sProduct, sUnit, product_quantity, price, pValue, rValue, sum, sPrice);
                pValue = 0;
                rValue = 0;
            }
            //while (readerResponsible.Read())
            //{
            //    product = Convert.ToInt32(readerResponsible["product"]);
            //    responsible = Convert.ToInt32(readerResponsible["responsible"]);
            //    traffic = Convert.ToInt32(readerResponsible["traffic"]);
            //    waybill = readerResponsible["waybill"].ToString();
            //    rProduct_quantity = Convert.ToSingle(readerResponsible["product_quantity"]);
            //    //-----
            //    commandProduct = new SqlCommand("SELECT * FROM dbo.Product WHERE product_id LIKE '%"+product+"'", productLoad.connection);
            //    readerProduct = commandProduct.ExecuteReader();
            //    readerProduct.Read();
            //    sProduct = readerProduct["product_name"].ToString();
            //    unit =Convert.ToInt32(readerProduct["product_unit"]);
            //    price = Convert.ToSingle(readerProduct["product_price"]);
            //    product_quantity = Convert.ToSingle(readerProduct["product_quantity"]);
            //    readerProduct.Close();
            //    //------
            //    commandUnit = new SqlCommand("SELECT * FROM dbo.Unit WHERE unit_id LIKE '%" + unit+"'", unitLoad.connection);
            //    readerUnit = commandUnit.ExecuteReader();
            //    readerUnit.Read();
            //    sUnit = readerUnit["unit_name"].ToString();
            //    readerUnit.Close();
            //    //------
            //    commandUser = new SqlCommand("SELECT * FROM dbo.Users WHERE user_id LIKE '%" + responsible + "'", userLoad.connection);
            //    readerUser = commandUser.ExecuteReader();
            //    readerUser.Read();
            //    sResponsible = readerUser["user_familija"].ToString() + " " + readerUser["user_imja"].ToString() + " " + readerUser["user_otchestvo"].ToString();
            //    readerUser.Close();
            //    this.Text += sProduct + sUnit;

            //    dataGridView1.Rows.Add(sProduct, sUnit, price);
            //}

            reportLoad.connection.Close();
            productLoad.connection.Close();
            unitLoad.connection.Close();
            userLoad.connection.Close();
            tLoad.connection.Close();
            //Закрыли коннекты

        }
    }
}
