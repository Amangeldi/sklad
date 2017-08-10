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
    public partial class Waybill : Form
    {
        int responsible, product, unit;
        string waybill, user, traffic, product_name, unit_name, date;
        float product_quantity, price, summa;
        public Waybill(string _waybill, string _traffic)
        {
            InitializeComponent();
            this.waybill = _waybill;
            this.traffic = _traffic;
        }

        private void Waybill_Load(object sender, EventArgs e)
        {
            var columnNo = new DataGridViewColumn();
            columnNo.HeaderText = "No";
            columnNo.Name = "nomer";
            columnNo.CellTemplate = new DataGridViewTextBoxCell();

            var columnPName = new DataGridViewColumn();
            columnPName.HeaderText = "Maddy gymmatlygyn ady";
            columnPName.Name = "productName";
            columnPName.CellTemplate = new DataGridViewTextBoxCell();

            var columnUnit = new DataGridViewColumn();
            columnUnit.HeaderText = "Maddy gymmatlygyn ady";
            columnUnit.Name = "productName";
            columnUnit.CellTemplate = new DataGridViewTextBoxCell();

            var columnMT = new DataGridViewColumn();
            columnMT.HeaderText = "Talap edileni";
            columnMT.Name = "talap";
            columnMT.CellTemplate = new DataGridViewTextBoxCell();

            var columnMG = new DataGridViewColumn();
            columnMG.HeaderText = "goyberileni";
            columnMG.Name = "goyber";
            columnMG.CellTemplate = new DataGridViewTextBoxCell();

            var columnPrice = new DataGridViewColumn();
            columnPrice.HeaderText = "Bahasy";
            columnPrice.Name = "price";
            columnPrice.CellTemplate = new DataGridViewTextBoxCell();

            var columnAdditional = new DataGridViewColumn();
            columnAdditional.HeaderText = "Gosmaca gymmaty";
            columnAdditional.Name = "additional";
            columnAdditional.CellTemplate = new DataGridViewTextBoxCell();

            dataGridView1.Columns.Add(columnNo);
            dataGridView1.Columns.Add(columnPName);
            dataGridView1.Columns.Add(columnUnit);
            dataGridView1.Columns.Add(columnMT);
            dataGridView1.Columns.Add(columnMG);
            dataGridView1.Columns.Add(columnPrice);
            dataGridView1.Columns.Add(columnAdditional);
            //-------
            this.Text = "Talapnama - yan haty No "+waybill;
            ConnOpen respLoad = new ConnOpen();
            ConnOpen productLoad = new ConnOpen();
            ConnOpen unitLoad = new ConnOpen();
            ConnOpen userLoad = new ConnOpen();
            ConnOpen respForDGV = new ConnOpen();
            respLoad.connection.Open();
            SqlCommand commandResp = new SqlCommand("SELECT * FROM dbo.Responsibility WHERE waybill = '" + waybill+"' AND traffic = '"+traffic+"'", respLoad.connection);
            SqlDataReader readerResp = commandResp.ExecuteReader();
            readerResp.Read();
            responsible = Convert.ToInt32(readerResp["responsible"]);
            readerResp.Close();
            respLoad.connection.Close();
            //-------
            userLoad.connection.Open();
            SqlCommand commandUser = new SqlCommand("SELECT * FROM dbo.Users WHERE user_id = '"+responsible+"'",userLoad.connection);
            SqlDataReader readerUser = commandUser.ExecuteReader();
            readerUser.Read();        
            user = readerUser["fio"].ToString();
            readerUser.Close();
            userLoad.connection.Close();
            //-------
            respForDGV.connection.Open();
            productLoad.connection.Open();
            unitLoad.connection.Open();
            SqlCommand commandForDGV = new SqlCommand("SELECT * FROM dbo.Responsibility WHERE waybill = '" + waybill + "' AND traffic = '" + traffic + "'", respForDGV.connection);
            SqlDataReader readerForDGV = commandForDGV.ExecuteReader();
            SqlCommand commandProduct;
            SqlDataReader readerProduct;
            SqlCommand commandUnit;
            SqlDataReader readerUnit;
            int n=0;
            while(readerForDGV.Read())
            {
                n++;
                product = Convert.ToInt32(readerForDGV["product"]);
                commandProduct = new SqlCommand("SELECT * FROM dbo.Product WHERE product_id = '"+product+"'", productLoad.connection);
                readerProduct = commandProduct.ExecuteReader();
                readerProduct.Read();
                unit = Convert.ToInt32(readerProduct["product_unit"]);
                product_name = readerProduct["product_name"].ToString();
                price = Convert.ToSingle(readerProduct["product_price"]);
                readerProduct.Close();
                commandUnit = new SqlCommand("SELECT * FROM dbo.Unit WHERE unit_id = '" + unit + "'", unitLoad.connection);
                readerUnit = commandUnit.ExecuteReader();
                readerUnit.Read();
                unit_name = readerUnit["unit_name"].ToString();
                readerUnit.Close();
                product_quantity =Convert.ToSingle( readerForDGV["product_quantity"]);
                summa = product_quantity * price;
                dataGridView1.Rows.Add(n, product_name, unit_name, " ", product_quantity, price, summa);

            }
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            respForDGV.connection.Close();
            productLoad.connection.Close();
            unitLoad.connection.Close();
            //-------
            label1.Text = "Kimin usti bilen \t";
            label2.Text = "Talap eden \t";
            if(traffic=="0")
            {
                label2.Text += "Аннаклычев Хакнепес Амангелдиевич";
                label1.Text += user;
            }
            else if(traffic=="1")
            {
                label1.Text += "Аннаклычев Хакнепес Амангелдиевич";
                label2.Text += user;
            }
            
        }
    }
}
