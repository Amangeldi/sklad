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
    public partial class Otchet : Form
    {
        public string sProduct = " ", report = "Материал ", sTraffic =" " ;
        public int product;
        public Otchet(string _product)
        {
            InitializeComponent();
            this.sProduct = _product;
            this.product =Convert.ToInt32(_product);
        }

        private void Otchet_Load(object sender, EventArgs e)
        {
            int responsible;
            string sResponsible, location, date, dateChas, dateMin, waybill, pName;
            float product_quantity;
            ConnOpen proLoad = new ConnOpen();
            ConnOpen resLoad = new ConnOpen();
            ConnOpen userLoad = new ConnOpen();
            proLoad.connection.Open();
            resLoad.connection.Open();
            userLoad.connection.Open();
            SqlCommand commandPro = new SqlCommand("SELECT * FROM dbo.Product WHERE product_id = '"+sProduct+"'",proLoad.connection );
            SqlDataReader readerPro = commandPro.ExecuteReader();
            readerPro.Read();
            pName = readerPro["product_name"].ToString();
            product_quantity = Convert.ToInt32(readerPro["product_quantity"]);
            readerPro.Close();
            SqlCommand commandRes = new SqlCommand("SELECT * FROM dbo.Responsibility WHERE product = '"+product+"'", resLoad.connection);
            SqlDataReader readerRes = commandRes.ExecuteReader();
            SqlCommand commandUser;
            SqlDataReader readerUser;
            this.Text = "Подробный отчет по материалу " + sProduct;
            label1.Text = "";
            while (readerRes.Read())
            {
                responsible = Convert.ToInt32(readerRes["responsible"]);
                commandUser = new SqlCommand("SELECT * FROM dbo.Users WHERE user_id = '" + responsible + "'",userLoad.connection);
                readerUser = commandUser.ExecuteReader();
                readerUser.Read();
                sResponsible = readerUser["fio"].ToString();
                location = readerRes["location"].ToString();
                date =String.Format("{0:dd.MM.yyyy}",readerRes["date"]);
                dateChas = String.Format("{0:HH}", readerRes["date"]);
                dateMin = String.Format("{0:mm}", readerRes["date"]);
                waybill = readerRes["waybill"].ToString();
                if(readerRes["traffic"].ToString()=="0")
                {
                    sTraffic = "был получен с ";
                }
                else if(readerRes["traffic"].ToString()=="1")
                {
                    sTraffic = "был отправлен к ";
                }
                readerUser.Close();
                report = "Материал " + pName+" "+sTraffic+sResponsible+" со склада "+location+" "+ date+" "+dateChas+" часов "+dateMin+" минут на основании накладной номер " +waybill;
                label1.Text += report+"\n";
            }
            proLoad.connection.Close();
            resLoad.connection.Close();
            userLoad.connection.Close();
        }
    }
}
