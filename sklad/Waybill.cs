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
        float product_quantity, price;
        public Waybill(string _waybill, string _traffic)
        {
            InitializeComponent();
            this.waybill = _waybill;
            this.traffic = _traffic;
        }

        private void Waybill_Load(object sender, EventArgs e)
        {
            this.Text = "Talapnama - yan haty No "+waybill;
            ConnOpen respLoad = new ConnOpen();
            ConnOpen productLoad = new ConnOpen();
            ConnOpen unitLoad = new ConnOpen();
            ConnOpen userLoad = new ConnOpen();
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
