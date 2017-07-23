using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
        int responsible, product, traffic;
        string sResponsible, sProduct, location, date, waybill;
        float product_quantity;
        private void Report_Load(object sender, EventArgs e)
        {

        }
    }
}
