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
    public partial class Waybill : Form
    {
        string waybill;
        public Waybill(string _waybill)
        {
            InitializeComponent();
            this.waybill = _waybill;
        }

        private void Waybill_Load(object sender, EventArgs e)
        {
            this.Text = waybill;
        }
    }
}
