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
    public partial class Add_user : Form
    {
        public Add_user()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            int role = comboBox1.SelectedIndex+1;
            User A = new User();
            A.add(textBox1.Text, textBox2.Text, textBox3.Text, dateTimePicker1.Value.ToString(), textBox4.Text, textBox5.Text, textBox6.Text, textBox7.Text, "", role, textBox8.Text, textBox9.Text);
        }
    }
}
