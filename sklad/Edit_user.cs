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
    public partial class Edit_user : Form
    {
        int uId;
        public Edit_user(int _id)
        {
            InitializeComponent();
            this.uId = _id;
        }

        private void Edit_user_Load(object sender, EventArgs e)
        {
            ConnOpen userLoad = new ConnOpen();
            userLoad.connection.Open();
            SqlCommand command = new SqlCommand("SELECT * FROM dbo.Users WHERE user_id LIKE '" + uId + "'", userLoad.connection);
            SqlDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                textBox1.Text = reader["user_familija"].ToString();
                textBox2.Text = reader["user_imja"].ToString();
                textBox3.Text = reader["user_otchestvo"].ToString();
                dateTimePicker1.Value = Convert.ToDateTime(reader["DOB"]);
                textBox4.Text = reader["user_tel"].ToString();
                textBox5.Text = reader["user_mail"].ToString();
                textBox6.Text = reader["user_login"].ToString();
                textBox7.Text = reader["user_password"].ToString();
                textBox8.Text = reader["place_of_work"].ToString();
                textBox9.Text = reader["position"].ToString();
                comboBox1.SelectedIndex = Convert.ToInt32(reader["role"]) - 1;
            }
            userLoad.connection.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            User u = new User();
            u.update(uId, textBox1.Text, textBox2.Text, textBox3.Text, dateTimePicker1.Value.ToShortDateString(), textBox4.Text, textBox5.Text, textBox6.Text, textBox7.Text, "", comboBox1.SelectedIndex + 1, textBox8.Text, textBox9.Text);
            MessageBox.Show("Пользователь изменен", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }
    }
}
