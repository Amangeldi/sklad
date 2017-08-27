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
    public partial class Users : Form
    {
        public Users()
        {
            InitializeComponent();
        }

        private void Users_Load(object sender, EventArgs e)
        {
            ConnOpen usersLoad = new ConnOpen();
            SqlDataAdapter adapter = new SqlDataAdapter();
            usersLoad.connection.Open();
            SqlCommand sqlCom = new SqlCommand("SELECT user_id, fio, DOB, user_tel, user_mail, role, user_login, place_of_work, position FROM dbo.users", usersLoad.connection);
            usersLoad.connection.Close();
            adapter.SelectCommand = sqlCom;
            DataSet dataset = new DataSet();
            adapter.Fill(dataset);
            dataGridView1.DataSource = dataset.Tables[0];
            adapter.Update(dataset);
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Add_user f = new Add_user();
            f.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            User u = new User();
            bool test = u.test_id(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            if (test == true)
            {
                MessageBox.Show("Материал с таким id не существует или был удален", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            else
            {
                Edit_user f = new Edit_user(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value));
                f.ShowDialog();
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            User u = new User();
            bool test = u.test_id(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            if (test == true)
            {
                MessageBox.Show("Пользаватель с таким id не существует или был удален", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            else
            {
                u.delete(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                MessageBox.Show("Удален пользователь " + dataGridView1.CurrentRow.Cells[0].Value.ToString(), "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
