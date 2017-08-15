using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sklad
{
    public class User
    {
        public ConnOpen add_user = new ConnOpen();
        public ConnOpen delete_user = new ConnOpen();
        public ConnOpen test_user = new ConnOpen();
        public ConnOpen update_user = new ConnOpen();
        public void add(string familija, string imja, string otchestvo, string DOB, string tel, string mail, string login, string password, string foto, int role, string place_of_work, string position)
        {
            add_user.connection.Open();
            string sql = string.Format("Insert Into Users" +
                       "(user_familija, user_imja, user_otchestvo, DOB, user_tel, user_mail, user_login, user_password, user_foto, role, place_of_work, position) Values(@user_familija, @user_imja, @user_otchestvo, @DOB, @user_tel, @user_mail, @user_login, @user_password, @user_foto, @role, @place_of_work, @position)");
            using (SqlCommand cmd = new SqlCommand(sql, add_user.connection))
            {
                cmd.Parameters.AddWithValue("@user_familija", familija);
                cmd.Parameters.AddWithValue("@user_imja", imja);
                cmd.Parameters.AddWithValue("@user_otchestvo", otchestvo);
                cmd.Parameters.AddWithValue("@DOB", DOB);
                cmd.Parameters.AddWithValue("@user_tel", tel);
                cmd.Parameters.AddWithValue("@user_mail", mail);
                cmd.Parameters.AddWithValue("@user_login", login);
                cmd.Parameters.AddWithValue("@user_password", password);
                cmd.Parameters.AddWithValue("@user_foto", foto);
                cmd.Parameters.AddWithValue("@role", role);
                cmd.Parameters.AddWithValue("@place_of_work", place_of_work);
                cmd.Parameters.AddWithValue("@position", position);
                cmd.ExecuteNonQuery();
            }
            add_user.connection.Close();
        }
        public bool test_id(string id)
        {
            test_user.connection.Open();
            bool test;
            SqlCommand sqlCom = new SqlCommand("SELECT * FROM dbo.Users WHERE user_id LIKE '%" + id + "'", test_user.connection);
            SqlDataReader dr = sqlCom.ExecuteReader();
            if (dr.HasRows == true)
            {
                test = false;
            }
            else
            {
                test = true;
            }
            test_user.connection.Close();
            return test;
        }
        public void delete(string id)
        {
            delete_user.connection.Open();
            string sql = string.Format("Delete from Users where user_id = '{0}'", id);
            using (SqlCommand cmd = new SqlCommand(sql, delete_user.connection))
            {
                cmd.ExecuteNonQuery();
            }
            delete_user.connection.Close();
        }
        public void update(int id, string familija, string imja, string otchestvo, string DOB, string tel, string mail, string login, string password, string foto, int role)
        {
            update_user.connection.Open();
            string sql = string.Format("Update User Set user_familija = " + familija + " Set user_imja = " + imja + " Set user_otchestvo = " + otchestvo + " Set DOB = " + DOB + " Set user_tel = " + tel + " Set user_mail = " + mail + " Set user_login = " + login + " Set user_password = " + password + " Set user_foto = " + foto + "Set role = " + role + " Where user_id = " + id.ToString() + ";");
            using (SqlCommand cmd = new SqlCommand(sql, update_user.connection))
            {
                cmd.ExecuteNonQuery();
            }
            update_user.connection.Close();
        }
    }
}
