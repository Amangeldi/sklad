using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sklad
{
    public class Product
    {
        public ConnOpen add_product = new ConnOpen();
        public ConnOpen delete_product = new ConnOpen();
        public ConnOpen test_product = new ConnOpen();
        public ConnOpen update_product = new ConnOpen();
        public void add(string kod, string name, float price, float quantity, int unit, string description, string foto, string receipt_date, string date_of_last_change, string location)
        {
            add_product.connection.Open();
            string sql = string.Format("Insert Into Product" +
                       "(product_kod, product_name, product_price, product_quantity, product_unit, product_description, product_foto, receipt_date, date_of_last_change, location) Values(@product_kod, @product_name, @product_price, @product_quantity, @product_unit, @product_description, @product_foto, @receipt_date, @date_of_last_change, @location)");
            using (SqlCommand cmd = new SqlCommand(sql, add_product.connection))
            {
                cmd.Parameters.AddWithValue("@product_kod", kod);
                cmd.Parameters.AddWithValue("@product_name", name);
                cmd.Parameters.AddWithValue("@product_price", price);
                cmd.Parameters.AddWithValue("@product_quantity", quantity);
                cmd.Parameters.AddWithValue("@product_unit", unit);
                cmd.Parameters.AddWithValue("@product_description", description);
                cmd.Parameters.AddWithValue("@product_foto", foto);
                cmd.Parameters.AddWithValue("@receipt_date", receipt_date);
                cmd.Parameters.AddWithValue("@date_of_last_change", date_of_last_change);
                cmd.Parameters.AddWithValue("@location", location);
                cmd.ExecuteNonQuery();
            }
            add_product.connection.Close();
        }
        public bool test_id(string id)
        {
            test_product.connection.Open();
            bool test;
            SqlCommand sqlCom = new SqlCommand("SELECT * FROM dbo.Product WHERE product_id LIKE '%" + id + "'", test_product.connection);
            SqlDataReader dr = sqlCom.ExecuteReader();
            if (dr.HasRows == true)
            {
                test = false;
            }
            else
            {
                test = true;
            }
            test_product.connection.Close();
            return test;
        }
        public void delete(string id)
        {
            delete_product.connection.Open();
            string sql = string.Format("Delete from Product where product_id = '{0}'", id);
            using (SqlCommand cmd = new SqlCommand(sql, delete_product.connection))
            {
                cmd.ExecuteNonQuery();
            }
            delete_product.connection.Close();
        }
        public void update(int id, string kod, string name, float price, float quantity, int unit, string description, string foto, string receipt_date, string date_of_last_change, string location)
        {
            update_product.connection.Open();
            string sql = string.Format("Update Product Set product_kod = "+kod+ " Set product_name = " +name+ " Set product_price = "+price+ " Set product_quantity = "+quantity+ " Set product_unit = "+unit+ " Set product_description = "+description+ " Set product_foto = "+foto+ " Set receipt_date = "+receipt_date+ " Set date_of_last_change = "+date_of_last_change+ "Set location = " +location+" Where product_id = " + id.ToString() + ";");
            using (SqlCommand cmd = new SqlCommand(sql, update_product.connection))
            {
                cmd.ExecuteNonQuery();
            }
            update_product.connection.Close();
        }
    }
}
