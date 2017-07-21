using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sklad
{
    public class ConnOpen
    {
        static string connStr = @"Data Source=AMAN\SQLEXPRESS;Initial Catalog=sklad;Integrated Security=True";
        public SqlConnection connection = new SqlConnection(connStr);
    }
}
