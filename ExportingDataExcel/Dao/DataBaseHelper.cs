using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportingDataExcel.Dao
{
    public class DataBaseHelper
    {
        private static string ConnectionString = ConfigurationManager
            .ConnectionStrings["ExportDataToExcel"].ConnectionString;

        public static SqlConnection GetConnection()
        {
            SqlConnection conn = new SqlConnection(ConnectionString);

            if (conn.State == ConnectionState.Closed)
                conn.Open();

            return conn;
        }

        public static void CloseConnection()
        {
            SqlConnection connection = GetConnection();

            if (connection.State == ConnectionState.Open)
                connection.Close();
        }
    }
}
