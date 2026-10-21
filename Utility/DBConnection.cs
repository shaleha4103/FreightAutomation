using System.Configuration;
using System.Data.SqlClient;

namespace FreightAutomation.Utility
{
    public class DBConnection
    {
        private readonly string connectionString;

        public DBConnection()
        {
            connectionString =
                ConfigurationManager
                .ConnectionStrings["CreateCon"]
                .ConnectionString;
        }

        public SqlConnection GetConnection()
        {
            return new SqlConnection(connectionString);
        }
    }
}