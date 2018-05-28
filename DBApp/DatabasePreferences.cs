using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DBApp
{
    class DatabasePreferences
    {
        static public SqlConnection sqlConnection;
        static public SqlConnection getDb()
        {
            string dbPathMyDocs = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string dbPath = Path.Combine(dbPathMyDocs,"Database.mdf");
            //string dbPath = Path.Combine(dbPathMyDocs, "DBApp", "DBApp", "Database.mdf");
            //string connectionString = @"Data Source=.\SQLEXPRESS; Initial Catalog=Database;User Id=User;Password=Pass";
            string connectionString = @"Server=(localdb)\MSSQLLocalDB; Integrated Security=True; AttachDbFileName=" + dbPath + "";

            sqlConnection = new SqlConnection(connectionString);
            sqlConnection.Open();
            //return true when database is connected 
            return sqlConnection;
        }

    }
}
