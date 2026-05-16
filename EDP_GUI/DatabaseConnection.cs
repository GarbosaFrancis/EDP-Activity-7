using MySql.Data.MySqlClient;

namespace EDP_GUI
{
    public class DatabaseConnection
    {
        // If your XAMPP/MySQL root user has a password, put it between the pwd=''; quotes!
        private string connectionString = "server=localhost;database=hobbyshop_db;uid=root;pwd=;";

        public MySqlConnection GetConnection()
        {
            return new MySqlConnection(connectionString);
        }
    }
}