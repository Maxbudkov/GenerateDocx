using System;
using System.IO;
using MySql.Data.MySqlClient;

namespace phpESH
{
    // Класс подключения к базе данных
    class DBConnection
    {
        private DBConnection()
        {

        }

        private string databaseName = string.Empty;
        private string UID = string.Empty;
        private string password = string.Empty;

        public string DatabaseName
        {
            get { return databaseName; }
            set { databaseName = value; }
        }

        public string Password { get; set; }
        private MySqlConnection connection = null;
        
        public string Password1 { get => password; set => password = value; }
        public string UID1 { get => UID; set => UID = value; }
        public MySqlConnection Connection { get => connection; set => connection = value; }

        private static DBConnection _instance = null;
        public static DBConnection Instance()
        {
            if (_instance == null)
                _instance = new DBConnection();
            return _instance;
        }

        public bool IsConnect()
        {
            if (Connection == null)
            {
                if (String.IsNullOrEmpty(databaseName))
                    return false;
                string connstring = string.Format("Server=localhost; database={0}; UID={1}; password={2}", databaseName, UID1, Password1);
                Connection = new MySqlConnection(connstring);
                try
                {
                    Connection.Open();
                } catch (Exception e)
                {
                    string path = @"log.txt";
                    File.AppendAllText(path, Environment.NewLine + DateTime.Now + "   Error: " + e.Message);
                    Console.WriteLine("Database error. Check log for details. ");
                    return false;
                }
            }

            return true;
        }

        public void Close()
        {
            Connection.Close();
        }
    }
}