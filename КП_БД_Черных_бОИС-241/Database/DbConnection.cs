using Npgsql;
using System;
using System.Data;

namespace КП_БД_Черных.Database
{
    public static class DbConnection
    {
       
        private static readonly string connectionString = "Host=localhost;Port=5432;Username=postgres;Password=2288;Database=computer_firm_db";

        public static DataTable ExecuteQuery(string query, NpgsqlParameter[] parameters = null)
        {
            DataTable dataTable = new DataTable();
            try
            {
                using (var connection = new NpgsqlConnection(connectionString))
                {
                    connection.Open();
                    using (var command = new NpgsqlCommand(query, connection))
                    {
                        if (parameters != null)
                            command.Parameters.AddRange(parameters);
                        using (var adapter = new NpgsqlDataAdapter(command))
                        {
                            adapter.Fill(dataTable);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Ошибка БД: " + ex.Message);
            }
            return dataTable;
        }

        // Метод для выполнения команд (INSERT, UPDATE, DELETE) – возвращает количество затронутых строк
        public static int ExecuteNonQuery(string query, NpgsqlParameter[] parameters = null)
        {
            try
            {
                using (var connection = new NpgsqlConnection(connectionString))
                {
                    connection.Open();
                    using (var command = new NpgsqlCommand(query, connection))
                    {
                        if (parameters != null)
                            command.Parameters.AddRange(parameters);
                        return command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Ошибка выполнения запроса: " + ex.Message);
                return -1;
            }
        }
    }
}