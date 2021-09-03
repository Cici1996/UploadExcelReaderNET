using dotNetUseCase.Constants;
using dotNetUseCase.DataAccess.Interfaces;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace dotNetUseCase.DataAccess
{
    public class ManageFileRepository : IManageFileRepository
    {
        private readonly IConfiguration _configuration;
        private readonly string _connectionString;

        public ManageFileRepository(IConfiguration configuration)
        {
            _configuration = configuration;
            _connectionString = _configuration.GetValue<string>(GeneralConstant.ConnectionString);
        }

        public void CreateTable(string tableName, List<string> columns)
        {
            try
            {
                StringBuilder queryTable = new StringBuilder();
                int index = 0;
                int totalColumn = columns.Count;
                queryTable.Append($"CREATE TABLE {tableName} (");
                foreach (string item in columns)
                {
                    string comma = ((index + 1).Equals(totalColumn)) ? "" : ",";
                    queryTable.Append($"{item} varchar(255) {comma}");
                    index++;
                }
                queryTable.Append($")");
                DropTable(tableName);
                using SqlConnection connection = new SqlConnection(_connectionString);
                connection.Open();
                using (SqlCommand command = new SqlCommand(queryTable.ToString(), connection))
                {
                    command.ExecuteNonQuery();
                }
                connection.Close();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void DropTable(string tableName)
        {
            StringBuilder queryTable = new StringBuilder();
            queryTable.Append($"DROP TABLE IF EXISTS {tableName}");
            using SqlConnection connection = new SqlConnection(_connectionString);
            connection.Open();
            using (SqlCommand command = new SqlCommand(queryTable.ToString(), connection))
            {
                command.ExecuteNonQuery();
            }
            connection.Close();
        }

        public void InsertBulkData(string tableName, DataTable data)
        {
            using SqlConnection connection = new SqlConnection(_connectionString);
            connection.Open();
            using SqlBulkCopy bulkCopy = new SqlBulkCopy(connection);
            foreach (DataColumn c in data.Columns)
                bulkCopy.ColumnMappings.Add(c.ColumnName.ToLower(), c.ColumnName.ToLower());

            bulkCopy.DestinationTableName = tableName;
            try
            {
                bulkCopy.WriteToServer(data);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            connection.Close();
        }

        public DataTable SelectData(string tableName)
        {
            using var conn = new SqlConnection(_connectionString);
            string command = $"SELECT * FROM {tableName}";
            DataTable table = new DataTable();
            using var cmd = new SqlCommand(command, conn);
            SqlDataAdapter adapt = new SqlDataAdapter(cmd);
            conn.Open();
            adapt.Fill(table);
            conn.Close();
            return table;
        }
    }
}