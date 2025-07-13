
using System;
using System.Data;
using Microsoft.Data.SqlClient;

namespace ExcelUploadApi.Repositories
{
    public class DynamicTableRepository
    {
        private readonly SqlConnection _connection;
        private readonly SqlTransaction _transaction;

        public DynamicTableRepository(SqlConnection connection, SqlTransaction transaction)
        {
            _connection = connection;
            _transaction = transaction;
        }

        public void CreateTable(string tableName, List<string> columns)
        {
            var columnList = string.Join(",", columns);


            using var command = new SqlCommand("uspUploadExcel", _connection, _transaction);
            command.CommandType = CommandType.StoredProcedure;

            command.Parameters.AddWithValue("@action", "create_table");
            command.Parameters.AddWithValue("@tableName", tableName);
            command.Parameters.AddWithValue("@columns", columnList);

            command.ExecuteNonQuery();
        }

        public void InsertRow(string tableName, List<string> columns, List<string> values)
        {
            var columnList = string.Join(",", columns);
            var valueList = string.Join(";", values);

            using var command = new SqlCommand("uspUploadExcel", _connection, _transaction);
            command.CommandType = CommandType.StoredProcedure;

            command.Parameters.AddWithValue("@action", "insert_row");
            command.Parameters.AddWithValue("@tableName", tableName);
            command.Parameters.AddWithValue("@columns", columnList);
            command.Parameters.AddWithValue("@values", valueList);

            command.ExecuteNonQuery();
        }

        public List<Dictionary<string, object>> ReadTable(string tableName)
        {
            var result = new List<Dictionary<string, object>>();
            using var command = new SqlCommand("uspUploadExcel", _connection, _transaction);
            command.CommandType = CommandType.StoredProcedure;
            command.Parameters.AddWithValue("@action", "read_table");
            command.Parameters.AddWithValue("@tableName", tableName);
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                var row = new Dictionary<string, object>();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    row[reader.GetName(i)] = reader.GetValue(i);
                }
                result.Add(row);
            }
            return result;
        }
    }

}
