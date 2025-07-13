using ExcelUploadApi.Helpers;
using ExcelUploadApi.Repositories;
using Microsoft.Data.SqlClient;

namespace ExcelUploadApi.Services
{
    public class UploadExcelService: IUploadExcelService
    {
        public readonly string _connectionString;

        public UploadExcelService(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DbConn");
        }

        public void ProcessExcel(IFormFile stream, string tableName)
        {
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var transaction = conn.BeginTransaction();

            try
            {
                var repo = new DynamicTableRepository(conn, transaction);

                using var streamReader = stream.OpenReadStream();
                var (headers, rows) = ExcelHelper.ReadExcel(streamReader);
                tableName = string.IsNullOrEmpty(tableName)?"excel_upload":tableName;
                if (headers.Count == 0 || rows.Count == 0)
                {
                    throw new ArgumentException("Excel file is empty or has no valid data.");
                }
                if (headers.Count != rows[0].Count)
                {
                    throw new ArgumentException("Header count does not match row data count.");
                }
                repo.CreateTable(tableName, headers);
                foreach (var row in rows)
                {
                    if (row.Count != headers.Count)
                    {
                        throw new ArgumentException("Row data count does not match header count.");
                    }
                    repo.InsertRow(tableName, headers, row);
                }
                transaction.Commit();
            }
            catch (Exception e)
            {
                transaction.Rollback();
                throw new Exception(e.Message);
            }          
        }
        public List<Dictionary<string, object>> ReadTable(string tableName)
        {
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var transaction = conn.BeginTransaction();
            try
            {
                var repo = new DynamicTableRepository(conn, transaction);
                return repo.ReadTable(tableName);
            }
            catch (Exception e)
            {
                transaction.Rollback();
                throw new Exception(e.Message);
            }
        }
    }
}
