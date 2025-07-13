namespace ExcelUploadApi.Services
{
    public interface IUploadExcelService
    {
        void ProcessExcel(IFormFile stream, string tableName);
        List<Dictionary<string, object>> ReadTable(string tableName);
    }
}
