using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelUploadApi.Helpers
{
    public static class ExcelHelper
    {
        public static (List<string> Headers, List<List<string>> Rows) ReadExcel(Stream stream)
        {
            var headers = new List<string>();
            var rows = new List<List<string>>();

            IWorkbook workbook = new XSSFWorkbook(stream);
            ISheet sheet = workbook.GetSheetAt(0);
            
            var headerRow = sheet.GetRow(0);
            for(int i = 0; i < headerRow.LastCellNum; i++)
            {
                var cell = headerRow.GetCell(i);
                headers.Add(cell?.ToString()?.Trim() ?? $"Column{i}");
            }

            for(int i = 1; i<= sheet.LastRowNum; i++) //start from row 1, karena row 0 adalaha header
            {
                var row = sheet.GetRow(i);
                if (row == null) continue;
                var rowData = new List<string>();
                for (int j = 0; j < row.LastCellNum; j++)
                {
                    var cell = row.GetCell(j);
                    rowData.Add(cell?.ToString()?.Trim() ?? string.Empty);
                }
                rows.Add(rowData);
            }

            return (headers, rows);
        }
    }
}
