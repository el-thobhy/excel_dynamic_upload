using ExcelUploadApi.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ExcelUploadApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UploadController : ControllerBase
    {
        private readonly IUploadExcelService _uploadExcelService;

        public UploadController(IUploadExcelService uploadExcelService)
        {
            _uploadExcelService = uploadExcelService;
        }
        [HttpPost("upload")]
        public IActionResult UploadExcel(IFormFile file, string tableName)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }
            try
            {
                _uploadExcelService.ProcessExcel(file, tableName);
                return Ok("File uploaded and processed successfully.");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }
        [HttpGet("read")]
        public IActionResult ReadTable(string tableName)
        {
            if (string.IsNullOrEmpty(tableName))
            {
                return BadRequest("Table name is required.");
            }
            try
            {
                var result = _uploadExcelService.ReadTable(tableName);
                return Ok(result);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }
    }
}
