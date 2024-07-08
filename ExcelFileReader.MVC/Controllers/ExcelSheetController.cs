using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;

namespace ExcelFileReader.MVC.Controllers
{
    public class ExcelSheetController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> UploadExcelSheet(IFormFile requestFile)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            if (requestFile is null || requestFile.Length < 1)
            {
                throw new Exception("No Excel File");
            }

            var uploadDirectory = $"{Directory.GetCurrentDirectory()}\\wwwroot\\Uploads";

            if (!Directory.Exists(uploadDirectory))
            {
                Directory.CreateDirectory(uploadDirectory);
            }

            // Upload File Start
            var filePath = Path.Combine(uploadDirectory, requestFile.FileName);

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await requestFile.CopyToAsync(stream);
            }

            // Upload File End

            using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                var excelData = new List<List<object>>();
                using var reader = ExcelReaderFactory.CreateReader(stream);
                do
                {
                    while (reader.Read())
                    {
                        var rowData = new List<object>();
                        for (int column = 0; column < reader.FieldCount; column++)
                        {
                            rowData.Add(reader.GetValue(column));
                        }
                        excelData.Add(rowData);
                    }
                } while (reader.NextResult());

                return Ok(excelData);
            }

            return View();
        }
    }
}
