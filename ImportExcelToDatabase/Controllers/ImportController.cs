using ImportExcelToDatabase.Helpers;
using ImportExcelToDatabase.Models;
using ImportExcelToDatabase.Repository;
using Microsoft.AspNetCore.Mvc;

namespace ImportExcelToDatabase.Controllers
{
    public class ImportController : Controller
    {
        private readonly IWebHostEnvironment _environment;
        private readonly List<string> _databaseTypes;   
        public ImportController(IWebHostEnvironment environment)
        {
            _environment = environment;
            _databaseTypes = new List<string>()
            {
                "MYSQL",
                "MSSQL",
                "POSTGRESQL",
                "ORACLE"
            };
        }

        [HttpPost]
        public async Task<IActionResult> UploadFile([FromForm] ImportDataModel model)
        {
            #region Check Condition
            
            if (string.IsNullOrEmpty(model.TableName))
            {
                return new OkObjectResult(new ResponseModel()
                {
                    Success = false,
                    Message = $"Table Name is required."
                });
            }

            if (string.IsNullOrEmpty(model.ConnectionString))
            {
                return new OkObjectResult(new ResponseModel()
                {
                    Success = false,
                    Message = $"Connection String is required."
                });
            }

            if (model.ExcelFile == null || model.ExcelFile.Length == 0)
            {
                return new OkObjectResult(new ResponseModel()
                {
                    Success = false,
                    Message = $"Excel File is required."
                });
            }

            if (string.IsNullOrEmpty(model.DatabaseType))
            {
                return new OkObjectResult(new ResponseModel()
                {
                    Success = false,
                    Message = $"Database Type is required."
                });
            }

            if (!_databaseTypes.Contains(model.DatabaseType.ToUpper()))
            {
                return new OkObjectResult(new ResponseModel()
                {
                    Success = false,
                    Message = $"Database Type \"{model.DatabaseType}\" do not supported."
                });
            }

            model.DatabaseType = model.DatabaseType.ToUpper();
            #endregion

            var uploadsFolder = Path.Combine(_environment.WebRootPath, "tempFile");
            if (!Directory.Exists(uploadsFolder))
            {
                Directory.CreateDirectory(uploadsFolder);
            }

            var fileName = Guid.NewGuid().ToString() + "_" + DateTime.Now.Ticks + Path.GetExtension(model.ExcelFile.FileName);
            var filePath = Path.Combine(uploadsFolder, fileName);

            using (var fileStream = new FileStream(filePath, FileMode.Create))
            {
                await model.ExcelFile.CopyToAsync(fileStream);
            }

            var data = ExcelHelper.ReadExcel(filePath);

            if (data.HasError)
            {
                System.IO.File.Delete(filePath);
                return new OkObjectResult(new ResponseModel()
                {
                    Success = false,
                    Message = data.ErrorMessage,
                    FileName = fileName
                });
            }

            var insertDataToDatabase = new InsertDataToDatabase();
            var result = await insertDataToDatabase.InsertAysnc(data.DataList, model);

            System.IO.File.Delete(filePath);

            return new OkObjectResult(result);
        }
    }
}
