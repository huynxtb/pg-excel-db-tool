using ImportExcelToDatabase.Models;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace ImportExcelToDatabase.Helpers
{
    public static class ExcelHelper
    {
        public static ExcelResponseModel ReadExcel(string filePath)
        {
            var listErrors = new List<string>();
            var dataList = new List<Dictionary<string, object>>();
            var fileInfo = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;
                var colCount = worksheet.Dimension.Columns;

                // Read header values
                var headers = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    string header = worksheet.Cells[1, col].Value.ToString();

                    if (string.IsNullOrEmpty(header) || string.IsNullOrWhiteSpace(header)) {
                        listErrors.Add($"At index row 1 and col {col} value can not null or have any white space");
                    }
                    else
                    {
                        if (header.Any(char.IsWhiteSpace))
                        {
                            listErrors.Add($"At index row 1 and col {col} value can not null or have any white space");
                        }
                        else
                        {
                            headers.Add(header);
                        }
                    }
                }

                if (listErrors.Count > 0)
                {
                    return new ExcelResponseModel()
                    {
                        HasError = true,
                        ErrorMessage = string.Join("; ", listErrors)
                    };
                }

                // Start from the second row to skip the header
                for (var row = 2; row <= rowCount; row++)
                {
                    var rowData = new Dictionary<string, object>();

                    for (int col = 1; col <= colCount; col++)
                    {
                        var header = headers[col - 1];
                        object cellValue = worksheet.Cells[row, col].Value;

                        // Check if the value is a number, and if so, try to convert it to the appropriate type
                        if (double.TryParse(cellValue?.ToString(), out double numberValue))
                        {
                            if (numberValue % 1 == 0)
                            {
                                rowData[header] = Convert.ToInt32(numberValue);
                            }
                            else
                            {
                                rowData[header] = numberValue;
                            }
                        }
                        else
                        {
                            rowData[header] = cellValue?.ToString();
                        }
                    }

                    dataList.Add(rowData);
                }
            }

            string jsonString = JsonConvert.SerializeObject(dataList);

            return new ExcelResponseModel()
            {
                JsonString = jsonString,
                HasError = false,
                DataList = dataList
            };
        }
    }
}
