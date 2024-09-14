namespace ImportExcelToDatabase.Models
{
    public class ImportDataModel
    {
        public IFormFile ExcelFile { get; set; }
        public string TableName { get; set; }

        public string DatabaseType { get; set; }
        public string ConnectionString { get; set; }
        public string FileName { get; set; }
    }
}
