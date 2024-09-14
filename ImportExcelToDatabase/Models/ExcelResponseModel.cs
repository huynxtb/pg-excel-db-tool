namespace ImportExcelToDatabase.Models
{
    public class ExcelResponseModel
    {
        public bool HasError { get; set; }
        public string ErrorMessage { get; set; }
        public string JsonString { get; set; }
        public List<Dictionary<string, object>> DataList { get; set; }
    }
}
