﻿namespace ImportExcelToDatabase.Models
{
    public class ResponseModel
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public string FileName { get; set; }
        public object Data { get; set; }    
    }
}
