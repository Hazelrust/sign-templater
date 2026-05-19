using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using ExcelDataReader;

namespace sign_templater.Services
{
    public static class DataService
    {
        public static DataTable LoadData(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    });
                    return result.Tables[0];
                }
            }
        }
    }
}
