using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;

namespace ReadExcel
{
    public class ExcelRepository : IReadFile
    {
        public ExcelRepository()
        {
        }

        private const string path = "/Users/houjingye/Documents/c#/ReadExcel/ReadExcel/Book.xlsx";

        public List<string> GetAllItems()
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {

                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {

                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    do
                    {
                        while (reader.Read())
                        {
                            Console.WriteLine(reader.GetValue(0));// reader.GetDouble(0);
                        }
                    } while (reader.NextResult());

                    // 2. Use the AsDataSet extension method
                    var result = reader.AsDataSet();
                    //Console.WriteLine(reader.RowCount);
                    // The result of each spreadsheet is in result.Tables
                }
            }
            return null;
        }
    }
}
