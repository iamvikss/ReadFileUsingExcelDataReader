using ExcelDataReader;
using System;
using System.Data;
using System.IO;
using System.Reflection;

namespace ReadFileUsingExcelDataReader
{
    class Program
    {
        static string currentDirectory = System.IO.Directory.GetCurrentDirectory();

        //Set temp path for reading sample file
        static string excelFilePath = String.Concat(currentDirectory, @"\SampleFile\Country.xls");
        static string csvFilePath = String.Concat(currentDirectory, @"\SampleFile\Country.csv");

        static void Main(string[] args)
        {
            ReadExcelFile();
            ReadCSVFile();
        }

        public static void ReadExcelFile()
        {
            using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
            {

                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Choose one of either 1 or 2:

                    #region 1. Use the reader methods

                    do
                    {
                        while (reader.Read())
                        {
                            //this will return value of column 1 
                            //reader.GetValue(0); 

                            //this will return value of column 2 
                            //reader.GetValue(1); 
                        }
                    } while (reader.NextResult());

                    #endregion

                    #region 2. Use the AsDataSet extension method

                    var result = reader.AsDataSet();

                    // The result of each spreadsheet is in result.Tables
                    var tables = result.Tables;

                    #endregion
                }
            }
        }

        private static void ReadCSVFile()
        {
            using (var stream = File.Open(csvFilePath, FileMode.Open, FileAccess.Read))
            {
                //  - Comma-Separated Values files (CSV format; *.csv)
                using (var reader = ExcelReaderFactory.CreateCsvReader(stream))
                {

                    // Choose one of either 1 or 2:

                    #region 1. Use the reader methods
                    
                    do
                    {
                        while (reader.Read())
                        {
                            //this will return value of column 1 
                            //reader.GetValue(0); 

                            //this will return value of column 2 
                            //reader.GetValue(1); 
                        }
                    } while (reader.NextResult());

                    #endregion

                    #region 2. Use the AsDataSet extension method

                    var result = reader.AsDataSet();

                    // The result of each spreadsheet is in result.Tables
                    var tables = result.Tables;

                    #endregion
                }
            }
        }

    }

}
