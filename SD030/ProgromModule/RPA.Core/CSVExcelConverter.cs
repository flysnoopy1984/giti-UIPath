using CsvHelper;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;

namespace RPA.Core
{
    public  class CSVExcelConverter
    {
        private string _FilePath;

        public delegate void ColumnSetting(ref Dictionary<int, int> colsType);
        public ColumnSetting columnSettingDele;

        public Action<DataTable> AdjustData;

        private DataTable ReadCSV()
        {
            DataTable dtResult = new DataTable();

            using (StreamReader SRFile = new StreamReader(_FilePath, Encoding.Default))
            {
                using (var csvReader = new CsvReader(SRFile, CultureInfo.InvariantCulture))
                {
                    using (var dr = new CsvDataReader(csvReader))
                    {
                        dtResult.Columns.Clear();
                        //    var dt = new DataTable();
                        dtResult.Load(dr);

                    

                    }
                }
            }
            return dtResult;
        }

        private string CreateXLSX(DataTable dtCsv)
        {
           
            string name = Path.GetFileNameWithoutExtension(_FilePath);
            FileInfo fi = new FileInfo(_FilePath);
            string xlsxFile = fi.Directory.FullName+"\\"+name+".xlsx";
            Dictionary<int, int> colTypes = new Dictionary<int, int>();
            int r = 1, c = 1;
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(new FileInfo(xlsxFile)))
                {
                    var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                    foreach (DataColumn col in dtCsv.Columns)
                    {
                        sheet1.Cells[r, c].Value = col.ColumnName.ToString();
                        colTypes[c] = 0;
                        c++;
                    }

                    if (columnSettingDele != null)
                    {
                        columnSettingDele(ref colTypes);
                    }

                    r = 2; c = 1;
                    foreach (DataRow row in dtCsv.Rows)
                    {
                        foreach (DataColumn col in dtCsv.Columns)
                        {
                            //默认是0，不做处理。 1 代表数字  
                            if (colTypes[c] == 1)
                                sheet1.Cells[r, c].Value = Convert.ToInt32(row[col]);
                            else
                                sheet1.Cells[r, c].Value = row[col];
                            c++;
                        }
                        r++;
                        c = 1;
                    }
                  //  sheet1.
                    package.Save();

                }
            }
            catch(Exception ex)
            {
                throw new Exception($"r:{r},c:{c}");
            }
           
            
            return xlsxFile;

        }

   

       // public Action<ExcelWorksheet> AdjustColumns;

        public void CSVToXLSX(string filePath)
        {
            _FilePath = filePath;

           var dt =  ReadCSV();
            if (AdjustData != null)
                AdjustData(dt);
            CreateXLSX(dt);
        }
    }
}
