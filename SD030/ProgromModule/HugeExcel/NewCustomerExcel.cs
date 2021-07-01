using OfficeOpenXml;
using RPA.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace HugeExcel
{
    public class NewCustomerExcel
    {
        public static RPACore _RPACore = RPACore.getInstance();

        public List<string> _FilePathList = new List<string>();

        public NewCustomerExcel()
        {
            var dirPath = _RPACore.Configuration["HugeExcel:newCustomerDir"];
            var fileDir = new DirectoryInfo(dirPath);
            var files = fileDir.GetFiles();
            if (files.Length > 0)
            {
                _FilePathList.Add(files[0].FullName);
            }
        }

        public void Run()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FinalExcel finalExcel = new FinalExcel();
            if (_FilePathList.Count > 0)
            {
                using (ExcelPackage packageFinal = new ExcelPackage(new FileInfo(finalExcel.FilePath)))
                {
                    var sheetFinal = packageFinal.Workbook.Worksheets["客户清单"];
                    var targetRow = sheetFinal.Dimension.End.Row;

                    foreach (var excelFilePath in _FilePathList)
                    {
                        using (ExcelPackage packageCustomer = new ExcelPackage(new FileInfo(excelFilePath)))
                        {
                            var sheetCustomer = packageCustomer.Workbook.Worksheets[0];
                            var rowCount = sheetCustomer.Dimension.End.Row;
                            var columnCount = sheetCustomer.Dimension.End.Column;

                            for (int r = 2; r <= rowCount; r++)
                            {
                                targetRow++;
                                for (int c = 1; c <= columnCount; c++)
                                {
                                    sheetFinal.Cells[targetRow, c].Value = sheetCustomer.Cells[r, c].Value;
                                }
                            }

                            //var testDataSheet = packageCustomer.Workbook.Worksheets["TestData"];
                            //sheetFinal = packageFinal.Workbook.Worksheets["表-数量"];
                            //targetRow = sheetFinal.Dimension.End.Row;
                            //if (testDataSheet != null)
                            //{
                            //    rowCount = testDataSheet.Dimension.End.Row;
                            //    columnCount = testDataSheet.Dimension.End.Column;
                            //    for (int r = 1; r <= rowCount; r++)
                            //    {
                            //        targetRow++;
                            //        for (int c = 1; c <= columnCount; c++)
                            //        {
                            //            sheetFinal.Cells[targetRow, c].Value = testDataSheet.Cells[r, c].Value;
                            //        }
                            //    }
                            //}
                        }
                    }
                    packageFinal.Save();

                }
            }
          


          
            
        }
    }
}
