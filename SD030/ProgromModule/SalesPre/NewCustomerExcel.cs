using OfficeOpenXml;
using RPA.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace SalesPre
{
    public class NewCustomerExcel
    {
        public static RPACore _RPACore = RPACore.getInstance();

        private string _FilePath_Customer = null;
        private string _FilePath_Process = null;


        public NewCustomerExcel()
        {
       
        }

        public void InitFilePath()
        {
            var dirPath = _RPACore.Configuration["SalesPre:processDir"];
            var  fileDir = new DirectoryInfo(dirPath);
            var files = fileDir.GetFiles();
            foreach (FileInfo f in files)
            {
                if (f.Name.StartsWith("customer"))
                {
                    _FilePath_Customer = f.FullName;

                    Program.HistoryFileList.Add(_FilePath_Customer);
                }

                if (f.Name.StartsWith("baseData"))
                {
                    _FilePath_Process = f.FullName;
                }
            }
        }

        public void Run()
        {
            InitFilePath();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    
            if (!string.IsNullOrEmpty(_FilePath_Customer))
            {
                using (ExcelPackage packageFinal = new ExcelPackage(new FileInfo(_FilePath_Process)))
                {
                    var sheetFinal = packageFinal.Workbook.Worksheets["表-客户清单"];
                    var targetRow = sheetFinal.Dimension.End.Row;

                    using (ExcelPackage packageCustomer = new ExcelPackage(new FileInfo(_FilePath_Customer)))
                    {
                        var sheetCustomer = packageCustomer.Workbook.Worksheets[0];

                        var rowCount = sheetCustomer.Dimension.End.Row;
                        var columnCount = sheetCustomer.Dimension.End.Column;

                     
                        for (int r = 2; r <= rowCount; r++)
                        {
                            targetRow++;
                            if (sheetCustomer.Cells[r, 1].Value == null) break;
                            for (int c = 1; c <= columnCount; c++)
                            {
                                sheetFinal.Cells[targetRow, c].Value = sheetCustomer.Cells[r, c].Value;
                            }
                        }
                    }

                    packageFinal.Save();

                }
            }
          


          
            
        }
    }
}
