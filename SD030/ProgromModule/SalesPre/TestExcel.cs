using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SalesPre
{
    public class TestExcel
    {
        public void Test()
        {
            FileInfo fi = new FileInfo(@"C:\Project\UIPath\SD030\ProgromModule\SalesPre\Process\abc.xlsx");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(fi))
            {
                var sheet = package.Workbook.Worksheets["Sheet1"];
                sheet.Cells[1, 4].Formula = "A1+F1";
                sheet.Cells[2, 4].Formula = "A2+F2";

                package.Save();
            }
        }
    }
}
