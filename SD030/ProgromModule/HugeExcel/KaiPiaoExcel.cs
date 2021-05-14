
using OfficeOpenXml;
using RPA.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace HugeExcel
{
    public class KaiPiaoExcel
    {
        public static RPACore _RPACore = RPACore.getInstance();

        public string _FilePath;
        public KaiPiaoExcel()
        {
            string fn = $"KaiPiao_{DateTime.Now.ToString("yyyyMMdd")}.xlsx";
            _FilePath = _RPACore.Configuration["HugeExcel:downloadDir"]+fn;
         }

        public void Test()
        {
            //    ExcelTools.CSVSaveasXLSX(_FilePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new FileStream(_FilePath, FileMode.Open)))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets[1];
                for (int row = 2; row < sheet.Dimension.End.Row; row++)
                {
                    string val = (string)sheet.Cells[row, 8].Value;
                    var t = val.Substring(0, 1);
                }
            }
        }

    }
}
