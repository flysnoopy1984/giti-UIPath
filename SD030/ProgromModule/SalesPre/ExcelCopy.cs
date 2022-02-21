using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SalesPre
{
    public class ExcelCopy
    {
        public void Run(string sourceFile,string targetFile)
        {
            if (File.Exists(targetFile))
            {
                File.Delete(targetFile);
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage targetPackage = new ExcelPackage(new FileInfo(targetFile)))
            {
                var targetGroup = targetPackage.Workbook.Worksheets.Add("集团");
                var targetCust = targetPackage.Workbook.Worksheets.Add("客户");

                ExcelPackage sourcePackage = null;
           
                try
                {
                    sourcePackage = new ExcelPackage(new FileInfo(sourceFile));
                    var sourceGroup = sourcePackage.Workbook.Worksheets["集团"];
                    var sourceCust = sourcePackage.Workbook.Worksheets["客户"];

                 
                    CopySheet(sourceGroup, targetGroup, sourcePackage,targetPackage);
                    CopySheet(sourceCust, targetCust, sourcePackage, targetPackage);

                   

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    if (sourcePackage != null)
                    {
                        sourcePackage.Dispose();
                    }
                }

                targetPackage.Save();
            }
        }
    
        public void CopySheet(ExcelWorksheet sourceSheet, ExcelWorksheet targetSheet, ExcelPackage sourcePackage, ExcelPackage targetPackage)
        {
            int allRowNum = sourceSheet.Dimension.End.Row;
            int allColNum = sourceSheet.Dimension.End.Column;

            ExcelRange cOld = sourceSheet.Cells[1, 1, allRowNum, allColNum];
            ExcelRange cNew = targetSheet.Cells[1, 1, allRowNum, allColNum];

            cOld.Copy(cNew, ExcelRangeCopyOptionFlags.ExcludeFormulas);


            for (int r = 1; r <= allRowNum; r++)
            {
                targetSheet.Row(r).Height = sourceSheet.Row(r).Height;
            }

            for (int c = 1; c <= allColNum; c++)
            {

                targetSheet.Column(c).Width = sourceSheet.Column(c).Width;

            }


        }
    }
}
