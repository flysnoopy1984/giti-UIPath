using OfficeOpenXml;
using RPA.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace HugeExcel_Step2
{
    public class ResultExcel
    {
        public static RPACore _RPACore = RPACore.getInstance();

        private string _FilePath;
        private DirectoryInfo _fileDir;
        public string DirPath
        {
            get
            {
                return _fileDir.FullName;
            }
        }

        public string FilePath
        {
            get { return _FilePath; }
        }
       

        public ResultExcel()
        {
            try
            {
                NLogUtil.cc_InfoTxt("Step2 ResultExcel Start");
                InitFilePath();
            }
            catch(Exception ex)
            {
                NLogUtil.cc_ErrorTxt("Step2 InitFilePath: " + ex.Message);
                throw new Exception("InitFilePath Error" + ex.Message);
            }
          
        }
        public void InitFilePath()
        {
            var dirPath = _RPACore.Configuration["HugeExcel:resultDir"];

            _fileDir = new DirectoryInfo(dirPath);
        
            var files = _fileDir.GetFiles();
            if (files.Length > 0)
            {
                _FilePath = files[0].FullName;
            }

        }

        public void Test_DrawLine()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(_FilePath)))
            {
                var sheet = package.Workbook.Worksheets[0];
                sheet.Cells[93, 14].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                sheet.Cells[93, 14].Style.Border.Left.Color.SetColor(Color.Black);
                sheet.Cells[94, 14].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                sheet.Cells[94, 14].Style.Border.Left.Color.SetColor(Color.Black);
                package.Save();

            }
       }

       public void CopyDrawLine(ExcelWorksheet sheet, int r, int c, int copyRow,bool isEnd = false)
        {
            try
            {
                sheet.Cells[r, c].StyleID = sheet.Cells[copyRow, c].StyleID;
                //   sheet.Cells[r, c].Style.Border = sheet.Cells[copyRow, c].Style.Border;
                if (isEnd)
                {
                    if (c == 217 || c == 200 || c == 185 || c == 170 || c == 155 || c == 140 || c == 125 || c == 110 || c == 95 || c == 80 || c == 65 || c == 50 || c == 35 || c == 20)
                        sheet.Cells[r, c].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
                    else
                        DrawLine(sheet, r, c, LinePos.Bottom);
                }
                else
                {
                    sheet.Cells[r, c].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;

                }
            }
            catch(Exception ex)
            {
                NLogUtil.cc_ErrorTxt("Step2 CopyDrawLine: " + ex.Message);
                throw ex;
            }
         
        }

        public void DrawLine(ExcelWorksheet sheet,int r,int c, LinePos linePos)
        {
            switch (linePos)
            {
                case LinePos.Left:
                    sheet.Cells[r, c].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    sheet.Cells[r, c].Style.Border.Left.Color.SetColor(Color.Black);
                    break;
                case LinePos.Bottom:
                    sheet.Cells[r, c].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    sheet.Cells[r, c].Style.Border.Bottom.Color.SetColor(Color.Black);
                    break;
                case LinePos.Right:
                    sheet.Cells[r, c].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    sheet.Cells[r, c].Style.Border.Right.Color.SetColor(Color.Black);
                    break;
                case LinePos.LeftBottom:
                    sheet.Cells[r, c].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    sheet.Cells[r, c].Style.Border.Left.Color.SetColor(Color.Black);
                    sheet.Cells[r, c].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    sheet.Cells[r, c].Style.Border.Bottom.Color.SetColor(Color.Black);
                    break;
                case LinePos.RightBottom:
                    sheet.Cells[r, c].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    sheet.Cells[r, c].Style.Border.Right.Color.SetColor(Color.Black);
                    sheet.Cells[r, c].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    sheet.Cells[r, c].Style.Border.Bottom.Color.SetColor(Color.Black);
                    break;
                default:break;
            }
          
        }

        private void ChangeParameterAndReCalcualte()
        {

        }

        public void Run()
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(new FileInfo(_FilePath)))
                {
                    var sheet = package.Workbook.Worksheets["A-全钢胎"];
                    int endRow = sheet.Dimension.Rows - 1;
                    int startRow = endRow;
                    int col = 19;
                    var val = sheet.Cells[startRow, col].Value as string;
                    while (string.IsNullOrEmpty(val))
                    {
                        val = sheet.Cells[--startRow, col].Value as string;
                    }
                    var copycol = 14;
                    var copyRow = startRow;
                    startRow = copyRow + 1;
                    while (startRow <= endRow)
                    {
                        for (int c = copycol; c <= sheet.Dimension.Columns; c++)
                        {
                            sheet.Cells[startRow, c].FormulaR1C1 = sheet.Cells[copyRow, c].FormulaR1C1;
                            //  sheet.Cells[startRow, c].Calculate();

                            CopyDrawLine(sheet, startRow, c, copyRow, startRow == endRow);

                        }
                        startRow++;
                    }

                    //设置月份参数，重新计算
                    sheet.Cells[1, 201].Value = DateTime.Now.AddMonths(-1).Month;
                    //   sheet.Cells.Calculate();

                    package.Save();
                }

            }
            catch(Exception ex)
            {
                NLogUtil.cc_ErrorTxt("Step2 run: " + ex.Message);
                throw new Exception("run Error:" + ex.Message);
            }
        
        }
    }
}
