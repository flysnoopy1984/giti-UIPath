using RPA.Core;
using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;
using System.Data;
using System.Drawing;

namespace HugeExcel
{
    public class FinalExcel
    {
        public static RPACore _RPACore = RPACore.getInstance();

        private string _FilePath;
        private DirectoryInfo _fileDir;
        public FinalExcel()
        {
            InitFilePath();
        }

        public string FilePath
        {
            get
            {
                return _FilePath;
            }
        }

        public string DirPath
        {
            get
            {
                return _fileDir.FullName;
            }
        }

        public void InitFilePath()
        {
            var dirPath = _RPACore.Configuration["HugeExcel:finalDir"];
            _fileDir = new DirectoryInfo(dirPath);
            var files = _fileDir.GetFiles();
            if (files.Length > 0)
            {
                _FilePath = files[0].FullName;
            }
          
        }

        //表-数量
        private void FillBaseData(ExcelPackage package, DataTable dt)
        {
            var sheet = package.Workbook.Worksheets["表-数量"];
            var addr = sheet.Dimension.Address;

            var startRow = sheet.Dimension.End.Row;
            int dtRow = 0;

            for (int r = startRow + 1; r <=startRow + dt.Rows.Count; r++)
            {

                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    if (c == 12 || c == 13 || c == 14)
                        sheet.Cells[r, c + 1].Value = Convert.ToInt32(dt.Rows[dtRow][c]);
                    else
                        sheet.Cells[r, c + 1].Value = dt.Rows[dtRow][c];

                    ExcelRange cOld = sheet.Cells[2, c + 1];
                    ExcelRange cNew = sheet.Cells[r, c + 1];
                    cNew.StyleID = cOld.StyleID;
                }


                //季度开始,按公式生成数据
                for (int nc = dt.Columns.Count + 1; nc < sheet.Dimension.End.Column; nc++)
                {
                    ExcelRange cOld = sheet.Cells[2, nc];
                    ExcelRange cNew = sheet.Cells[r, nc];
                    cNew.StyleID = cOld.StyleID;
                    cNew.FormulaR1C1 = cOld.FormulaR1C1;
                }
                dtRow++;
            }
        }

        //图-半钢胎 //图-全钢胎
        private void SetMainStyleAndFormula(ExcelPackage package,ExcelWorksheet sheet)
        {
            int FirstMonthRow = 7;
            int curMonth = DateTime.Now.Month-1;
            int curMonthRow = FirstMonthRow + curMonth;
            int curCol = 22;
            int colorStartCol = 14;
            int c=0;
            ExcelRange cOld = null;
            ExcelRange cNew = null;
            if (curMonthRow == 8)
            {
                //TODO Send Email
                throw new Exception("第一个月无法处理");
            }
            try
            {
              
                for (c = curCol; c < 43; c++)
                {
                    cOld = sheet.Cells[curMonthRow - 1, c];
                    cNew = sheet.Cells[curMonthRow, c];
                    if (!string.IsNullOrEmpty(cOld.FormulaR1C1) && string.IsNullOrEmpty(cNew.FormulaR1C1))
                    {
                        cNew.FormulaR1C1 = cOld.FormulaR1C1;
                    }
                  
                }
                for(c = colorStartCol; c < 43; c++)
                {
                    cOld = sheet.Cells[curMonthRow - 1, c];
                    cNew = sheet.Cells[curMonthRow, c];
                    if (cOld.Style.Fill.BackgroundColor.Rgb != null)
                    {
                        cNew.StyleID = cOld.StyleID;

                        cOld.Style.Fill.BackgroundColor.SetColor(Color.White);
                    }
                }
           
                curMonthRow += 23;
                for (c = curCol; c < 41; c++)
                {
                    cOld = sheet.Cells[curMonthRow - 1, c];
                    cNew = sheet.Cells[curMonthRow, c];
                    if (!string.IsNullOrEmpty(cOld.FormulaR1C1) && string.IsNullOrEmpty(cNew.FormulaR1C1))
                    {
                        cNew.FormulaR1C1 = cOld.FormulaR1C1;
                    }
                }

                for (c = colorStartCol; c < 41; c++)
                {
                    cOld = sheet.Cells[curMonthRow - 1, c];
                    cNew = sheet.Cells[curMonthRow, c];
                    if (cOld.Style.Fill.BackgroundColor.Rgb != null)
                    {
                        cNew.StyleID = cOld.StyleID;

                        cOld.Style.Fill.BackgroundColor.SetColor(Color.White);
                    }
                }

            }
            catch(Exception ex)
            {
                Console.WriteLine($"SetMainSheets Error: row：{curMonthRow}  col:{c}");
                throw ex;
            }
            //图-全钢胎
            //     sheet = package.Workbook.Worksheets["图-全钢胎"];
            //     sheet.Cells[curMonthRow, curCol].Formula = $"=G{curMonthRow}";

        }

        public void Run(DataTable dt)
        {
            if (!string.IsNullOrEmpty(_FilePath))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(new FileInfo(_FilePath)))
                {
                
                     FillBaseData(package, dt);
                  
                    var sheet = package.Workbook.Worksheets["图-半钢胎"];
                    SetMainStyleAndFormula(package, sheet);
                    sheet = package.Workbook.Worksheets["图-全钢胎"];
                    SetMainStyleAndFormula(package, sheet);
                    package.Save();
                }
            }
        }
    }
}
