using RPA.Core;
using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;
using System.Data;
using System.Drawing;
using System.Linq;

namespace SalesPre
{
    public class ProcessExcel
    {
        public static RPACore _RPACore = RPACore.getInstance();

        private string _FilePath;
        private DirectoryInfo _fileDir;
        public ProcessExcel()
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
            var dirPath = _RPACore.Configuration["SalesPre:processDir"];
            _fileDir = new DirectoryInfo(dirPath);
            var files = _fileDir.GetFiles();
            foreach(FileInfo f in files)
            {
                if (f.Name.StartsWith("baseData"))
                {
                    _FilePath = f.FullName;
                }
            }
        }

        //表-数量
        private void FillBaseData(ExcelPackage package, DataTable dt)
        {
            var sheet = package.Workbook.Worksheets["数据"];
       //     var addr = sheet.Dimension.Address;

            var startRow = sheet.Dimension.End.Row;
            int dtRow = 0;
            int r = 0; int c =0;
            try
            {
                for (r = startRow + 1; r <= startRow + dt.Rows.Count; r++)
                {
                    Console.WriteLine($"Process On :{r}. num:{dtRow}");
                    for (c = 0; c < dt.Columns.Count; c++)
                    {
                        if (c == 17 || c == 18 || c == 20)
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
                        cNew.Calculate();
                    }
                    dtRow++;
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine($"get error in row:{r}-col:{c} --{ex.Message}");
            }
          
            
        }
        /// <summary>
        /// 1月 123 2月234 更新1月的23 11月 11 12 01 更新 11 12 12月 12 01 02 更新 
        /// </summary>
       public void UpdateOldPlanData(ExcelPackage package)
        {


            Console.WriteLine("Start RemoveOldPlanData");
            //  18 年 //19 月
            var sheet = package.Workbook.Worksheets["数据"];
            var allRow = sheet.Dimension.End.Row;
           
             int year = DateTime.Now.Year;
             int month = DateTime.Now.Month;
             int nextYear = DateTime.Now.AddYears(1).Year;
             int nextMonth = DateTime.Now.AddMonths(1).Month;

            /* 测试版本 */
            if (Program.IsTest)
            {
                year = DateTime.Now.Year;
                month = 8;
                nextYear = DateTime.Now.AddYears(1).Year;
                nextMonth = 9;
            }
      

            int removeNum = 0;
            for (int r = 2; r <= allRow; r++)
            {
                //数据细类
                if (Convert.ToString(sheet.Cells[r, 23].Value) == "滚动计划2版")
                {
                    // allRange.Start
                    int cellYear = Convert.ToInt32(sheet.Cells[r, 18].Value);
                    int cellMonth = Convert.ToInt32(sheet.Cells[r, 19].Value);
                    if (month == 12)
                    {
                        if ((cellYear == year && cellMonth == month) ||
                            (cellYear == nextYear && cellMonth == 1))
                        {

                            sheet.Cells[r, 24].Value = "失效";
                            Console.WriteLine($"Update Roll:{r}. Year-{cellYear},Month-{cellMonth}");
                            removeNum++;
                        }
                    }
                    else
                    {
                        if ((cellYear == year && cellMonth == month) ||
                            (cellYear == year && cellMonth == nextMonth))
                        {

                            sheet.Cells[r, 24].Value = "失效";
                            Console.WriteLine($"Update Roll:{r}. Year-{cellYear},Month-{cellMonth}");
                            removeNum++;
                        }

                    }
                }
              
            }

            Console.WriteLine($"End UpdatedOldPlanData.Update Num:{removeNum}");


        }

        public void Run(DataTable dt, bool hasPlanData = false)
        {
            if (!string.IsNullOrEmpty(_FilePath))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(new FileInfo(_FilePath)))
                {

                    if (hasPlanData)
                    {
                        UpdateOldPlanData(package);
                    }

                    FillBaseData(package, dt);

                    UpdateNumberData(package);

                    UpdateRedColumn(package);

                    package.Save();
                }
            }
        }

     
        //更新数字行，控制月份的那行数据
        public void UpdateNumberData(ExcelPackage package )
        {
            var sheetCustomer = package.Workbook.Worksheets["客户"];
         
            int curMonth = DateTime.Now.Month;
            int prevMonth = DateTime.Now.AddMonths(-1).Month;

            //Test Begin
            if (Program.IsTest)
            {
                 curMonth = 8;
                 prevMonth = 7;
            }
          
            //Test End


            //修改ET 4列<7
            var sheet = package.Workbook.Worksheets["集团"];
            var month = curMonth;
            if (month == 1) month = 13;
            string val = $"<{month}";
            sheet.Cells[4, 150].Value = val;
            sheet.Cells[4, 151].Value = val;
            sheet.Cells[4, 152].Value = val;
            sheet.Cells[4, 153].Value = val;
            sheet.Cells[7, 150].Value = $"1—{prevMonth} 月累计（实际）";

            sheetCustomer.Cells[5, 158].Value = val;
            sheetCustomer.Cells[5, 159].Value = val;
            sheetCustomer.Cells[5, 160].Value = val;
            sheetCustomer.Cells[5, 161].Value = val;
            sheetCustomer.Cells[8, 158].Value = $"1—{prevMonth} 月累计（实际）";

            month = curMonth + 3; //比如8月跑的是7月实际+3个月滚动
            if (month > 13) month = 13;
            val = $"<{month}";
            sheet.Cells[4, 165].Value = val;
            sheet.Cells[4, 166].Value = val;
            sheet.Cells[4, 167].Value = val;
            sheet.Cells[4, 168].Value = val;

            sheetCustomer.Cells[5, 173].Value = val;
            sheetCustomer.Cells[5, 174].Value = val;
            sheetCustomer.Cells[5, 175].Value = val;
            sheetCustomer.Cells[5, 176].Value = val;

            month--;
            if (month >= 13) month = 12;
            sheet.Cells[7, 165].Value = $"1—{month} 月累计（实际+滚动）";
            sheetCustomer.Cells[8, 173].Value = $"1—{month} 月累计（实际+滚动）";

         
        }

        private void UpdateRedColumn(ExcelPackage package)
        {
            var groupSheet = package.Workbook.Worksheets["集团"];
            var custSheet = package.Workbook.Worksheets["客户"];
           

            Dictionary<int, string> custMap = new Dictionary<int, string>();
            Dictionary<int, string> groupMap = new Dictionary<int, string>();
            groupMap.Add(1, "BQ");
            groupMap.Add(2, "BX");
            groupMap.Add(3, "CE");
            groupMap.Add(4, "CL");
            groupMap.Add(5, "CS");
            groupMap.Add(6, "CZ");
            groupMap.Add(7, "DG");
            groupMap.Add(8, "DN");
            groupMap.Add(9, "DU");
            groupMap.Add(10, "EB");
            groupMap.Add(11, "EI");
            groupMap.Add(12, "EP");

            custMap.Add(1, "BY");
            custMap.Add(2, "CF");
            custMap.Add(3, "CM");
            custMap.Add(4, "CT");
            custMap.Add(5, "DA");
            custMap.Add(6, "DH");
            custMap.Add(7, "DO");
            custMap.Add(8, "DV");
            custMap.Add(9, "EC");
            custMap.Add(10, "EJ");
            custMap.Add(11, "EQ");
            custMap.Add(12, "EX");

          
            SetRedColumnFormula(groupMap, groupSheet, "EV", 169);

            SetRedColumnFormula(custMap, custSheet, "FD", 177);
        }
    
        private void SetRedColumnFormula(Dictionary<int, string> map,ExcelWorksheet sheet,string actTxt,int targetCol)
        {
           
            var month = DateTime.Now.Month;

            //测试-Begin
            if (Program.IsTest)
                 month = 8;
            //测试-End
            for (int r = 10; r < sheet.Dimension.End.Row; r++)
            {
                string tr = r.ToString();
                var formula = "";
                if (month > 1 && month <= 10)
                {
                    formula = $"{actTxt+tr}+{map[month] +tr}+{map[month + 1] + tr}+{map[month + 2] + tr}";
                }
                else if (month == 11)
                {
                    formula = $"{actTxt + tr}+{map[month] + tr}+{map[month + 1] + tr}";
                }
                else if (month == 12)
                {
                    formula = $"{actTxt + tr}+{map[month] + tr}";
                }
                else if (month == 1)
                {
                    formula = $"{actTxt + tr}";
                }
                sheet.Cells[r, targetCol].Formula = formula;
            }
        }
    }
}
