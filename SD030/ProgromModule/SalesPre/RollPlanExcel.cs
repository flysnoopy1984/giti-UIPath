using OfficeOpenXml;
using RPA.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace SalesPre
{
    public class RollPlanExcel
    {
        public static RPACore _RPACore = RPACore.getInstance();

        private string _FilePath_RollPlan= null;

        //private DataTable _FileDataTable;

        //public DataTable FileDataTable
        //{
        //    get { return _FileDataTable; }
        //}

        //将要合并的数据
        private DataTable _ComBineDataTable;
        //  private string _FilePath_BaseData;

        private DirectoryInfo _fileDir;

        public void InitFilePath()
        {
            var dirPath = _RPACore.Configuration["SalesPre:processDir"];
            _fileDir = new DirectoryInfo(dirPath);
            var files = _fileDir.GetFiles();
            foreach (FileInfo f in files)
            {
                if (f.Name.StartsWith("rollPlan"))
                {
                    _FilePath_RollPlan = f.FullName;

                    Program.HistoryFileList.Add(_FilePath_RollPlan);
                }
            }
        }

        //和Combine DataTable 保持一致
        public void Run(DataTable comBineDataTable)
        {
            _ComBineDataTable = comBineDataTable;

            InitFilePath();

            if (!string.IsNullOrEmpty(_FilePath_RollPlan))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(new FileInfo(_FilePath_RollPlan)))
                {

                    GenDataTable(package);
                   // package.Save();
                }
            }
        }

        public void GenDataTable(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets["Sheet1"];    
            int r = 0; int c = 0;
            /*正式版   */
            var curMonth = DateTime.Now.Month;
            var curYear = DateTime.Now.Year;
            var nextYear = DateTime.Now.AddYears(1).Year;
            var nextMonth = DateTime.Now.AddMonths(1).Month;
            var nextnextMonth = DateTime.Now.AddMonths(2).Month;

            if (Program.IsTest)
            {
                 curMonth = 8;
                 curYear = 2021;
                 nextYear = 2022;
                 nextMonth = 9;
                 nextnextMonth = 10;
            }
          
            var endRowNum  = sheet.Dimension.End.Row;   
            for (r = 3; r <= endRowNum; r++)
            {
                if (sheet.Cells[r, 4].Value == null || string.IsNullOrEmpty(sheet.Cells[r, 4].Value.ToString())) 
                    continue;
                for(int i = 0; i < 3; i++)
                {
                    int year = curYear;
                    int month = curMonth;
                    object qty = null;
                    
                    var newRow = _ComBineDataTable.NewRow();
                    newRow[9] = "内销配套";
                    newRow[11] = sheet.Cells[r, 2].Value; //区域/分部
                    newRow[12] = sheet.Cells[r, 4].Value; //客户编码
                    newRow[13] = sheet.Cells[r, 5].Value; //客户名称
                    newRow[14] = "101半钢外胎";//sheet.Cells[r, 2]; //产品类别 ??
                    newRow[15] = sheet.Cells[r, 8].Value; //标准品号
                    newRow[16] = sheet.Cells[r, 9].Value; //描述

                    newRow[21] = "发货";// sheet.Cells[r, 2]; //数据大类
                    newRow[22] = "滚动计划2版";// sheet.Cells[r, 2]; //数据细类 
                    newRow[23] = "有效"; //有效性
                    newRow[24] = DateTime.Now.ToString("yyyy/M/dd"); //更新日期

                    if (i == 0)
                    {
                        qty = sheet.Cells[r, 17].Value;
                    }
                    else if (i == 1)
                    {

                        qty = sheet.Cells[r, 18].Value;
                        if (curMonth == 12)
                            month = 1;
                        else
                        {
                            month = nextMonth;
                        }
                      
                    }
                    else if (i == 2)
                    {
                        qty = sheet.Cells[r, 19].Value;
                        if (curMonth == 11)
                        {
                            year = nextYear;
                            month = 1;
                        }  
                        if (curMonth == 12)
                        {
                            year = nextYear;
                            month = 2;
                        }
                        else
                        {
                            month = nextnextMonth;
                        }
                    }
                    newRow[17] = year;// sheet.Cells[r, 2]; //年
                    newRow[18] = month; //月
                    newRow[19] = ""; //周
                    newRow[20] = qty; //数量

                    _ComBineDataTable.Rows.Add(newRow);
                }
            }
        }
    }
}
