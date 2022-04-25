using OfficeOpenXml;
using RPA.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace HugeExcel
{
    public class CombineExcel
    {
        public static RPACore _RPACore = RPACore.getInstance();
        public string _FilePath;
        public CombineExcel()
        {
            string fn = $"ComBine_{DateTime.Now.ToString("yyyyMMdd")}.xlsx";
            _FilePath = _RPACore.Configuration["HugeExcel:downloadDir"] + fn;

            Program.HistoryFileList.Add(_FilePath);
         
        }

        private void Init()
        {
            if (File.Exists(_FilePath))
            {
                File.Delete(_FilePath);
            }
        }
        public DataTable Run_GetTableData_2022(DataRow[] rows, int colLen)
        {


            DataTable dt = new DataTable();
            colLen += 8;
            for (int i = 0; i < colLen; i++)
            {
                dt.Columns.Add();
            }
            for (int r = 0; r < rows.Length; r++)
            {
                DataRow row = dt.NewRow();
                for (int c = 0; c < dt.Columns.Count; c++)
                {

                    if (c >= 0 && c <= 4)
                        row[c] = "";
                    else if (c == 5)
                        row[c] = "内销配套";
                    else if (c == 17)
                        row[c] = "有效";
                    else if (c == 18)
                        row[c] = DateTime.Now.ToString("yyyy/M/dd");
                    else
                        row[c] = rows[r][c - 6];

                }
                dt.Rows.Add(row);
            }
            return dt;

        }

        public DataTable Run_GetTableData(DataRow[] rows, int colLen)
        {

         
            DataTable dt = new DataTable();
            colLen += 8;
            for(int i = 0; i < colLen; i++)
            {
                dt.Columns.Add();
            }
            for(int r = 0; r <rows.Length; r++)
            {
                DataRow row = dt.NewRow();
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                   
                    if (c >= 0 && c <= 4)
                        row[c] = "";
                    else if (c == 5)
                        row[c] = "内销配套";
                    else if (c == 17)
                        row[c] = "有效";
                    else if (c == 18)
                        row[c] = DateTime.Now.ToString("yyyy/M/dd");
                    else
                        row[c] = rows[r][c-6];
                  
                }
                dt.Rows.Add(row);
            }
            return dt;
          
        }

        public DataTable Run_2022(DataRow[] rows, int colLen)
        {
            Init();

            DataTable dt = new DataTable();
            int dtColNum = 17;  //17 写死的直接看Combine 得出

            for (int i = 0; i < dtColNum; i++)
            {
                dt.Columns.Add();
            }

            int r = 1, c = 5;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(_FilePath)))
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
             

                foreach (DataRow row in rows)
                {
                    c = 5;
                    sheet1.Cells[r, 4].Value = "内销配套";
                    sheet1.Cells[r, 16].Value = "有效";
                    sheet1.Cells[r, 17].Value = DateTime.Now.ToString("yyyy/M/dd");
                    for (var j = 0; j < colLen; j++)
                    {
                        if (j == 6 || j == 7 || j == 8)
                            sheet1.Cells[r, c].Value = Convert.ToInt32(row[j]);
                        else
                            sheet1.Cells[r, c].Value = row[j];
                        c++;
                    }
                  
                    //写入内存DataTable
                    DataRow dtRow = dt.NewRow();
                    for (var dtc = 0; dtc < dtColNum; dtc++)
                    {
                        dtRow[dtc] = sheet1.Cells[r, dtc + 1].Value;
                    }
                    dt.Rows.Add(dtRow);

                    r++;
                }
                package.Save();


            }
            return dt;
        }

        public bool Run(DataRow[] rows, int colLen)
        {
            Init();

            int r = 1,c =7;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(_FilePath)))
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                foreach (DataRow row in rows)
                {
                    c = 7;
                    sheet1.Cells[r, 6].Value = "内销配套";
                    sheet1.Cells[r, 18].Value = "有效";
                    sheet1.Cells[r, 19].Value = DateTime.Now.ToString("yyyy/M/dd");
                    for (var j= 0; j < colLen; j++)
                    {
                        if(j == 6 || j == 7 || j ==8)
                            sheet1.Cells[r, c].Value = Convert.ToInt32(row[j]);
                        else
                            sheet1.Cells[r, c].Value = row[j];
                        c++;
                    }
                    r++;
                }
                package.Save();

              
            }
            return true;
       }
    }
}
