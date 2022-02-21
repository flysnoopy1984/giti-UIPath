using OfficeOpenXml;
using RPA.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace SalesPre
{
    public class CombineExcel
    {
        public static RPACore _RPACore = RPACore.getInstance();
        public string _FilePath;

        private DataTable _FileDataTable; 

        public DataTable FileDataTable
        {
            get { return _FileDataTable; }
        }
        public CombineExcel()
        {
            string fn = $"ComBine_{DateTime.Now.ToString("yyyyMMdd")}.xlsx";
            _FilePath = _RPACore.Configuration["SalesPre:downloadDir"] + fn;

            Program.HistoryFileList.Add(_FilePath);
         
        }

        private void Init()
        {
            if (File.Exists(_FilePath))
            {
                File.Delete(_FilePath);
            }
        }

        

        public DataTable Run_GetTableData(DataRow[] rows, int colLen)
        {

         
            DataTable dt = new DataTable();
            colLen =25;
            for(int i = 0; i < colLen; i++)
            {
                dt.Columns.Add();
            }
            for(int r = 0; r <rows.Length; r++)
            {
                DataRow row = dt.NewRow();
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                   
                    if (c >= 0 && c <= 8)
                        row[c] = "";
                    else if (c == 9)
                        row[c] = "内销配套";
                    else if (c == 24)
                        row[c] = "有效";
                    else if (c == 25)
                        row[c] = DateTime.Now.ToString("yyyy/M/dd");
                    else
                        row[c] = rows[r][c-6];
                  
                }
                dt.Rows.Add(row);
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
                //写入Excel
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                foreach (DataRow row in rows)
                {   
                    sheet1.Cells[r, 10].Value = "内销配套";
                    sheet1.Cells[r, 24].Value = "有效";
                    sheet1.Cells[r, 25].Value = DateTime.Now.ToString("yyyy/M/dd");
                    c = 12; // 11 汽车集团
                    for (var j= 0; j < colLen; j++)
                    {
                        if(j == 6 || j == 7 || j ==8)
                            sheet1.Cells[r, c].Value = Convert.ToInt32(row[j]);
                        else
                            sheet1.Cells[r, c].Value = row[j];
                        if (c == 19) c++;
                        c++;
                    }
                    r++;
                }

                //写入DataTable

                _FileDataTable = new DataTable();
                int allColNum = sheet1.Dimension.End.Column;
                for (int i = 0; i < allColNum; i++)
                {
                    _FileDataTable.Columns.Add();
                }

                r = sheet1.Dimension.Start.Row;
                while (r<= sheet1.Dimension.End.Row)
                {
                    var newRow = _FileDataTable.NewRow();
                     c = 1;
                    while (c <= allColNum)
                    {
                        newRow[c-1] = sheet1.Cells[r, c].Value;
                        c++;
                    }
                    _FileDataTable.Rows.Add(newRow);
                    r++;
                }
             
                package.Save();

              
            }
            return true;
       }
    }
}
