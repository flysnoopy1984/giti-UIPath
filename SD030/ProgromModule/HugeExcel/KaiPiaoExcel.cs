
using OfficeOpenXml;
using RPA.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace HugeExcel
{
    public class KaiPiaoExcel
    {
        public static RPACore _RPACore = RPACore.getInstance();

        public string _FilePath;

        private DataTable _dataTable;

        public DataTable DataTable
        {
            get
            {
                return _dataTable;
            }
        }
        public KaiPiaoExcel()
        {
           
        }

        public void Init()
        {
            string fn = $"KaiPiao_{DateTime.Now.ToString("yyyyMMdd")}.csv";
            _FilePath = _RPACore.Configuration["HugeExcel:downloadDir"] + fn;

            string excelFile = _RPACore.Configuration["HugeExcel:downloadDir"]+ $"KaiPiao_{DateTime.Now.ToString("yyyyMMdd")}.xlsx";
            if (File.Exists(excelFile))
            {
                File.Delete(excelFile);
            }

            Program.HistoryFileList.Add(_FilePath);
            Program.HistoryFileList.Add(excelFile);
        }


        public bool Run()
        {
            try
            {
                Init();
                if (File.Exists(_FilePath))
                {
                    CSVExcelConverter converter = new CSVExcelConverter();
                    converter.AdjustData = AdjustData;
                    converter.columnSettingDele = ColumnSettingDelegate;
                    converter.CSVToXLSX(_FilePath);
                    return true;
                }
                return false;
            }
            catch(Exception ex)
            {
                return false;
            }

         
        }

        public void ColumnSettingDelegate(ref Dictionary<int, int> colsType)
        {
            colsType[7] = 1;
            colsType[8] = 1;
            colsType[9] = 1;
        }

      

        public void AdjustData(DataTable dt)
        {
           var col1 = dt.Columns.Add("add1");
           var col2 =  dt.Columns.Add("add2");
           DataRow row = null;
            

            foreach (System.Data.DataColumn col in dt.Columns) 
                col.ReadOnly = false;

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                for(int c = 0; c < dt.Columns.Count; c++)
                {
                    //月份
                    if(c == 7)
                    {
                        var month = dt.Rows[r][c].ToString();
                        month = month.Substring(0,2).Trim();
                        row = dt.Rows[r];
                        row.BeginEdit();
                        row[c] = Convert.ToInt32(month);
                        row.EndEdit();
                     
                    }
                    //新增列
                    if (c == 9)
                    {
                        row = dt.Rows[r];
                        row.BeginEdit();
                        dt.Rows[r][c] = "开票";
                        row.EndEdit();
                     //   dt.AcceptChanges();
                    }  
                    if (c == 10)
                    {
                        row = dt.Rows[r];
                        row.BeginEdit();
                        dt.Rows[r][c] = "开票实际";
                        row.EndEdit();
                     //   dt.AcceptChanges();
                    }
                   
                }
            }

            dt.AcceptChanges();

            _dataTable = dt;

       
        }

    }
}
