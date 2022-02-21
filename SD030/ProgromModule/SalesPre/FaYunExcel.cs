using RPA.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.IO;

namespace SalesPre
{
    
    public class FaYunExcel
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
        public FaYunExcel()
        {
           
        }

        private void Init()
        {
            string fn = $"FaHuo_{DateTime.Now.ToString("yyyyMMdd")}.csv";
            _FilePath = _RPACore.Configuration["SalesPre:downloadDir"] + fn;

            string excelFile = _RPACore.Configuration["SalesPre:downloadDir"] + $"FaHuo_{DateTime.Now.ToString("yyyyMMdd")}.xlsx";
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
            var col2 = dt.Columns.Add("add2");
            DataRow row = null;
            DataColumn delCol = dt.Columns[6];
            foreach (System.Data.DataColumn col in dt.Columns)
                col.ReadOnly = false;

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {

                    //新增列
                    if (c == 10)
                    {
                        row = dt.Rows[r];
                        row.BeginEdit();
                        dt.Rows[r][c] = "发货";
                        row.EndEdit();
                        //   dt.AcceptChanges();
                    }
                    if (c == 11)
                    {
                        row = dt.Rows[r];
                        row.BeginEdit();
                        dt.Rows[r][c] = "发货实际";
                        row.EndEdit();
                        //   dt.AcceptChanges();
                    }

                }
            }

            dt.AcceptChanges();
            dt.Columns.Remove(delCol);

            _dataTable = dt;
        }


    }
}
