using RPA.Core;
using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Collections.Generic;

namespace HugeExcel
{
    class Program
    {
        public static RPACore _RPACore = RPACore.getInstance();

        public static List<string> HistoryFileList = new List<string>();

        private static string _HistoryDir;
        public static string HistoryDir
        {
            get
            {
                if (string.IsNullOrEmpty(_HistoryDir))
                {
                    _HistoryDir = _RPACore.Configuration["HugeExcel:historyDir"];
                    _HistoryDir += DateTime.Now.ToString("yyyy_MM_dd");
                    DirectoryInfo dir = new DirectoryInfo(_HistoryDir);
                    if (!dir.Exists)
                        dir.Create();  
                }
                return _HistoryDir;
            }
        }

        public static void MoveToHistory(string filePath)
        {
            FileInfo fi = new FileInfo(filePath);
            if (fi.Exists)
            {
                var targetFile = HistoryDir + "\\" + fi.Name;
                File.Move(fi.FullName, targetFile,true);
            }
        }

        private static void EndProgram()
        {
            foreach(var filepath in HistoryFileList)
            {
                MoveToHistory(filepath);
            }
        }
        static void Main(string[] args)
        {
            try
            {
                _RPACore.InitSystem(typeof(Program));
                bool result = false;


                /* Test Begin */
                //FinalExcel finalExcel = new FinalExcel();
                //   finalExcel.Test();
                // finalExcel.Run(null);
                /* Test End */

                KaiPiaoExcel kaiPiaoExcel = new KaiPiaoExcel();
                result = kaiPiaoExcel.Run();

                FaYunExcel faYunExcel = new FaYunExcel();
                result = faYunExcel.Run();

                if (result)
                {

                    DataRow[] rows = new DataRow[faYunExcel.DataTable.Rows.Count + kaiPiaoExcel.DataTable.Rows.Count];
                    faYunExcel.DataTable.Rows.CopyTo(rows, 0);
                    kaiPiaoExcel.DataTable.Rows.CopyTo(rows, faYunExcel.DataTable.Rows.Count);

                    CombineExcel combineExcel = new CombineExcel();
                    int colLen = kaiPiaoExcel.DataTable.Columns.Count;
                    combineExcel.Run(rows, colLen);
                    var dt = combineExcel.Run_GetTableData(rows, colLen);


                    FinalExcel finalExcel = new FinalExcel();
                    finalExcel.Run(dt);

                    EndProgram();
                    //    combineExcel.Run(rows, colLen);
                }

                Console.WriteLine("Done");
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
    }
}
