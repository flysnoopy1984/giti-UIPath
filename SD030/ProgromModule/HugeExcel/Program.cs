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

        private static FinalExcel _finalExcel;
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

        private static void EndStep1Program()
        {
            foreach(var filepath in HistoryFileList)
            {
                MoveToHistory(filepath);
            }

            //将处理后的数据文件Copy到Result文件夹中(Step再删除)
            ResultExcel resultExcel = new ResultExcel();
            resultExcel.DeleteDirFIles();

            var dirResultPath = resultExcel.DirPath;
            var finalfilePath = _finalExcel.FilePath;
            FileInfo fi = new FileInfo(finalfilePath);
            var name = fi.Name;
            fi.CopyTo(dirResultPath + "\\"+name);
        }

        public static void StepOne()
        {
            bool result = false;
            try
            {
                /* 可以使用 但暂时不用*/
                NewCustomerExcel newCustomerExcel = new NewCustomerExcel();
                newCustomerExcel.Run();
     

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


                    _finalExcel = new FinalExcel();
                    _finalExcel.Run(dt);

                    EndStep1Program();
                    //    combineExcel.Run(rows, colLen);
                }
                else
                {
                    Console.WriteLine("没有Download 文件！");
                    Console.ReadLine();
                }

                Console.WriteLine("Done");
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }

        private static void StepTwo()
        {
            ResultExcel resultExcel = new ResultExcel();
            resultExcel.InitFilePath();
           // resultExcel.RunStepTwo();

            FinalExcel finalExcel = new FinalExcel();
            
            //将Result的文件移动到FinalFile,由RPA清理
            FileInfo fi = new FileInfo(resultExcel.FilePath);
            fi.CopyTo(finalExcel.DirPath + "\\"+fi.Name,true);
            
            Console.WriteLine("Done");
        }

        static void Main(string[] args)
        {
            try
            {
                _RPACore.InitSystem(typeof(Program));

                StepOne();
                //ResultExcel resultExcel = new ResultExcel();
                //if (!resultExcel.ExistFilePath())
                //    StepOne();
                //else
                //    StepTwo();


            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
    }
}
