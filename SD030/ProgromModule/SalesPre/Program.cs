using RPA.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace SalesPre
{
    class Program
    {
        public static RPACore _RPACore = RPACore.getInstance();

        private static FileMonitor _FileMonitor ;
        public static List<string> HistoryFileList = new List<string>();
        private static string _HistoryDir;
        public static string HistoryDir
        {
            get
            {
                if (string.IsNullOrEmpty(_HistoryDir))
                {
                    _HistoryDir = _RPACore.Configuration["SalesPre:historyDir"];
                    _HistoryDir += DateTime.Now.ToString("yyyy_MM_dd");
                    DirectoryInfo dir = new DirectoryInfo(_HistoryDir);
                    if (!dir.Exists)
                        dir.Create();
                }
                return _HistoryDir;
            }
        }

        public static bool IsTest
        {
            get
            {
                return Convert.ToBoolean(_RPACore.Configuration["SalesPre:IsTest"]);
            }
        }
        static void Main(string[] args)
        {
            try
            {
                _RPACore.InitSystem(typeof(Program), false);

                string filePath = _RPACore.Configuration["SalesPre:monitorFile"];
                _FileMonitor = new FileMonitor(filePath);
                _FileMonitor.StartMonitor();

              // bool isTest = Convert.ToBoolean(_RPACore.Configuration["SalesPre:IsTest"]);
                //   Test();
              //  if (!IsTest)
                    RunWithTime();
          
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                _FileMonitor.EndMonitor();
            }
        }

       
      
        public static void RunWithTime()
        {
            var startTime = DateTime.Now;
            string log = "用时:";

        

            StepOne();

            var endTime= DateTime.Now;
            TimeSpan ts = endTime - startTime;
            double diffSec = ts.TotalSeconds;


            if (diffSec < 60)
                log += $"{diffSec} 秒";
            else
            {
                log += $"{diffSec / 60} 分 {diffSec % 60} 秒";
            }
            Console.WriteLine(log);
        }

        public static void Test()
        {
            int startSec = DateTime.Now.Second;
            string log = "用时:";
       
        
            int endSec = DateTime.Now.Second;
            int diffSec = endSec - startSec;
            if (diffSec < 60)
                log += $"{diffSec} 秒";
            else
            {
                log += $"{diffSec / 60} 分 {diffSec % 60} 秒"; 
            }
            Console.WriteLine(log);
        }

        public static void StepOne()
        {
            bool result = false;
            try
            {
                NewCustomerExcel newCustomerExcel = new NewCustomerExcel();
                newCustomerExcel.Run();
            //    return;
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

                    int dataCount = combineExcel.FileDataTable.Rows.Count;
                    Console.WriteLine($"ComBine Count:{dataCount}");

                    RollPlanExcel rollPlanExcel = new RollPlanExcel();
                    rollPlanExcel.Run(combineExcel.FileDataTable);
                    int added = combineExcel.FileDataTable.Rows.Count - dataCount;
                    dataCount = combineExcel.FileDataTable.Rows.Count;
                    Console.WriteLine($"After RollPlan Count:{dataCount}. Added:{added}");
                  
                    //Process Excel
                    ProcessExcel processExcel = new ProcessExcel();
                    processExcel.Run(combineExcel.FileDataTable, added>0);
                  //  processExcel.RemoveOldPlanData();
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
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }

        private static void EndStep1Program()
        {
            foreach (var filepath in HistoryFileList)
            {
                MoveToHistory(filepath);
            }

            //将处理后的数据文件Copy到Result文件夹中(Step再删除)
            //ResultExcel resultExcel = new ResultExcel();
            //resultExcel.DeleteDirFIles();

            //var dirResultPath = resultExcel.DirPath;
            //var finalfilePath = _finalExcel.FilePath;
            //FileInfo fi = new FileInfo(finalfilePath);
            //var name = fi.Name;
            //fi.CopyTo(dirResultPath + "\\" + name);
        }

        public static void MoveToHistory(string filePath)
        {
            FileInfo fi = new FileInfo(filePath);
            if (fi.Exists)
            {
                var targetFile = HistoryDir + "\\" + fi.Name;
                File.Move(fi.FullName, targetFile, true);
            }
        }
    }
}
