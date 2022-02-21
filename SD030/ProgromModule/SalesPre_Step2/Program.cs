using RPA.Core;
using System;
using System.IO;

namespace SalesPre_Step2
{
    class Program
    {

        public static RPACore _RPACore = RPACore.getInstance();

        private static FileMonitor _FileMonitor;

        static void Main(string[] args)
        {
            try
            {
                _RPACore.InitSystem(typeof(Program), false);
                string filePath = _RPACore.Configuration["SalesPre:monitorFile"];
                _FileMonitor = new FileMonitor(filePath);
                _FileMonitor.StartMonitor();

                run();

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


        public static void run()
        {
            var dirPath = _RPACore.Configuration["SalesPre:processDir"];
            var fileDir = new DirectoryInfo(dirPath);
            var files = fileDir.GetFiles();
            string sourceFilePath = null;
            foreach (FileInfo f in files)
            {
                if (f.Name.StartsWith("baseData"))
                {
                    sourceFilePath = f.FullName;
                }
            }

            DateTime dt = DateTime.Now.AddMonths(-1);

            ExcelCopy excelCopy = new ExcelCopy();
            string targetFilePath = _RPACore.Configuration["SalesPre:exportReportDir"] + $"配套市场销售预算、实际追踪分析-{dt.Year}年{dt.Month}月.xlsx";
            fileDir = new DirectoryInfo(_RPACore.Configuration["SalesPre:exportReportDir"]);
            files = fileDir.GetFiles();
            foreach (FileInfo f in files)
            {
                f.Delete();
            }     
            excelCopy.Run(sourceFilePath, targetFilePath);
        }
    }
}
