using Microsoft.Extensions.Configuration;
using RPA.Core;
using System;
using System.Collections.Generic;

namespace FinSplitSalesCustomer
{
    class Program
    {
        public static RPACore _RPACore = RPACore.getInstance();
        private static FileMonitor _FileMonitor;

        static void Main(string[] args)
        {
            try
            {
                _FileMonitor = new FileMonitor();
                _RPACore.InitSystem(typeof(Program), false);
                _FileMonitor.StartMonitor();

                var runApp = Convert.ToString(_RPACore.Configuration["runApp"]).ToLower();
              

                SplitExcel splitExcel = new SplitExcel(runApp);
                if (runApp == "regunship")
                    splitExcel.runRegUnShip();
                else
                    splitExcel.runSalesCustomer();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if(_FileMonitor !=null)
                    _FileMonitor.EndMonitor();
            }
         
            //  TestZip();

            //SplitExcel splitExcel = new SplitExcel();
            //splitExcel.run();
        }
        static void TestEmail()
        {

        }
        static void TestZip()
        {
            IConfiguration cfg = Program._RPACore.Configuration;
            string path = RPAZip.ZipDir(@"C:\Project\UIPath\SD030\ProgromModule\FinSplitSalesCustomer\History\20220216\sent\");
            List<string> sentList = new List<string>();
            List<string> attachList = new List<string>();
            sentList.Add("song.fuwei@giti.com");
            attachList.Add(path);

            RPAEmail.Sent(sentList, Convert.ToString(cfg["Email:title"]), Convert.ToString(cfg["Email:body"]), null, null, attachList);
        }
    }
}
