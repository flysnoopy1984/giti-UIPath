using Microsoft.Extensions.Configuration;
using RPA.Core;
using System;
using System.IO;

namespace CSVAnalysis
{
    class Program
    {
     
        public static RPACore _RPACore = RPACore.getInstance();
        static void Main(string[] args)
        {
            try
            {
                _RPACore.InitSystem(typeof(Program),false);
      //          if (Convert.ToBoolean(_RPACore.Configuration["setting:KillIEprocess"]))
            //    BaseTask.KillProcess("iexplore");

             //   Console.WriteLine($"start Root :{_RPACore.CurrentDirectory}");
                SD030Task task = new SD030Task();
                task.run();
                if (Convert.ToBoolean(_RPACore.Configuration["setting:KillIEprocess"]))
                    BaseTask.KillProcess("iexplore");
                Console.WriteLine($"Done");

                if (Convert.ToBoolean(_RPACore.Configuration["setting:showEnding"]))
                    Console.ReadLine();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
           
        }


    }
}
