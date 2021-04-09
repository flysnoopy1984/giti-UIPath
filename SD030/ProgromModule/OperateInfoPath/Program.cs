
using RPA.Core;
using System;

namespace OperateInfoPath
{
    class Program
    {
        public static RPACore _RPACore = RPACore.getInstance();
        static void Main(string[] args)
        {
            try
            {
                _RPACore.InitSystem(typeof(Program));
                OperateInfoPathTask task = new OperateInfoPathTask();
             //   task.test();
                task.run();
                Console.WriteLine("Done");
                if (Convert.ToBoolean(_RPACore.Configuration["setting:showEnding"]))
                    Console.ReadLine();
            }
            catch(Exception ex)
            {
                throw ex;
            }
        
        }
    }
}
