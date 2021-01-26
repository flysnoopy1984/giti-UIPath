
using RPA.Core;
using System;

namespace OperateInfoPath
{
    class Program
    {
        public static RPACore _RPACore = RPACore.getInstance();
        static void Main(string[] args)
        {
            _RPACore.InitSystem(typeof(Program));
            OperateInfoPathTask task = new OperateInfoPathTask();
            task.run();
            Console.WriteLine("Done");
            Console.Read();
        }


    }
}
