using RPA.Core;
using System;
using System.Collections.Generic;
using System.IO;

namespace TestUtil
{
    class Program
    {
        public static RPACore _RPACore = RPACore.getInstance();
        static void Main(string[] args)
        {
            //     _RPACore.InitSystem(typeof(Program), false);

            //     List<string> sentList = new List<string>();
            //     sentList.Add("song.fuwei@giti.com");

            ////     var a = Convert.ToString(_RPACore.Configuration["Email:fromAddr"]);
            //    // bool is126 = Convert.ToString(_RPACore.Configuration["Email:fromAddr"]).Split("@")[1] == "126.com";

            //     RPA.Core.RPAEmail.Sent_126(sentList,"aa","bb");
            //     Console.WriteLine("Hello World!");
            TestFile();
        }

        static void TestFile()
        {
            FileInfo fi = new FileInfo(@"c:\1.xls");
            if (fi.Exists)
            {

            }
        }
    }
}
