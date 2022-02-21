using System;
using System.IO;

namespace TestUtil
{
    class Program
    {
        static void Main(string[] args)
        {
            DateTime cur = DateTime.Now;
            string d = cur.AddDays(1 - cur.Day).ToString("yyyy-MM-dd");

             cur = DateTime.Now;
            string e = cur.AddDays(1 - cur.Day).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
            //   string d =  DateTime.Now.ToString("yyyy-MM-01");
            System.IO.DirectoryInfo di = new DirectoryInfo("c:\\test");
            if (di.Exists)
                di.Delete(true);
            di.Create();

            Console.WriteLine(d);
            Console.WriteLine(e);
            Console.WriteLine("Hello World!");
        }
    }
}
