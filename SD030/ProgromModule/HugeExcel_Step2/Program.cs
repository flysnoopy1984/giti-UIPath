using RPA.Core;
using System;
using System.IO;

namespace HugeExcel_Step2
{
    class Program
    {
        public static RPACore _RPACore = RPACore.getInstance();
        static ResultExcel resultExcel = null;
        static void Main(string[] args)
        {
           
            try
            {
                _RPACore.InitSystem(typeof(Program));

                resultExcel = new ResultExcel();
                //  finalExcel.Test_DrawLine();
                resultExcel.Run();
           //     CopyFileToTwo();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }

        static void CopyFileToTwo()
        {
            FileInfo fi = new FileInfo(resultExcel.FilePath);
            fi.CopyTo(resultExcel.DirPath + "\\report1.xlsx");
            fi.CopyTo(resultExcel.DirPath + "\\report2.xlsx");
        }
    }
}
