using System;
using System.Diagnostics;
using System.Reflection;


namespace RPA.Core
{
    public static class ExcelTools
    {

        private static void QuertExcel()
        {
            Process[] excels = Process.GetProcessesByName("EXCEL");
            foreach (var item in excels)
            {
                item.Kill();
            }
        }

        public static string CSVSaveasXLSX(string FilePath)
        {
            return null;
       
        }
    }
}
