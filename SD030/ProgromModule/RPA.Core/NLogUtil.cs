using NLog;
using System;
using System.Collections.Generic;
using System.Text;

namespace RPA.Core
{
    public static class NLogUtil
    {

        private static Logger _FileLogger_cc = LogManager.GetLogger("ccInfoLog");
        private static Logger _FileErrorLogger_cc = LogManager.GetLogger("ccErrorLog");

        public static void cc_InfoTxt(string txt, bool isPrint = true)
        {
            try
            {
                _FileLogger_cc.Info(txt);
                if (isPrint) Console.WriteLine(txt);
            }
            catch (Exception ex)
            {

            }


        }

        public static void cc_ErrorTxt(string txt, bool isPrint = true)
        {
            try
            {
                _FileErrorLogger_cc.Error(txt);
                if (isPrint) Console.WriteLine(txt);
            }
            catch (Exception ex)
            {

            }
        }

       
    }
}
