using RPA.Core;
using System;

namespace HugeExcel
{
    class Program
    {
        public static RPACore _RPACore = RPACore.getInstance();
        static void Main(string[] args)
        {
            try
            {
                _RPACore.InitSystem(typeof(Program));

                KaiPiaoExcel kaiPiaoExcel = new KaiPiaoExcel();
                kaiPiaoExcel.Test();
               // NLogUtil.cc_InfoTxt("test");
               //  _RPACore.Db.Queryable("TPRT_Client_TYPE", "tct");
                Console.WriteLine("Done");
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
          
            //     string s = "54455371417145580116401361072062165223103482211721783466421773985821043329861";

            //   Console.WriteLine(s.Length.ToString());
            //  Console.ReadLine();
        }
    }
}
