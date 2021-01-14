using Microsoft.Extensions.Configuration;
using RPA.Core;
using System;
using System.IO;

namespace CSVAnalysis
{
    class Program
    {
        //public static readonly IConfigurationBuilder ConfigurationBuilder = new ConfigurationBuilder();
        //public static IConfiguration _configuration;
        //public static string _CurrentDirectory;
        public static RPACore _RPACore = RPACore.getInstance();
        static void Main(string[] args)
        {
            try
            {
                _RPACore.InitSystem(typeof(Program));

                Console.WriteLine($"start Root :{_RPACore.CurrentDirectory}");
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

        //public static void InitSystem()
        //{

        //    //    dynamic type = (new Program()).GetType();
        //    dynamic type = typeof(Program);
        //    _CurrentDirectory = Path.GetDirectoryName(type.Assembly.Location);

        //    _configuration = ConfigurationBuilder.SetBasePath(_CurrentDirectory)
        //          .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
        //          .Build();
        //}


    }
}
