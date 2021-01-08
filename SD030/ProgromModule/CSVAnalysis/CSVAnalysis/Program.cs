using Microsoft.Extensions.Configuration;
using System;
using System.IO;

namespace CSVAnalysis
{
    class Program
    {
        public static readonly IConfigurationBuilder ConfigurationBuilder = new ConfigurationBuilder();
        public  static IConfiguration _configuration;
        public static string _CurrentDirectory;

        static void Main(string[] args)
        {
            try
            {
       
                InitSystem();
                Console.WriteLine($"start Root :{_CurrentDirectory}");
                SD030Task task = new SD030Task();
                task.run();
                if (Convert.ToBoolean(_configuration["setting:KillIEprocess"]))
                    BaseTask.KillProcess("iexplore");
                Console.WriteLine($"Done");

                if(Convert.ToBoolean(_configuration["setting:showEnding"]))
                     Console.ReadLine();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
           
        }

        public static void InitSystem()
        {
            dynamic type = (new Program()).GetType();
            _CurrentDirectory = Path.GetDirectoryName(type.Assembly.Location);

            _configuration = ConfigurationBuilder.SetBasePath(_CurrentDirectory)
                  .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                  .Build();
        }


    }
}
