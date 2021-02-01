using System;
using System.IO;
using Microsoft.Extensions.Configuration;
namespace RPA.Core
{
    public class RPACore
    {
        private  IConfigurationBuilder _ConfigurationBuilder = new ConfigurationBuilder();
        private IConfiguration _configuration;
        private string _CurrentDirectory;

        public IConfiguration Configuration
        {
            get { return _configuration; }
        }

        public string CurrentDirectory
        {
            get { return _CurrentDirectory; }
        }
   //     public IConfigurationBuilder

        private static RPACore _instance = null;
        public static RPACore getInstance()
        {
            if (_instance == null) 
                _instance = new RPACore();
            return _instance;
        }
        private RPACore()
        {
          
        }
        public  void InitSystem(Type type) 
        {
           
            _CurrentDirectory = Path.GetDirectoryName(type.Assembly.Location);
            Console.WriteLine("CurrentDirectory:" + _CurrentDirectory);
            _configuration = _ConfigurationBuilder.SetBasePath(_CurrentDirectory)
                  .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                  .Build();
        }
    }
}
