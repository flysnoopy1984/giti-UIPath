using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using SqlSugar;

namespace RPA.Core
{
    public class RPACore
    {
        private  IConfigurationBuilder _ConfigurationBuilder = new ConfigurationBuilder();
        private IConfiguration _configuration;
        private string _CurrentDirectory;
        private SqlSugarClient _Db = null;

        public IConfiguration Configuration
        {
            get { return _configuration; }
        }

        public SqlSugarClient Db
        {
            get
            {
                return _Db;
            }
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
        //    FileInfo fi = new FileInfo(@"C:\Project\UIPath\ImportUserToWeb_Inner\Do\HistoryFiles\Inner\B2B客户新建账号--07 03 _1.xlsx");
         //   string ee = fi.Extension;
        }
        public  void InitSystem(Type type,bool needDB=true) 
        {
           
            _CurrentDirectory = Path.GetDirectoryName(type.Assembly.Location);
            Console.WriteLine("CurrentDirectory:" + _CurrentDirectory);
            _configuration = _ConfigurationBuilder.SetBasePath(_CurrentDirectory)
                  .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                  .Build();

            if(needDB)
                SqlSugarSetup();
        
        }

        private void SqlSugarSetup()
        {
            var appcfg = _configuration.GetSection("DataBases").GetChildren();
            List<ConnectionConfig> ccList = new List<ConnectionConfig>();

            foreach (var cfg in appcfg)
            {
                ConnectionConfig cc = new ConnectionConfig
                {
                    DbType = DbType.SqlServer,
                    ConfigId = cfg["ConfId"],
                    IsAutoCloseConnection = true,
                
                    ConnectionString = cfg["Connection"],

                };
                ccList.Add(cc);
            }
            _Db = new SqlSugarClient(ccList[0]);



        }
    }
}
