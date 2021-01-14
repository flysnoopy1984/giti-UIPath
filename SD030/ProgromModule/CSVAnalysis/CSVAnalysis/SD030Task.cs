using CsvHelper;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Linq;
using System.Net.Mail;
using System.Diagnostics;
using System.ComponentModel;
using RPA.Core;

namespace CSVAnalysis
{
    public class SD030Task: BaseTask
    {
        private string _runRoot;
        private string _excelDir,_historyDir;
     
        private DateTime _startDate;
        private DateTime _endDate;
        private List<SC030CSVEntity> _result;
        private FileInfo _todoExcel;
        public static RPACore _RPACore = RPACore.getInstance();

        public SD030Task()
        {
            _runRoot = _RPACore.CurrentDirectory; //Program._CurrentDirectory;

            int y = DateTime.Now.Year;
            int  m = DateTime.Now.Month;
            int d = this.getLastDayOfMonth(DateTime.Now).Day;
            _endDate = DateTime.Parse(y +"-"+ m +"-"+ d);
            int beforeMonth = Convert.ToInt32(_RPACore.Configuration["setting:beforeMonth"]);
            var tempDate = _endDate.AddMonths(-beforeMonth);
         
            _startDate = DateTime.Parse(tempDate.Year + "-" + tempDate.Month + "-" + 1);

            var dir = _RPACore.Configuration["setting:excelDir"];
            string parentDir = Path.GetFullPath("..");
            _excelDir = parentDir + "\\" + dir;
            dir = _RPACore.Configuration["setting:historyDir"];
            _historyDir = parentDir + "\\" + dir;

        }


        public bool FindAvaliableExcel()
        {
            DirectoryInfo di = new DirectoryInfo(_excelDir);
            var files = di.GetFiles();
            foreach(var f in files)
            {
                if(this.verifyTodoFile(f))
                {
                    _todoExcel = f;
                    return true;
                }
            }
            return false;
        }
        public void run()
        {
            if (this.FindAvaliableExcel())
            {
                _result = this.readCSV(_todoExcel.FullName);
                this.writeCSV();

                var historyFileFullName = _historyDir + _todoExcel.Name;

                File.Move(_todoExcel.FullName, historyFileFullName, true);
            }
            else
            {
                throw new Exception("没有找到合适的文件");
            }

          
        }

        private DateTime getLastDayOfMonth(DateTime datetime)
        {
            return datetime.AddDays(1 - datetime.Day).AddMonths(1).AddDays(-1);
        }

        private bool verifyTodoFile(FileInfo fi)
        {
            var name = Path.GetFileNameWithoutExtension(fi.Name);
            if (name.EndsWith("_result"))
            {
              //  Console.WriteLine("没有找到Excel");
                return false;
            }
            return true;
        }

        private List<SC030CSVEntity> readCSV(string filePath)
        {
            using (StreamReader SRFile = new StreamReader(filePath, Encoding.Default))
            {
                using (var csvReader = new CsvReader(SRFile, CultureInfo.InvariantCulture))
                {

                    var datas = csvReader.GetRecords<SC030CSVEntity>().ToList();
                    Console.WriteLine($"读取文件行--{datas.Count}");
                    _result = datas.FindAll(a => Convert.ToDateTime(a.a1) >= _startDate && Convert.ToDateTime(a.a1) <= _endDate);
                    Console.WriteLine($"过滤结果--{_result.Count}");
                }
            }
          
            if (_result == null)
                _result = new List<SC030CSVEntity>();
            return _result;
        }

        private void writeCSV()
        {
            if (_result == null || _result.Count == 0) 
                return;
            var name = Path.GetFileNameWithoutExtension(_todoExcel.Name);

            var outputFile = _excelDir + name + "_result.csv";

            using (var writer = new StreamWriter(outputFile,false,Encoding.UTF8))
            {
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    csv.WriteHeader<SC030CSVEntity>();
                    csv.NextRecord();
                    foreach (var record in _result)
                    {
                        csv.WriteRecord(record);
                        csv.NextRecord();
                    }
                    Console.WriteLine($"生成结果文件");
                }
            }
        }

     
       
    }
}
