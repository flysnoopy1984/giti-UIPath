using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace RPA.Core
{
    public class FileMonitor
    {
        private FileInfo _MonitorFile;

        public FileMonitor(string filePath = null)
        {
            if(string.IsNullOrEmpty(filePath))
                _MonitorFile = new FileInfo(Environment.CurrentDirectory + "\\mon.sfw");
            else
                _MonitorFile = new FileInfo(filePath);
        }
       
        public void StartMonitor()
        {
         
            if (!_MonitorFile.Exists)
            {
                using (var fs = _MonitorFile.Create())
                {
                    fs.Close();
                }
                _MonitorFile.Refresh();

            }
        }

        public void EndMonitor()
        {
        
            File.Delete(_MonitorFile.FullName);
         
        }
    }
}
