using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace CSVAnalysis
{
    public class AnalyInfoPath
    {
        private string filePath = @"C:\Project\UIPath\SD030\ProgromModule\CSVAnalysis\test.xml";
        public void run()
        {
            XmlDocument xml = new XmlDocument();
            xml.LoadXml(filePath);
        }
    }
}
