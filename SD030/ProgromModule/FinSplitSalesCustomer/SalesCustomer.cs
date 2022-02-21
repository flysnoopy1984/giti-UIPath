using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace FinSplitSalesCustomer
{
    public class EmailSales
    {
        public string SalesMail { get; set; }

        public string AttachFilePath { get; set; }
    }
    public class SalesData
    {
        public string SalesMail { get; set; }

        public List<int> rowIndexList { get; set; } = new List<int>();
    }
    public class SalesCustomer
    {
        public string CusCode { get; set; }

        public string SalesMail { get; set; }
    }
}
