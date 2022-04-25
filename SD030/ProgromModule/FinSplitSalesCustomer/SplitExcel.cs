using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using RPA.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace FinSplitSalesCustomer
{
    public class SplitExcel
    {
        private string _ExcelPath;
        private string _MapPath;
        private string _SentDir;

        private List<SalesCustomer> salesCustomers { get; } = new List<SalesCustomer>();

        private List<EmailSales> emailSalesList { get; } = new List<EmailSales>();
       
        private Dictionary<string, SalesData> salesDatas= new Dictionary<string, SalesData>();

        public SplitExcel(string runApp)
        {
            Init(runApp);
        }

        private void Init(string runApp)
        {
           if(runApp == "regunship")
                _ExcelPath = Program._RPACore.Configuration["historyDir"]+DateTime.Now.ToString("yyyyMMdd")+ "\\RegUnShip.xlsx";
           else
                _ExcelPath = Program._RPACore.Configuration["historyDir"] + DateTime.Now.ToString("yyyyMMdd") + "\\exceldata.xlsx";

            _SentDir = Program._RPACore.Configuration["historyDir"] +  DateTime.Now.ToString("yyyyMMdd") + "\\sent";
            DirectoryInfo di = new DirectoryInfo(_SentDir);
            if (di.Exists)
                di.Delete(true);
           
             di.Create();
            _MapPath = Program._RPACore.Configuration["mappingFile"];


        }

     
        private void initMapping()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage packMap = new ExcelPackage(new FileInfo(_MapPath)))
            {
                var sheet = packMap.Workbook.Worksheets["Data"];
                var rowCount = sheet.Dimension.End.Row;

                for (int r = 2; r <= rowCount; r++)
                {
                    string cc = Convert.ToString(sheet.Cells[r, 1].Value);
                    string sm = Convert.ToString(sheet.Cells[r, 3].Value);
                    var getOne = salesCustomers.Find(sc => sc.CusCode == cc && sc.SalesMail == sm);
                    if (getOne == null)
                        salesCustomers.Add(new SalesCustomer
                        {
                            CusCode = cc,
                            SalesMail = sm
                        });
                }
            }
        }

        private void splitSalesCustomerData()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage packMap = new ExcelPackage(new FileInfo(_ExcelPath)))
            {
                var sheet = packMap.Workbook.Worksheets["data"];
                var rowCount = sheet.Dimension.End.Row;

                for (int r = 3; r <= rowCount; r++)
                {

                    string cc = Convert.ToString(sheet.Cells[r, 1].Value);
                    //根据客户寻找对应的销售
                    var scList = salesCustomers.FindAll(sc => sc.CusCode == cc);
                    if (scList != null)
                    {
                        foreach (var sc in scList)
                        {
                            //数据中是否存在此销售
                            try
                            {
                                salesDatas[sc.SalesMail].rowIndexList.Add(r);
                            }
                            catch
                            {
                                var sd = new SalesData();
                                sd.rowIndexList.Add(r);
                                sd.SalesMail = sc.SalesMail;
                                salesDatas.Add(sc.SalesMail, sd);
                            }
                        }
                    }
                }

                createEmailFiles(sheet);
            }
        }

        private void splitRegData()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage packMap = new ExcelPackage(new FileInfo(_ExcelPath)))
            {
                var sheet = packMap.Workbook.Worksheets["Sheet1"];
                var rowCount = sheet.Dimension.End.Row;

                for (int r = 2; r <= rowCount; r++)
                {
                    
                    string cc = Convert.ToString(sheet.Cells[r, 9].Value);
                    //根据客户寻找对应的销售
                    var scList = salesCustomers.FindAll(sc => sc.CusCode == cc);
                    if (scList != null)
                    {
                        foreach(var sc in scList)
                        {
                            //数据中是否存在此销售
                            try
                            {
                                salesDatas[sc.SalesMail].rowIndexList.Add(r);
                            }
                            catch
                            {
                                var sd = new SalesData();
                                sd.rowIndexList.Add(r);
                                sd.SalesMail = sc.SalesMail;
                                salesDatas.Add(sc.SalesMail, sd);
                            }
                        }
                    }
                }

                createEmailFiles(sheet);
            }
        }

        private void CopyRow(ExcelWorksheet source, ExcelWorksheet target,int sourceIndex,int targetIndex)
        {
            var colCount = source.Dimension.End.Column;
            for (int c = 1; c <= colCount; c++)
            {
                target.Cells[targetIndex, c].Value = source.Cells[sourceIndex, c].Value;
             }
        }
      
        private void createEmailFiles(ExcelWorksheet dataSheet)
        {
        
            foreach (var sd in this.salesDatas)
            {
                var name = sd.Key.Split('@')[0];
                string newExcelPath = _SentDir + "\\" + name + ".xlsx";
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage packNew = new ExcelPackage(new FileInfo(newExcelPath)))
                {
                    var newSheet = packNew.Workbook.Worksheets.Add(name);
                    var newRow = 2;
                    //如果有数据
                    if (sd.Value.rowIndexList.Count > 0)
                    {
                        //Copy 头
                        CopyRow(dataSheet, newSheet, 1,1);
                        foreach (int r in sd.Value.rowIndexList)
                        {
                            CopyRow(dataSheet, newSheet, r, newRow);
                            newRow++;
                        }
                        emailSalesList.Add(new EmailSales
                        {
                            SalesMail = sd.Key,
                            AttachFilePath = newExcelPath,
                        });
                        packNew.Save();
                    }
                }
            }
        }

        private void SentEmail()
        {
            IConfiguration cfg = Program._RPACore.Configuration;
            bool isTest = Convert.ToBoolean(cfg["IsTest"]);
            bool is126 = Convert.ToString(cfg["Email:fromAddr"]).Split("@")[1] == "126.com";
           
            if (isTest)
            {
                //List<string> sentList = new List<string>();
                //List<string> attachList = new List<string>();
                //List<string> bccList = new List<string>();
                //string path = RPAZip.ZipDir(_SentDir);
              
                //sentList.Add("song.fuwei@giti.com");
                //attachList.Add(path);
                //if(is126)
                //    RPAEmail.Sent_126(sentList, Convert.ToString(cfg["Email:title"]), Convert.ToString(cfg["Email:body"]), null, null, attachList);
                //else
                //    RPAEmail.Sent(sentList, Convert.ToString(cfg["Email:title"]), Convert.ToString(cfg["Email:body"]), null, null, attachList);

            }
            else
            {
                foreach (var es in emailSalesList)
                {
                    List<string> sentList = new List<string>();
                    List<string> attachList = new List<string>();
                    List<string> bccList = new List<string>();

                    object bcc = cfg["Email:bcc"];
                    if (bcc!=null)
                    {
                        string bccStr = Convert.ToString(bcc);
                        if(!string.IsNullOrEmpty(bccStr))
                            bccList =  bccStr.Split(";").ToList();
                    }

                    sentList.Add(es.SalesMail);
                    attachList.Add(es.AttachFilePath);
                    Console.WriteLine(es.SalesMail);
                  //  RPAEmail.Sent(sentList, Convert.ToString(cfg["Email:title"]), Convert.ToString(cfg["Email:body"]), null, bccList, attachList);
                }
            }
           
        }

        public void runSalesCustomer()
        {
            if (File.Exists(_ExcelPath))
            {
                initMapping();

                splitSalesCustomerData();

                SentEmail();
            }
            else
            {
                Console.WriteLine($"{_ExcelPath} not got");
                Console.ReadLine();

            }
        }
        public void runRegUnShip()
        {
            if (File.Exists(_ExcelPath))
            {
                initMapping();

                splitRegData();

                SentEmail();
             }
            else
            {
                Console.WriteLine($"{_ExcelPath} not got");
                Console.ReadLine();
            }
        }
    }
}
