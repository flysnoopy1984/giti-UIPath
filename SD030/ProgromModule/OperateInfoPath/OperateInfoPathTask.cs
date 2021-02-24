using Newtonsoft.Json;
using OfficeOpenXml;
using RPA.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Xml;

namespace OperateInfoPath
{
    public class OperateInfoPathTask:BaseTask
    {
        private XmlDocument _doc = new XmlDocument();
        private List<FileInfo> _xmlFilesList = new List<FileInfo>();
        private XmlNamespaceManager _nsmgr;
        public static RPACore _RPACore = RPACore.getInstance();

        public List<ccJsonEntity> _ccDataList = new List<ccJsonEntity>();

        private B2BJsonReault _B2BJsonReault = new B2BJsonReault();

        public const string gAddressType = "100000002";

        private ExcelWorksheet _SheetPC1, _SheetPC2, _SheetTB1, _SheetTB2,_SheetClientType;

        public void run()
        {
            //对照表
            InitCCDataJson();
            
            //找到需要转换的infoPath
            InitXmlFiles();
         
            //将一个或多个infoPath转换成B2BEntity对象
            ConvertInfoPathToTarget();

            //转换Json到需要上传到B2b格式的xml,保存
            if (Convert.ToBoolean(_RPACore.Configuration["setting:createFXml"]))
                SaveToXmlFile();

           SaveToCRMExcel();

            TempJsonExcel();
        }

        private void InitXmlFiles()
        {
            string dirInPath = _RPACore.Configuration["setting:infoPathInDir"];//  @"C:\Project\UIPath\SD030\ProgromModule\OperateInfoPath\test.xml";

            DirectoryInfo di = new DirectoryInfo(dirInPath);
            _xmlFilesList = di.GetFiles().ToList();

        }
        /// <summary>
        /// 初始化手工表，excel 已经转成json
        /// </summary>
        private void InitCCDataJson()
        {
            string jsonFile = _RPACore.CurrentDirectory + "\\ccDatas.json";
            string jsonStr = JsonHelper.GetJsonFile(jsonFile);
            _ccDataList = JsonConvert.DeserializeObject<List<ccJsonEntity>>(jsonStr); // JsonSerializer.Deserialize<List<ccJsonEntity>>(jsonStr);
        }


        public void ConvertInfoPathToTarget()
        {
           foreach(var fi in _xmlFilesList)
            {
                _doc.Load(fi.FullName);
                _nsmgr = new XmlNamespaceManager(_doc.NameTable);
                _nsmgr.AddNamespace("my", "http://schemas.microsoft.com/office/infopath/2003/myXSD/2017-10-17T09:00:01");

                var B2BObj = this.CreateB2BSheetEntity();
                _B2BJsonReault.Data.Add(B2BObj);
            }
        }

       //手工表中找到记录
        public ccJsonEntity FindinJsonDatas(string branch,bool isPC)
        {
            if (isPC)
            {
                return _ccDataList.Find(a => a.PCBranch == branch);
            }
            else
                return _ccDataList.Find(a => a.TBBranch == branch);
        }
        public B2BEntity CreateB2BSheetEntity()
        {
            B2BEntity obj = new B2BEntity();
            AddressEntity addrObj = new AddressEntity();
            try
            {
                addrObj.Addresstype = gAddressType;
                addrObj.Region = "China";
                addrObj.County = "中国";
                addrObj.Province = "";
                addrObj.City = "";
                addrObj.AddressNum = "1";
                addrObj.Addressline = getNodeInnerText("//my:收货地址");

                obj.AddressList.Add(addrObj);

                obj.AccountNum = getNodeInnerText("//my:ERP编码");
                obj.CustomerType = "1";
                obj.Name = getNodeInnerText("//my:营业执照名称"); 

                obj.BusinessContact = getNodeInnerText("//my:业务联系人");
                obj.PhoneNumber = getNodeInnerText("//my:联系人手机"); 

                obj.Ifpc = getNodeInnerText("//my:乘用or商用") == "乘用车胎" ? true : false;
                obj.Iftb = !obj.Ifpc;
                obj.Ifbias = false;
                
                dbssInfo info = new dbssInfo();

                info.district = getNodeInnerText("//my:我司信息/my:区域");
                info.branch = getNodeInnerText("//my:我司信息/my:分部");
                info.SalesContact = getNodeInnerText("//my:我司信息/my:销售代表");
                info.SalesEmail = getNodeInnerText("//my:我司信息/my:邮箱");

                var ccEntity = FindinJsonDatas(info.branch, obj.Ifpc);
                if (obj.Ifpc)
                {
                    obj.InitPCInfo(info, ccEntity);
                    obj.applyBrand = getNodeInnerText("//my:授权信息/my:申请乘用授权/my:乘用授权/my:乘用授权及目标/my:乘-品牌");
                }
                else
                {
                    obj.InitTbInfo(info, ccEntity);
                    obj.applyBrand = getNodeInnerText("//my:授权信息/my:申请商用授权/my:商用授权/my:商用授权及目标/my:商-品牌");
                }
                   
                obj.Status = "0";
                obj.Createdon = DateTime.Now.ToString("yyyy/MM/dd");
                obj.K2District = getNodeInnerText("//my:发运组织");

                #region  CRM
                obj.crmId = Guid.NewGuid().ToString()+"_"+obj.AccountNum;
                this.getAuthListData(obj);
                //string pathSQ = "//my:合同信息/my:合同授权/my:授权区域";
                //obj.crmBrand = getNodeInnerText($"{pathSQ }/ my:授权品牌");
                //obj.crmKind = getNodeInnerText("//my:乘用or商用");
                //obj.crmShen = getNodeInnerText($"{pathSQ }/my:省份");
                //obj.crmVMode = getNodeInnerText($"{pathSQ }/my:经营方式");
                //if(obj.crmJYFangShi == 1) obj.crmLinShouAddr = getNodeInnerText($"{pathSQ }/my:零售地址");
                //    obj.crmDI = getNodeInnerText($"{pathSQ }/my:地级市");
                //var IsAllXian = getNodeInnerText($"{pathSQ }/my:全部or部分区县") == "部分"?false:true;

                //if (IsAllXian) obj.crmXian = "所有县";
                //else obj.crmXian = this.getXianList();

                if (obj.Ifpc)
                {
                    getPCSheet2(obj);
                }
                else
                {
                    //obj.crmSCXFEN = this.getMarketList(obj);
                    //obj.crmSeries = "全系列";
                    //obj.crmLinShouAddr = "(空白)";

                    var hasKeYun = this.getNodeInnerText("//my:合同信息/my:目标-商用/my:商用目标/my:商用车胎目标/my:客运市场");
                    obj.hasKeYun = (!string.IsNullOrEmpty(hasKeYun) && hasKeYun == "客运");
                    if (obj.hasKeYun)
                    {
                        var xifen1 = this.getNodeInnerText("//my:客运授权/my:清单及考察期/my:客运细分1");
                        if(!string.IsNullOrEmpty(xifen1)) obj.KeYunXiFenList.Add(xifen1);

                        var xifen2 = this.getNodeInnerText("//my:客运授权/my:清单及考察期/my:客运细分2");
                        if (!string.IsNullOrEmpty(xifen2))  obj.KeYunXiFenList.Add(xifen2);

                        //   var hasKeYun = this.getNodeInnerText("//my:合同信息/my:目标-商用/my:商用目标/my:商用车胎目标/my:客运市场");
                    }
                    getTBSheet2(obj);
                }   
                #endregion

            }
            catch (Exception ex)
            {
            //    throw ex;
                Console.WriteLine(ex.Message);
            }
            return obj;
        }

        //合同授权信息
        private void getAuthListData(B2BEntity obj)
        {
            string pathSQ = "//my:合同信息/my:合同授权/my:授权区域";
            var authNodeList = _doc.SelectNodes(pathSQ, _nsmgr);
           for(int i = 0; i < authNodeList.Count; i++)
            {
                var authNode = authNodeList[i];
             //   var bd = this.getNodeValue(authNode.SelectSingleNode("//my:授权品牌",_nsmgr));
                AuthEntity authEntity = new AuthEntity()
                {
                   Brand = this.getNodeValue(authNode.SelectSingleNode("my:授权品牌", _nsmgr)),
                   Shen = this.getNodeValue(authNode.SelectSingleNode("my:省份", _nsmgr)),
                   VMode = this.getNodeValue(authNode.SelectSingleNode("my:经营方式", _nsmgr)),
                   DI = this.getNodeValue(authNode.SelectSingleNode("my:地级市", _nsmgr)),
                };
                if (authEntity.VMModeType == 1)
                    authEntity.LinShouAddr = this.getNodeValue(authNode.SelectSingleNode("my:零售地址", _nsmgr));
                
                authEntity.Kind = getNodeInnerText("my:乘用or商用");

                var IsAllXian = authNode.SelectSingleNode("my:全部or部分区县", _nsmgr).InnerText == "部分" ? false : true;

                if (IsAllXian) 
                    authEntity.Xian = "所有县";
                else 
                    authEntity.Xian = this.getXianList(authNode);

                authEntity.XFSCStr = this.getMarketList(authNode, obj);

                obj.AuthList.Add(authEntity);
            }


        }

        /// <summary>
        /// Sheet2 目标
        /// </summary>
        private void getPCSheet2(B2BEntity b2BEntity)
        {
            string rootPath = "//my:合同信息/my:目标-乘用/my:乘用目标/my:半钢胎目标";
            b2BEntity.crmSheet2PCList = new List<sheet2>();

            sheet2 sheetPC2 = new sheet2(b2BEntity);
            b2BEntity.crmSheet2PCList.Add(sheetPC2);
            sheetPC2.Brand = this.getNodeInnerText(rootPath + "/my:半钢-品牌");
            sheetPC2.S_Y1 = this.getNodeInnerText(rootPath + "/my:半-Q1-额");
            sheetPC2.S_T1 = this.getNodeInnerText(rootPath + "/my:半-Q1-量");
            sheetPC2.S_Y2 = this.getNodeInnerText(rootPath + "/my:半-Q2-额");
            sheetPC2.S_T2 = this.getNodeInnerText(rootPath + "/my:半-Q2-量");
            sheetPC2.S_Y3 = this.getNodeInnerText(rootPath + "/my:半-Q3-额");
            sheetPC2.S_T3 = this.getNodeInnerText(rootPath + "/my:半-Q3-量");
            sheetPC2.S_Y4 = this.getNodeInnerText(rootPath + "/my:半-Q4-额");
            sheetPC2.S_T4 = this.getNodeInnerText(rootPath + "/my:半-Q4-量");
            sheetPC2.S_Y = this.getNodeInnerText(rootPath + "/my:半-全年-额");
            sheetPC2.S_T = this.getNodeInnerText(rootPath + "/my:半-全年-量");

            sheetPC2 = new sheet2(b2BEntity);
            sheetPC2.Brand = this.getNodeInnerText(rootPath + "/my:轮辋18");
            sheetPC2.S_Y1 = this.getNodeInnerText(rootPath + "/my:半-Q1-额-18");
            sheetPC2.S_T1 = this.getNodeInnerText(rootPath + "/my:半-Q1-量-18");
            sheetPC2.S_Y2 = this.getNodeInnerText(rootPath + "/my:半-Q2-额-18");
            sheetPC2.S_T2 = this.getNodeInnerText(rootPath + "/my:半-Q2-量-18");
            sheetPC2.S_Y3 = this.getNodeInnerText(rootPath + "/my:半-Q3-额-18");
            sheetPC2.S_T3 = this.getNodeInnerText(rootPath + "/my:半-Q3-量-18");
            sheetPC2.S_Y4 = this.getNodeInnerText(rootPath + "/my:半-Q4-额-18");
            sheetPC2.S_T4 = this.getNodeInnerText(rootPath + "/my:半-Q4-量-18");
            sheetPC2.S_Y = this.getNodeInnerText(rootPath + "/my:半-全年-额-18");
            sheetPC2.S_T = this.getNodeInnerText(rootPath + "/my:半-全年-量-18");
            b2BEntity.crmSheet2PCList.Add(sheetPC2);

            sheetPC2 = new sheet2(b2BEntity);
            sheetPC2.Brand = this.getNodeInnerText(rootPath + "/my:轮辋1617");
            sheetPC2.S_Y1 = this.getNodeInnerText(rootPath + "/my:半-Q1-额-1617");
            sheetPC2.S_T1 = this.getNodeInnerText(rootPath + "/my:半-Q1-量-1617");
            sheetPC2.S_Y2 = this.getNodeInnerText(rootPath + "/my:半-Q2-额-1617");
            sheetPC2.S_T2 = this.getNodeInnerText(rootPath + "/my:半-Q2-量-1617");
            sheetPC2.S_Y3 = this.getNodeInnerText(rootPath + "/my:半-Q3-额-1617");
            sheetPC2.S_T3 = this.getNodeInnerText(rootPath + "/my:半-Q3-量-1617");
            sheetPC2.S_Y4 = this.getNodeInnerText(rootPath + "/my:半-Q4-额-1617");
            sheetPC2.S_T4 = this.getNodeInnerText(rootPath + "/my:半-Q4-量-1617");
            sheetPC2.S_Y = this.getNodeInnerText(rootPath + "/my:半-全年-额-1617");
            sheetPC2.S_T = this.getNodeInnerText(rootPath + "/my:半-全年-量-1617");
            b2BEntity.crmSheet2PCList.Add(sheetPC2);
        }

        //商用sheet2
        private void getTBSheet2(B2BEntity b2BEntity)
        {
            string rootPath = "//my:合同信息/my:目标-商用/my:商用目标/my:商用车胎目标";
            b2BEntity.crmSheet2TBList = new List<sheet2>();

            sheet2 sheetTB2 = new sheet2(b2BEntity);
            sheetTB2.Brand = this.getNodeInnerText(rootPath + "/my:商用-品牌");
            sheetTB2.S_Y1 = this.getNodeInnerText(rootPath + "/my:商用-Q1-额");
            sheetTB2.S_T1 = this.getNodeInnerText(rootPath + "/my:商用-Q1-量");
            sheetTB2.S_Y2 = this.getNodeInnerText(rootPath + "/my:商用-Q2-额");
            sheetTB2.S_T2 = this.getNodeInnerText(rootPath + "/my:商用-Q2-量");
            sheetTB2.S_Y3 = this.getNodeInnerText(rootPath + "/my:商用-Q3-额");
            sheetTB2.S_T3 = this.getNodeInnerText(rootPath + "/my:商用-Q3-量");
            sheetTB2.S_Y4 = this.getNodeInnerText(rootPath + "/my:商用-Q4-额");
            sheetTB2.S_T4 = this.getNodeInnerText(rootPath + "/my:商用-Q4-量");
            sheetTB2.S_Y = this.getNodeInnerText(rootPath + "/my:商用-全年-额");
            sheetTB2.S_T = this.getNodeInnerText(rootPath + "/my:商用-全年-量");
            //  sheetTB2.SCXIFEN
            sheetTB2.SCXIFEN = b2BEntity.XFSC_HuoYunStr;
            b2BEntity.crmSheet2TBList.Add(sheetTB2);

            var hasKeYun = this.getNodeInnerText(rootPath + "/my:客运市场");
            if(!string.IsNullOrEmpty(hasKeYun) && hasKeYun == "客运")
            {
                sheetTB2 = new sheet2(b2BEntity);
                sheetTB2.Brand = "佳通";// this.getNodeInnerText(rootPath + "/my:商用-品牌");
                sheetTB2.S_Y1 = this.getNodeInnerText(rootPath + "/my:客运-Q1-额");
                sheetTB2.S_T1 = this.getNodeInnerText(rootPath + "/my:客运-Q1-量");
                sheetTB2.S_Y2 = this.getNodeInnerText(rootPath + "/my:客运-Q2-额");
                sheetTB2.S_T2 = this.getNodeInnerText(rootPath + "/my:客运-Q2-量");
                sheetTB2.S_Y3 = this.getNodeInnerText(rootPath + "/my:客运-Q3-额");
                sheetTB2.S_T3 = this.getNodeInnerText(rootPath + "/my:客运-Q3-量");
                sheetTB2.S_Y4 = this.getNodeInnerText(rootPath + "/my:客运-Q4-额");
                sheetTB2.S_T4 = this.getNodeInnerText(rootPath + "/my:客运-Q4-量");
                sheetTB2.S_Y = this.getNodeInnerText(rootPath + "/my:客运-全年-额");
                sheetTB2.S_T = this.getNodeInnerText(rootPath + "/my:客运-全年-量");
                sheetTB2.SCXIFEN = this.getKeYunXiFen_TBSheet2();

                b2BEntity.crmSheet2TBList.Add(sheetTB2);
            }
        

        
        }
        
        //CRM 使用：部分区县List
        private string getXianList(XmlNode rootNode)
        {
            string resultList = "";
           var nodes = rootNode.SelectNodes("//my:县级/my:区县", _nsmgr);
           for(int i = 0; i < nodes.Count; i++)
            {
                var node = nodes[i];
                if(node!=null && !string.IsNullOrEmpty(node.InnerText))
                {
                    resultList += node.InnerText+"，";
                    //if (i != nodes.Count-1)
                    //resultList +=  "，";
                }
            }
            resultList = resultList.Remove(resultList.Length - 1);
            return resultList;
        }

        //细分市场
        private string getMarketList(XmlNode rootNode, B2BEntity b2BEntity)
        {
            string resultList = "";
            //  var nodes = _doc.SelectNodes("//my:合同信息/my:合同授权/my:授权区域/my:全钢细分市场/child::node()", _nsmgr);
            var nodes = rootNode.SelectNodes("my:全钢细分市场/child::node()", _nsmgr);

            for (int i = 0; i < nodes.Count; i++)
            {
                var node = nodes[i];
                if (node != null && !string.IsNullOrEmpty(node.InnerText))
                {
                    resultList += node.InnerText + ",";
                    var findResult = b2BEntity.XFSCList.Find(a => a.Contains(node.InnerText));
                    if (findResult == null)
                    {
                        b2BEntity.XFSCList.Add(node.InnerText);
                    }
                   // if (b2BEntity != null)
                   // {
                   ////     b2BEntity.XFSCList.Find(a=>a.Contains(a.))
                   //     b2BEntity.XFSCList.Add(node.InnerText);
                   // }
                        

                }
            }
            if(resultList.Length>0)
                resultList = resultList.Remove(resultList.Length - 1);
         
            return resultList;
        }
        
        /// <summary>
        /// 获取TB sheet2 中客运细分
        /// </summary>
        private string getKeYunXiFen_TBSheet2()
        {
            string resultList = "";
            var nodes = _doc.SelectSingleNode("//my:客运授权/my:客运细分显示", _nsmgr);
            resultList = nodes.InnerText;

            return resultList;
        }       

        public string getNodeValue(XmlNode node)
        {
            if (node != null) return node.InnerText;
            else return "";
        }
        public string getNodeInnerText(string xpath)
        {
            var node = _doc.SelectSingleNode(xpath, _nsmgr);
            return getNodeValue(node);
        }
         
        public void SaveToXmlFile()
        {
            string outputFileName = _RPACore.Configuration["setting:infoPathOutDir"] + "\\F.xml";

            string json = JsonConvert.SerializeObject(_B2BJsonReault);
            string b2bTemplate = _RPACore.CurrentDirectory + "\\B2BTemplate.xml";
            using (StreamReader sr = new StreamReader(b2bTemplate))
            {
                string xml = sr.ReadToEnd();
                xml = string.Format(xml, json);
                XmlDocument resultXml = new XmlDocument();
                resultXml.LoadXml(xml);
                resultXml.Save(outputFileName);
            }

            #region 不用xml对象生成
            //XmlDocument doc = new XmlDocument();

            // XmlTextWriter tr = new XmlTextWriter(outputFileName, Encoding.UTF8);

            //tr.WriteStartDocument();

            //tr.WriteStartElement("soap", "Envelope", "http://schemas.xmlsoap.org/soap/envelope/");
            //tr.WriteAttributeString("xmlns", "soap", null, "http://schemas.xmlsoap.org/soap/envelope/");
            //tr.WriteAttributeString("xmlns", "xsi", null, "http://www.w3.org/2001/XMLSchema-instance");
            //tr.WriteAttributeString("xmlns", "xsd", null, "http://www.w3.org/2001/XMLSchema");

            //tr.WriteStartElement("Body", "http://schemas.xmlsoap.org/soap/envelope/");

            //string json = JsonConvert.SerializeObject(_B2BJsonReault);

            //tr.WriteStartElement(null, "GetAccountInfoResponse", "http://tempuri.org/");
            #endregion
        }

        public void SaveToCRMExcel()
        {
            var dtName = DateTime.Now.ToString("yyyyMMddhhmm");
            string outputFileName = _RPACore.Configuration["setting:infoPathOutDir"] + $"\\CRM_{dtName}.xlsx";
           
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(outputFileName))) 
            {
                _SheetPC1 = package.Workbook.Worksheets.Add("TBIZ_AuthArea_c_2012(乘用)");
                genSheetPC1_header(_SheetPC1);
                _SheetPC2 = package.Workbook.Worksheets.Add("tbiz_saletarget_c_2012(乘用)");
                genSheetPC2_header(_SheetPC2);

                _SheetTB1 = package.Workbook.Worksheets.Add("tbiz_autharea_2012（商用）");
                genSheetTB1_header(_SheetTB1);
                _SheetTB2 = package.Workbook.Worksheets.Add("TBIZ_SaleTarget_2012（商用）");
                genSheetTB2_header(_SheetTB2);

                _SheetClientType = package.Workbook.Worksheets.Add("TPRT_Client_TYPE（全部）");
                genSheetClientType_header(_SheetClientType);

                int sheet1row = 2;
                int sheet2row= 2;
                int sheet3row = 2;
                int sheet4row = 2;

                foreach (var entity in _B2BJsonReault.Data)
                {
                    if (entity.Ifpc)
                    {
                        genSheetPC1(entity, _SheetPC1, ref sheet1row);
                     //   sheet1row++;
                        genSheetPC2(entity, _SheetPC2, ref sheet2row);
                   //    sheet2row++;
                    }
                    else
                    {
                        genSheetTB1(entity, _SheetTB1, ref sheet3row);
                     //   sheet3row++;
                        genSheetTB2(entity, _SheetTB2, ref sheet4row);
                    //    sheet4row++;
                    }
                }

                genSheecClientType();

                package.Save();
            }
        }
        
        //临时使用
        public void TempJsonExcel()
        {
            var dtName = DateTime.Now.ToString("yyyyMMddhhmm");
            string outputFileName = _RPACore.Configuration["setting:infoPathOutDir"] + $"\\TempB2B_{dtName}.xlsx";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(outputFileName)))
            {
                int row = 1, col = 1;
                var b2bSheet = package.Workbook.Worksheets.Add("B2B");
                b2bSheet.Cells[row, col++].Value = "Addresstype";
                b2bSheet.Cells[row, col++].Value = "AccountNum";
                b2bSheet.Cells[row, col++].Value = "CustomerType";
                b2bSheet.Cells[row, col++].Value = "Name";
                b2bSheet.Cells[row, col++].Value = "Region";
                b2bSheet.Cells[row, col++].Value = "County";
                b2bSheet.Cells[row, col++].Value = "Province";
                b2bSheet.Cells[row, col++].Value = "City";
                b2bSheet.Cells[row, col++].Value = "AddressNum";
                b2bSheet.Cells[row, col++].Value = "Addressline";
                b2bSheet.Cells[row, col++].Value = "BusinessContact";
                b2bSheet.Cells[row, col++].Value = "PhoneNumber";
                b2bSheet.Cells[row, col++].Value = "CustomerEmail";
                b2bSheet.Cells[row, col++].Value = "Ifpc";
                b2bSheet.Cells[row, col++].Value = "Pcdistrict";
                b2bSheet.Cells[row, col++].Value = "Pcbranch";
                b2bSheet.Cells[row, col++].Value = "PcSalesContact";
                b2bSheet.Cells[row, col++].Value = "PcSalesEmail";
                b2bSheet.Cells[row, col++].Value = "PcSalesSupervisor";
                b2bSheet.Cells[row, col++].Value = "PcSalesSupervisorEmail";
                b2bSheet.Cells[row, col++].Value = "Iftb";
                b2bSheet.Cells[row, col++].Value = "Tbdistrict";
                b2bSheet.Cells[row, col++].Value = "Tbbranch";
                b2bSheet.Cells[row, col++].Value = "TbSalesContact";
                b2bSheet.Cells[row, col++].Value = "TbSalesEmail";
                b2bSheet.Cells[row, col++].Value = "TbSalesSupervisor";
                b2bSheet.Cells[row, col++].Value = "TbSalesSupervisorEmail";
                b2bSheet.Cells[row, col++].Value = "Ifbias";
                b2bSheet.Cells[row, col++].Value = "Biasdistrict";
                b2bSheet.Cells[row, col++].Value = "Biasbranch";
                b2bSheet.Cells[row, col++].Value = "BiasSalesContact";
                b2bSheet.Cells[row, col++].Value = "BiasSalesEmail";
                b2bSheet.Cells[row, col++].Value = "BiasSalesSupervisor";
                b2bSheet.Cells[row, col++].Value = "BiasSalesSupervisorEmail";
                b2bSheet.Cells[row, col++].Value = "ICSContact";
                b2bSheet.Cells[row, col++].Value = "ICSEmail";
                b2bSheet.Cells[row, col++].Value = "ICSSupervisor";
                b2bSheet.Cells[row, col++].Value = "ICSSupervisorEmail";
                b2bSheet.Cells[row, col++].Value = "Status";
                b2bSheet.Cells[row, col++].Value = "Createdon";
                b2bSheet.Cells[row, col++].Value = "CustomerServiceDistribution";
                b2bSheet.Cells[row, col++].Value = "K2District";

                row++;col = 1;
                foreach(var rd in _B2BJsonReault.Data)
                {
                    AddressEntity addr = rd.AddressList[0];
                    if (addr == null) addr = new AddressEntity();
                    b2bSheet.Cells[row, col++].Value = addr.Addresstype;// "Addresstype";
                    b2bSheet.Cells[row, col++].Value = rd.AccountNum;
                    b2bSheet.Cells[row, col++].Value = rd.CustomerType; 
                    b2bSheet.Cells[row, col++].Value = rd.Name; 
                    b2bSheet.Cells[row, col++].Value = addr.Region;
                    b2bSheet.Cells[row, col++].Value = addr.County;
                    b2bSheet.Cells[row, col++].Value = addr.Province;
                    b2bSheet.Cells[row, col++].Value = addr.City;
                    b2bSheet.Cells[row, col++].Value = addr.AddressNum;
                    b2bSheet.Cells[row, col++].Value = addr.Addressline;
                    b2bSheet.Cells[row, col++].Value = rd.BusinessContact;
                    b2bSheet.Cells[row, col++].Value = rd.PhoneNumber;
                    b2bSheet.Cells[row, col++].Value = rd.CustomerEmail;
                    b2bSheet.Cells[row, col++].Value = rd.Ifpc?"True":"False";
                    b2bSheet.Cells[row, col++].Value = rd.Pcdistrict;
                    b2bSheet.Cells[row, col++].Value = rd.Pcbranch;
                    b2bSheet.Cells[row, col++].Value = rd.PcSalesContact;
                    b2bSheet.Cells[row, col++].Value = rd.PcSalesEmail;
                    b2bSheet.Cells[row, col++].Value = rd.PcSalesSupervisor;
                    b2bSheet.Cells[row, col++].Value = rd.PcSalesSupervisorEmail;
                    b2bSheet.Cells[row, col++].Value = rd.Iftb?"True":"False";
                    b2bSheet.Cells[row, col++].Value = rd.Tbdistrict;
                    b2bSheet.Cells[row, col++].Value = rd.Tbbranch;
                    b2bSheet.Cells[row, col++].Value = rd.TbSalesContact;
                    b2bSheet.Cells[row, col++].Value = rd.TbSalesEmail;
                    b2bSheet.Cells[row, col++].Value = rd.TbSalesSupervisor;
                    b2bSheet.Cells[row, col++].Value = rd.TbSalesSupervisorEmail;
                    b2bSheet.Cells[row, col++].Value = "False";
                    b2bSheet.Cells[row, col++].Value = rd.Biasdistrict;
                    b2bSheet.Cells[row, col++].Value = rd.Biasbranch;
                    b2bSheet.Cells[row, col++].Value = rd.BiasSalesContact;
                    b2bSheet.Cells[row, col++].Value = rd.BiasSalesEmail;
                    b2bSheet.Cells[row, col++].Value = rd.BiasSalesSupervisor;
                    b2bSheet.Cells[row, col++].Value = rd.BiasSalesSupervisorEmail;
                    b2bSheet.Cells[row, col++].Value = rd.ICSContact;
                    b2bSheet.Cells[row, col++].Value = rd.ICSEmail;
                    b2bSheet.Cells[row, col++].Value = rd.ICSSupervisor;
                    b2bSheet.Cells[row, col++].Value = rd.ICSSupervisorEmail;
                    b2bSheet.Cells[row, col++].Value = rd.Status;
                    b2bSheet.Cells[row, col++].Value = rd.Createdon;
                    b2bSheet.Cells[row, col++].Value = rd.CustomerServiceDistribution;
                    b2bSheet.Cells[row, col++].Value = rd.K2District;
                    row++; col = 1;
                }

                row = 1;col = 1;
                var QGTSheet = package.Workbook.Worksheets.Add("全钢胎");
                QGTSheet.Cells[row, col++].Value = "ERP编码";
                QGTSheet.Cells[row, col++].Value = "授权品牌";
                QGTSheet.Cells[row, col++].Value = "细分市场";
                row++; col = 1;
                foreach (var rd in _B2BJsonReault.Data)
                {
                    if (rd.Iftb)
                    {
                        foreach(string huoyunXifen in rd.XFSCList)
                        {

                            QGTSheet.Cells[row, col++].Value = rd.AccountNum;
                            QGTSheet.Cells[row, col++].Value = rd.applyBrand;
                            QGTSheet.Cells[row, col++].Value = huoyunXifen;
                            row++; col = 1;
                        }
                        foreach (string keyunXifen in rd.KeYunXiFenList)
                        {
                            QGTSheet.Cells[row, col++].Value = rd.AccountNum;
                            QGTSheet.Cells[row, col++].Value = rd.applyBrand;
                            QGTSheet.Cells[row, col++].Value = keyunXifen;
                            row++; col = 1;
                        }
                    }
                 
                }
                 row = 1; col = 1;
                var CYCTSheet = package.Workbook.Worksheets.Add("乘用车胎");
                CYCTSheet.Cells[row, col++].Value = "ERP编码";
                CYCTSheet.Cells[row, col++].Value = "授权品牌";
                CYCTSheet.Cells[row, col++].Value = "区域";
                row++; col = 1;
                foreach (var rd in _B2BJsonReault.Data)
                {
                    if (rd.Ifpc)
                    {
                        CYCTSheet.Cells[row, col++].Value = rd.AccountNum;
                        CYCTSheet.Cells[row, col++].Value = rd.applyBrand;
                        CYCTSheet.Cells[row, col++].Value = rd.Ifpc?rd.Pcdistrict:rd.Tbbranch;
                        row++; col = 1;
                    }
                }
                    package.Save();
            }
        }
       
        private void genSheetPC1_header(ExcelWorksheet sheet)
        {
            sheet.Cells[1, 1].Value = "ID";
            sheet.Cells[1, 2].Value = "CustomerCode";
            sheet.Cells[1, 3].Value = "BRAND";
            sheet.Cells[1, 4].Value = "KIND";
            sheet.Cells[1, 5].Value = "SHEN";
            sheet.Cells[1, 6].Value = "VMODE";
            sheet.Cells[1, 7].Value = "ADDRESS";
            sheet.Cells[1, 8].Value = "DI";
            sheet.Cells[1, 9].Value = "XIAN";
            sheet.Cells[1, 10].Value = "LastUpdateTime";
         
        }
        
        //乘用1
        private void genSheetPC1(B2BEntity entity,ExcelWorksheet sheet,ref int row)
        {
            foreach (var auth in entity.AuthList)
            {
                sheet.Cells[row, 1].Value = entity.crmId;
                sheet.Cells[row, 2].Value = entity.AccountNum;
                sheet.Cells[row, 3].Value = auth.Brand;
                sheet.Cells[row, 4].Value = auth.Kind;
                sheet.Cells[row, 5].Value = auth.Shen;
                sheet.Cells[row, 6].Value = auth.VMode;
                sheet.Cells[row, 7].Value = auth.LinShouAddr;
                sheet.Cells[row, 8].Value = auth.DI;
                sheet.Cells[row, 9].Value = auth.Xian;
                sheet.Cells[row, 10].Value = entity.Createdon;
            }
            //    sheet.Cells[row, 1].Value = entity.crmId;
            //sheet.Cells[row, 2].Value = entity.AccountNum;
            //sheet.Cells[row, 3].Value = entity.crmBrand;
            //sheet.Cells[row, 4].Value = entity.crmKind;
            //sheet.Cells[row, 5].Value = entity.crmShen;
            //sheet.Cells[row, 6].Value = entity.crmVMode;
            //sheet.Cells[row, 7].Value = entity.crmLinShouAddr;
            //sheet.Cells[row, 8].Value = entity.crmDI;
            //sheet.Cells[row, 9].Value = entity.crmXian;
            //sheet.Cells[row, 10].Value = entity.Createdon;
        }
           
        private void genSheetPC2_header(ExcelWorksheet sheet)
        {
            sheet.Cells[1, 1].Value = "ID";
            sheet.Cells[1, 2].Value = "KIND";
            sheet.Cells[1, 3].Value = "BRAND";
            sheet.Cells[1, 4].Value = "S_Y1";
            sheet.Cells[1, 5].Value = "S_T1";
            sheet.Cells[1, 6].Value = "S_Y2";
            sheet.Cells[1, 7].Value = "S_T2";
            sheet.Cells[1, 8].Value = "S_Y3";
            sheet.Cells[1, 9].Value = "S_T3";
            sheet.Cells[1, 10].Value = "S_Y4";
            sheet.Cells[1, 11].Value = "S_T4";
            sheet.Cells[1, 12].Value = "S_Y";
            sheet.Cells[1, 13].Value = "S_T";
        }
        
        //乘用2
        private void genSheetPC2(B2BEntity entity, ExcelWorksheet sheet,ref int row)
        {

            string emptyValue = "（空白）";
            foreach (var crmS2Data in entity.crmSheet2PCList)
            {
                sheet.Cells[row, 1].Value = entity.crmId;
                sheet.Cells[row, 2].Value = entity.crmKind;
                sheet.Cells[row, 3].Value = crmS2Data.Brand;
                if (string.IsNullOrEmpty(crmS2Data.S_Y1))  sheet.Cells[row, 4].Value = emptyValue;
                else sheet.Cells[row, 4].Value =  Convert.ToInt32(crmS2Data.S_Y1);

                if (string.IsNullOrEmpty(crmS2Data.S_T1)) sheet.Cells[row, 5].Value = emptyValue;
                else sheet.Cells[row, 5].Value = Convert.ToInt32(crmS2Data.S_T1);

                if (string.IsNullOrEmpty(crmS2Data.S_Y2)) sheet.Cells[row, 6].Value = emptyValue;
                else sheet.Cells[row, 6].Value = Convert.ToInt32(crmS2Data.S_Y2);

                if (string.IsNullOrEmpty(crmS2Data.S_T2)) sheet.Cells[row, 7].Value = emptyValue;
                else sheet.Cells[row, 7].Value = Convert.ToInt32(crmS2Data.S_T2);

                if (string.IsNullOrEmpty(crmS2Data.S_Y3)) sheet.Cells[row, 8].Value = emptyValue;
                else sheet.Cells[row, 8].Value = Convert.ToInt32(crmS2Data.S_Y3);

                if (string.IsNullOrEmpty(crmS2Data.S_T3)) sheet.Cells[row, 9].Value = emptyValue;
                else sheet.Cells[row, 9].Value = Convert.ToInt32(crmS2Data.S_T3);

                if (string.IsNullOrEmpty(crmS2Data.S_Y4)) sheet.Cells[row, 10].Value = emptyValue;
                else sheet.Cells[row, 10].Value = Convert.ToInt32(crmS2Data.S_Y4);

                if (string.IsNullOrEmpty(crmS2Data.S_T4)) sheet.Cells[row, 11].Value = emptyValue;
                else sheet.Cells[row, 11].Value = Convert.ToInt32(crmS2Data.S_T4);

                if (string.IsNullOrEmpty(crmS2Data.S_Y)) sheet.Cells[row, 12].Value = emptyValue;
                else sheet.Cells[row, 12].Value = Convert.ToInt32(crmS2Data.S_Y);

                if (string.IsNullOrEmpty(crmS2Data.S_T)) sheet.Cells[row, 13].Value = emptyValue;
                else sheet.Cells[row, 13].Value = Convert.ToInt32(crmS2Data.S_T);

                //sheet.Cells[row, 5].Value = crmS2Data.S_T1;
                //sheet.Cells[row, 6].Value = crmS2Data.S_Y2;
                //sheet.Cells[row, 7].Value = crmS2Data.S_T2;
                //sheet.Cells[row, 8].Value = crmS2Data.S_Y3;
                //sheet.Cells[row, 9].Value = crmS2Data.S_T3;
                //sheet.Cells[row, 10].Value = crmS2Data.S_Y4;
                //sheet.Cells[row, 11].Value = crmS2Data.S_T4;
                //sheet.Cells[row, 12].Value = crmS2Data.S_Y;
                //sheet.Cells[row, 13].Value = crmS2Data.S_T;
                row++;
            }
        }

        private void genSheetTB1_header(ExcelWorksheet sheet)
        {
            sheet.Cells[1, 1].Value = "ID";
            sheet.Cells[1, 2].Value = "CustomerCode";
            sheet.Cells[1, 3].Value = "PINZ";
            sheet.Cells[1, 4].Value = "PINP";
            sheet.Cells[1, 5].Value = "SCXFEN";
            sheet.Cells[1, 6].Value = "TIRETYPE";
            sheet.Cells[1, 7].Value = "SHEN";
            sheet.Cells[1, 8].Value = "VMODE";
            sheet.Cells[1, 9].Value = "ADDRESS";
            sheet.Cells[1, 10].Value = "DI";
            sheet.Cells[1, 11].Value = "XIAN";
            sheet.Cells[1, 12].Value = "LastUpdateTime";

        }

        //商用1
        private void genSheetTB1(B2BEntity entity, ExcelWorksheet sheet, ref int row)
        {
            foreach(var auth in entity.AuthList)
            {
                sheet.Cells[row, 1].Value = entity.crmId;
                sheet.Cells[row, 2].Value = entity.AccountNum;
                sheet.Cells[row, 3].Value = auth.Kind;
                sheet.Cells[row, 4].Value = auth.Brand;
                sheet.Cells[row, 5].Value = auth.XFSCStr;
                //   sheet.Cells[row, 5].Value = auth.SCXFEN;
                sheet.Cells[row, 6].Value = auth.Series;
                sheet.Cells[row, 7].Value = auth.Shen;
                sheet.Cells[row, 8].Value = auth.VMode;
                sheet.Cells[row, 9].Value = auth.LinShouAddr;
                sheet.Cells[row, 10].Value = auth.DI;
                sheet.Cells[row, 11].Value = auth.Xian;
                sheet.Cells[row, 12].Value = entity.Createdon;
                row++;
            }
        //    sheet.Cells[row, 3].Value = entity.crmKind;
        //    sheet.Cells[row, 4].Value = entity.crmBrand;
        ////    sheet.Cells[row, 5].Value = entity.crmSCXFEN;
        //    sheet.Cells[row, 6].Value = entity.crmSeries;
        //    sheet.Cells[row, 7].Value = entity.crmShen;
        //    sheet.Cells[row, 8].Value = entity.crmVMode;
        //    sheet.Cells[row, 9].Value = entity.crmLinShouAddr;
        //    sheet.Cells[row, 10].Value = entity.crmDI;
        //    sheet.Cells[row, 11].Value = entity.crmXian;

        }
        
        private void genSheetTB2_header(ExcelWorksheet sheet)
        {
            sheet.Cells[1, 1].Value = "ID";
            sheet.Cells[1, 2].Value = "BINZHONG";
            sheet.Cells[1, 3].Value = "PINPAI";
            sheet.Cells[1, 4].Value = "SCXIFEN";
            sheet.Cells[1, 5].Value = "S_Y1";
            sheet.Cells[1, 6].Value = "S_T1";
            sheet.Cells[1, 7].Value = "S_Y2";
            sheet.Cells[1, 8].Value = "S_T2";
            sheet.Cells[1, 9].Value = "S_Y3";
            sheet.Cells[1, 10].Value = "S_T3";
            sheet.Cells[1, 11].Value = "S_Y4";
            sheet.Cells[1, 12].Value = "S_T4";
            sheet.Cells[1, 13].Value = "S_Y";
            sheet.Cells[1, 14].Value = "S_T";

        }

        //商用2
        private void genSheetTB2(B2BEntity entity, ExcelWorksheet sheet, ref int row)
        {
            string emptyValue = "（空白）";
            foreach (var crmS2Data in entity.crmSheet2TBList)
            {

                sheet.Cells[row, 1].Value = entity.crmId;
                sheet.Cells[row, 2].Value = entity.crmKind;
                sheet.Cells[row, 3].Value = crmS2Data.Brand;
                sheet.Cells[row, 4].Value = crmS2Data.SCXIFEN;

                if (string.IsNullOrEmpty(crmS2Data.S_Y1)) sheet.Cells[row, 5].Value = emptyValue;
                else sheet.Cells[row, 5].Value = Convert.ToInt32(crmS2Data.S_Y1);

                if (string.IsNullOrEmpty(crmS2Data.S_T1)) sheet.Cells[row, 6].Value = emptyValue;
                else sheet.Cells[row, 6].Value = Convert.ToInt32(crmS2Data.S_T1);

                if (string.IsNullOrEmpty(crmS2Data.S_Y2)) sheet.Cells[row, 7].Value = emptyValue;
                else sheet.Cells[row, 7].Value = Convert.ToInt32(crmS2Data.S_Y2);

                if (string.IsNullOrEmpty(crmS2Data.S_T2)) sheet.Cells[row, 8].Value = emptyValue;
                else sheet.Cells[row, 8].Value = Convert.ToInt32(crmS2Data.S_T2);

                if (string.IsNullOrEmpty(crmS2Data.S_Y3)) sheet.Cells[row, 9].Value = emptyValue;
                else sheet.Cells[row, 9].Value = Convert.ToInt32(crmS2Data.S_Y3);

                if (string.IsNullOrEmpty(crmS2Data.S_T3)) sheet.Cells[row, 10].Value = emptyValue;
                else sheet.Cells[row, 10].Value = Convert.ToInt32(crmS2Data.S_T3);

                if (string.IsNullOrEmpty(crmS2Data.S_Y4)) sheet.Cells[row, 11].Value = emptyValue;
                else sheet.Cells[row, 11].Value = Convert.ToInt32(crmS2Data.S_Y4);

                if (string.IsNullOrEmpty(crmS2Data.S_T4)) sheet.Cells[row, 12].Value = emptyValue;
                else sheet.Cells[row, 12].Value = Convert.ToInt32(crmS2Data.S_T4);

                if (string.IsNullOrEmpty(crmS2Data.S_Y)) sheet.Cells[row, 13].Value = emptyValue;
                else sheet.Cells[row, 13].Value = Convert.ToInt32(crmS2Data.S_Y);

                if (string.IsNullOrEmpty(crmS2Data.S_T)) sheet.Cells[row, 14].Value = emptyValue;
                else sheet.Cells[row, 14].Value = Convert.ToInt32(crmS2Data.S_T);

                //sheet.Cells[row, 5].Value = crmS2Data.S_T1;
                //sheet.Cells[row, 6].Value = crmS2Data.S_Y2;
                //sheet.Cells[row, 7].Value = crmS2Data.S_T2;
                //sheet.Cells[row, 8].Value = crmS2Data.S_Y3;
                //sheet.Cells[row, 9].Value = crmS2Data.S_T3;
                //sheet.Cells[row, 10].Value = crmS2Data.S_Y4;
                //sheet.Cells[row, 11].Value = crmS2Data.S_T4;
                //sheet.Cells[row, 12].Value = crmS2Data.S_Y;
                //sheet.Cells[row, 13].Value = crmS2Data.S_T;
                row++;
            }
        }

        private void genSheetClientType_header(ExcelWorksheet sheet)
        {
            sheet.Cells[1, 1].Value = "CUST_NUMBER";
            sheet.Cells[1, 2].Value = "IS_PCR";
            sheet.Cells[1, 3].Value = "IS_TBR";
            sheet.Cells[1, 4].Value = "IS_BIAS";
            sheet.Cells[1, 5].Value = "LAST_UPDATE";
        }
    
        private void genSheecClientType()
        {
            var allData = _B2BJsonReault.Data;
            int row = 2;
            foreach(var rowData in allData)
            {
                _SheetClientType.Cells[row, 1].Value = rowData.AccountNum;
                _SheetClientType.Cells[row, 2].Value = rowData.Ifpc ? "Y" : "N";
                _SheetClientType.Cells[row, 3].Value = rowData.Ifpc ? "N" : "Y";
                _SheetClientType.Cells[row, 4].Value = "N";
                _SheetClientType.Cells[row, 5].Value = rowData.Createdon;
                row++;
            }
        }
    }
}
