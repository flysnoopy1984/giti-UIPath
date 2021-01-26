using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace OperateInfoPath
{
    //区域，分部，业务员，业务员邮箱
    public class dbssInfo
    {
        public string district { get; set; }

        public string branch { get; set; }

        public string SalesContact { get; set; }

        public string SalesEmail { get; set; }

        public string SalesSupervisor { get; set; }

        public string SalesSupervisorEmail { get; set; }
    }

    public class AddressEntity
    {
        public string Addresstype { get; set; }

        public string AddressNum { get; set; }

        public string Addressline { get; set; }

        public string Region { get; set; }

        public string County { get; set; }

        public string Province { get; set; }

        public string City { get; set; }
        public string Mainaddr { get; set; }

        public string Zipcode { get; set; }

    }
   
    public class sheet2
    {
        public sheet2(B2BEntity obj)
        {
            this.Id = obj.crmId;
            this.Kind = obj.crmKind;
      
        }
        public string Id { get; set; }
        public string Kind { get; set; }
        public string Brand { get; set; }

        public string S_Y1 { get; set; }
        public string S_T1 { get; set; }
        public string S_Y2 { get; set; }
        public string S_T2 { get; set; }
        public string S_Y3 { get; set; }
        public string S_T3 { get; set; }
        public string S_Y4 { get; set; }
        public string S_T4 { get; set; }
        public string S_Y { get; set; }
        public string S_T { get; set; }

        //细分市场
        public string SCXIFEN { get; set; }
       

    }
    public class B2BEntity
    {
        private List<AddressEntity> _AddressList;
        public List<AddressEntity> AddressList {
            get
            {
                if (_AddressList == null) _AddressList = new List<AddressEntity>();
                return _AddressList;
            }
        }

        //   public string mmDateTime { get; set; }

        public string AccountNum { get; set; }

        public string CustomerType { get; set; }

        public string Name { get; set; }

        public string BusinessContact { get; set; }

        public string PhoneNumber { get; set; }

        public string CustomerEmail { get; set; }

        public bool Ifpc { get; set; }

        public string Pcdistrict { get; set; }

        public string Pcbranch { get; set; }

        public string PcSalesContact { get; set; }

        public string PcSalesEmail { get; set; }

        public string PcSalesSupervisor { get; set; }

        public string PcSalesSupervisorEmail { get; set; }

        public bool Iftb { get; set; }

        public string Tbdistrict { get; set; }

        public string Tbbranch { get; set; }

        public string TbSalesContact { get; set; }

        public string TbSalesEmail { get; set; }

        public string TbSalesSupervisor { get; set; }

        public string TbSalesSupervisorEmail { get; set; }

        public bool Ifbias { get; set; }

        public string Biasdistrict { get; set; }

        public string Biasbranch { get; set; }

        public string BiasSalesContact { get; set; }

        public string BiasSalesEmail { get; set; }

        public string BiasSalesSupervisor { get; set; }

        public string BiasSalesSupervisorEmail { get; set; }

        public string ICSContact { get; set; }

        public string ICSEmail { get; set; }

        public string ICSSupervisor { get; set; }

        public string ICSSupervisorEmail { get; set; }

        public string Status { get; set; }

        public string Createdon { get; set; }

        public string CustomerServiceDistribution { get; set; }

        public string K2District { get; set; }

        /* orig xml B*/
        public string TradeTerm { get; set; }

        public string StandardPaymentTerm { get; set; }

        /* orig xml E*/

        #region CRM Excel
        [JsonIgnore]
        public string crmBrand { get; set; }

        [JsonIgnore]
        public string crmKind { get; set; }

        [JsonIgnore]
        public string crmShen { get; set; }

        /// <summary>
        /// 经营方式
        /// </summary>
        private string _crmVMode;

        [JsonIgnore]
        public string crmVMode
        {
            get { return _crmVMode; }
            set
            {
                _crmVMode = value;
                if (value == "零售")
                    crmJYFangShi = 1;
                else
                    crmJYFangShi = 0;
            }
        }

        /// <summary>
        /// 经营方式 0 批发 1 零售 
        /// </summary>
        [JsonIgnore]
        public int crmJYFangShi { get; set; }
        /// <summary>
        /// 零售地址
        /// </summary>
        [JsonIgnore]
        public  string crmLinShouAddr { get; set; }

        [JsonIgnore]
        public string crmDI { get; set; }

        [JsonIgnore]
        public string crmXian { get; set; }

        /// <summary>
        /// 自定义RelationId
        /// </summary>
        [JsonIgnore]
        public string crmId { get; set; }

        [JsonIgnore]
        public List<sheet2> crmSheet2PCList { get; set; }

        [JsonIgnore]
        public string crmSCXFEN { get; set; }

        //货运细分市场
        [JsonIgnore]
        public List<string> XFSCList { get; set; } = new List<string>();

        //客运细分市场
        public List<string> KeYunXiFenList { get; set; } = new List<string>();

        [JsonIgnore]
        public string crmSeries { get; set; }

        [JsonIgnore]
        public List<sheet2> crmSheet2TBList { get; set; } = new List<sheet2>();

        [JsonIgnore]
        public string applyBrand { get; set; }

        //是否客运
        [JsonIgnore]
        public bool hasKeYun { get; set; }

        #endregion


        public void InitPCInfo(dbssInfo info,ccJsonEntity ccJsonEntity)
        {
            if (ccJsonEntity != null)
            {
                ICSContact = ccJsonEntity.PCICSContact;
                CustomerEmail = ccJsonEntity.PCICSEmail;
                CustomerServiceDistribution = ccJsonEntity.PCDistrict;
            }
           
            Pcbranch = info.branch;
            Pcdistrict = info.district;
            //PcSalesContact = info.SalesContact;
            //PcSalesEmail = info.SalesEmail;
            //PcSalesSupervisor = ccJsonEntity.PCSalesSupervisor;
            //PcSalesSupervisorEmail = ccJsonEntity.PCICSSupervisorEmail;


        }

        public void InitTbInfo(dbssInfo info, ccJsonEntity ccJsonEntity)
        {
            if (ccJsonEntity!=null)
            {
                ICSContact = ccJsonEntity.TBICSContact;
                CustomerEmail = ccJsonEntity.TBICSEmail;
                CustomerServiceDistribution = ccJsonEntity.TBDistrict;
            }
          
            Tbbranch = info.branch;
            Tbdistrict = info.district;
            //TbSalesContact = info.SalesContact;
            //TbSalesEmail = info.SalesEmail;
            //TbSalesSupervisor = ccJsonEntity.TBSalesSupervisor;
            //TbSalesSupervisorEmail = ccJsonEntity.TBICSSupervisorEmail;



        }
    }

    public class B2BJsonReault
    {
        public string ErrorCode { get; set; }

        public string ErrorMsg { get; set; }

        public string Status { get; set; } = "success";

        private List<B2BEntity> _Data;
        public List<B2BEntity> Data {
            get
            {
                if (_Data == null) _Data = new List<B2BEntity>();
                return _Data;
            }
            set { _Data = value; }
        }
    }
}
