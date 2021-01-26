using System;
using System.Collections.Generic;
using System.Text;

namespace OperateInfoPath
{
    public class ccJsonList
    {
        public List<ccJsonEntity> List { get; set; }
    }
    public class ccJsonEntity
    {
        public string PCBranch { get; set; }
        public string PCSalesSupervisor { get; set; }
        public string PCICSContact { get; set; }
        public string PCDistrict { get; set; }
        public string PCICSEmail { get; set; }
        public string PCICSSupervisor { get; set; }
        public string PCICSSupervisorEmail { get; set; }
        public string TBBranch { get; set; }
        public string TBSalesSupervisor { get; set; }
        public string TBICSContact { get; set; }
        public string TBDistrict { get; set; }
        public string TBICSEmail { get; set; }
        public string TBICSSupervisor { get; set; }
        public string TBICSSupervisorEmail { get; set; }
    }

}
