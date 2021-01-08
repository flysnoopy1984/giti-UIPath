using CsvHelper.Configuration.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace CSVAnalysis
{
    public class SC030CSVEntity
    {
        [Name("月描述")]
        public string a1 { get; set; }

        [Name("乘用销售区域")]
        public string a2 { get; set; }

        [Name("乘用销售分部")]
        public string a3 { get; set; }

        [Name("客户编码")]
        public string a4 { get; set; }

        [Name("客户名称")]
        public string a5 { get; set; }
        [Name("产品大类")]
        public string a6 { get; set; }
        [Name("花纹")]
        public string a7 { get; set; }
        [Name("轮辋（乘用）")]
        public string a8 { get; set; }

        [Name("车型")]
        public string a9 { get; set; }

        [Name("规格")]
        public string a10 { get; set; }

        [Name("品牌（中文）")]
        public string a11 { get; set; }

        [Name("销售订单数量")]
        public string a12 { get; set; }
    }
}
