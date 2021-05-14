using Microsoft.Extensions.Configuration;
using OperateInfoPath.CRMModels;
using RPA.Core;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OperateInfoPath
{
    public class CRMImport
    {
        private SqlSugarClient _Db;
        private List<B2BEntity> _dataList;
   
        public CRMImport(List<B2BEntity> dataList, SqlSugarClient db)
        {
            _dataList = dataList;
            _Db = db;
        }

        public void CreateModels()
        {
            _Db.DbFirst.CreateClassFile(@"C:\Project\UIPath\SD030\ProgromModule\OperateInfoPath\CRMModels", "OperateInfoPath.CRMModels");
        }

        public  void InsertToDataBase()
        {
            try
            {
                var result = _Db.UseTran(() =>
                {
                    int n;
                    foreach (var data in _dataList)
                    {
                        foreach (var pc1 in data.dbPC1)
                        {
                             n = _Db.Insertable(pc1).ExecuteCommand();
                        }
                        foreach (var pc2 in data.dbPC2)
                        {
                            n = _Db.Insertable(pc2).ExecuteCommand();
                        }
                        foreach (var tb1 in data.dbTB1)
                        {
                            n = _Db.Insertable(tb1).ExecuteCommand();
                        }
                        foreach (var tb2 in data.dbTB2)
                        {
                            n = _Db.Insertable(tb2).ExecuteCommand();
                        }
                        foreach (var summery in data.dbSummery)
                        {
                            n = _Db.Insertable(summery).ExecuteCommand();
                        }
                    }
                });
                if (!result.IsSuccess)
                {
                    NLogUtil.cc_ErrorTxt("CRM Import InsertToDataBase: " + result.ErrorException);
                }
               
            }
            catch(Exception ex)
            {
                NLogUtil.cc_ErrorTxt("CRM Import InsertToDataBase: "+ex.Message);
            }
          

        }

    }
}
