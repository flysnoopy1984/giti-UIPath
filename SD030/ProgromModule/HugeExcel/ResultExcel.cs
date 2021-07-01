using OfficeOpenXml;
using RPA.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace HugeExcel
{
    public class ResultExcel
    {
        public static RPACore _RPACore = RPACore.getInstance();

        private string _FilePath;
        private DirectoryInfo _fileDir;

        public string DirPath
        {
            get
            {
                return _fileDir.FullName;
            }
        }

        public string FilePath
        {
            get { return _FilePath; }
        }

        public ResultExcel()
        {
            var dirPath = _RPACore.Configuration["HugeExcel:resultDir"];
            _fileDir = new DirectoryInfo(dirPath);
        }
        public void DeleteDirFIles()
        {
            var files = _fileDir.GetFiles();
            foreach(var file in files)
            {
                file.Delete();
            }
        }
        public void InitFilePath()
        {
            var files = _fileDir.GetFiles();
            if (files.Length > 0)
            {
                _FilePath = files[0].FullName;
            }
        }
        public bool ExistFilePath()
        {
            var files = _fileDir.GetFiles();
            return files.Length > 0;
        }

        //public void RunStepTwo()
        //{
        //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        //    using (ExcelPackage package = new ExcelPackage(new FileInfo(_FilePath)))
        //    {
        //        var sheet = package.Workbook.Worksheets["图-全钢胎"];

        //        var sheet2 = package.Workbook.Worksheets["图-半钢胎"];
        //        //    sheet.Cells[1, 2].Value = "有效";
        //        //sheet.Cells[2, 2].Value = "内销配套";
        //        //sheet.Cells[3, 2].Value = "101半钢外胎";

        //        //sheet.Cells[24, 2].Value = "有效";
        //        //sheet.Cells[25, 2].Value = "内销配套";
        //        //sheet.Cells[26, 2].Value = "101半钢外胎";

        //        //sheet = package.Workbook.Worksheets["图-全钢胎"];
        //        //sheet.Cells[1, 2].Value = "有效";
        //        //sheet.Cells[2, 2].Value = "内销配套";
        //        //sheet.Cells[3, 2].Value = "106全钢外胎";

        //        //sheet.Cells[24, 2].Value = "有效";
        //        //sheet.Cells[25, 2].Value = "内销配套";
        //        //sheet.Cells[26, 2].Value = "106全钢外胎";
        //        sheet.PivotTables[0].PageFields[0].Items[0].Text = "有效";
        //        sheet.PivotTables[0].PageFields[1].Items[0].Text = "内销配套";
        //    //    var a = sheet.PivotTables[0].Filters[0].Value1;
        //        package.Save();
        //    }
        //}
    }
}
