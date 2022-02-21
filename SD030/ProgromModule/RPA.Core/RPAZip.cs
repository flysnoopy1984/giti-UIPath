using Ionic.Zip;
using System;
using System.Collections.Generic;

using System.Text;

namespace RPA.Core
{
    public class RPAZip
    {
        public static string ZipDir(string dirPath,string zipName = "data.zip")
        {
            string zipPath = dirPath + "\\"+zipName;
            using (ZipFile zip = new ZipFile())
            {
                zip.AddDirectory(dirPath);
                zip.Save(zipPath);
            }
            return zipPath;
        }
    }
}
