using DataTransformation.Librairie;
using System;
using System.Collections.Generic;
using System.Linq;
using DataTransformation.Interface;
using DataTransformation.CollectionData;
using DataTransformation.Models;
using System.IO;
using System.Configuration;

namespace DataTransformation
{
    class Program
    {
        static void Main(string[] args)
        {
            IDataInterface oDataInterface;

            var factory = new Factory();
            var oExcelModelsList = new List<ExcelModels>();

            var _Path = ConfigurationManager.AppSettings["path"];
            var FolderList = Dir(_Path);
            List<string> FileList;
         
            foreach (string Folder in FolderList)
            {
                FileList = Dir(Folder);
                foreach (var File in FileList)
                {
                    oDataInterface = factory.GetObject(Path.GetExtension(File), File);
                    if (oDataInterface != null)
                    {
                        CollectionOfData oCollectionOfData = new CollectionOfData(oDataInterface, File);
                        var t = oCollectionOfData.GetDataAsync( File);
                        oCollectionOfData.DisplayData();

                    }
                }
            }
            Console.ReadKey();
        }
        static List<string> Dir(string directory)
        {
            string[] files;

            files = Directory.GetFileSystemEntries(directory);

            return files.ToList();

        }
    }
}
