using System;
using System.Collections.Generic;
using DataTransformation.Interface;
using DataTransformation.Models;
using DataTransformation.Librairie;

namespace DataTransformation.CollectionData
{
    class CollectionOfData
    {
        public List<ExcelModels> oExcelModelsList = new List<ExcelModels>();
        private IDataInterface oSourceFolder;

        public CollectionOfData(IDataInterface _SourceFolder, string Path)
        {
            this.oSourceFolder = _SourceFolder;            
        }

        public  string GetDataAsync( string Path)
        {
            oSourceFolder.ReadXLSFile(Path);

            oExcelModelsList = ((ExcelFolder)oSourceFolder).oExcelModelsList;
            return "oExcelModelsListTest";
        }

        public  string DisplayData(string _resourcesData ="Excel")
        {
            string Info = string.Empty;
            Console.WriteLine(Environment.NewLine + Environment.NewLine + " Begin " + _resourcesData);
            foreach (var item in oExcelModelsList)
            {
                Info = "Firstname : " + item.FirstName + "   LastName : " + item.LastName + "       Phone : " + item.Phone + Environment.NewLine;
                Console.WriteLine(Info);
                Info = string.Empty;
            }

            oExcelModelsList.Clear();
            Console.WriteLine(" end " + _resourcesData + Environment.NewLine + "**********************************");
            return Info;
        }
    }
}
