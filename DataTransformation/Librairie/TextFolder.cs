using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataTransformation.Models;
using DataTransformation.Interface;

namespace DataTransformation.Librairie
{
    public class TextFolder :IDataInterface
    {
        public List<ExcelModels> oTexteModelsList;
        public TextFolder()
        {
            oTexteModelsList = new List<ExcelModels>();
        }
        public async Task<List<ExcelModels>> ReadFile(string pathFile)
        {
            string[] lines = System.IO.File.ReadAllLines(pathFile);

            foreach (var line in lines)
            {
                ExcelModels oTexteModels = new ExcelModels();
                string[] temp = line.Split(' ');

                oTexteModels.FirstName = temp[0];
                oTexteModels.LastName = temp[1];
                oTexteModels.Phone = temp[2];
                //oTexteModels.BirthDate = temp[3];

                oTexteModelsList.Add(oTexteModels);

            }
            return oTexteModelsList;

        }

        public Task ReadXLSFile(string pathFile)
        {
            throw new NotImplementedException();
        }
    }
}
