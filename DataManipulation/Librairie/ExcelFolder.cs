using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataManipulation.Models;
using Microsoft.Office.Interop.Excel;

namespace DataManipulation.Librairie
{
    public class ExcelFolder
    {
        #region prepreties
        public string FileName { get; set; }

        public List<ExcelModels> oExcelModelsList ;
        #endregion

        #region Folder Manipulation
        public void ReadFile(string pathFile)
        {
            oExcelModelsList = new List<ExcelModels>();

            
            _Application application = new ApplicationClass();
            _Workbook workbook = application.Workbooks.Open(pathFile, Type.Missing, Type.Missing, Type.Missing,
                                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                            Type.Missing, Type.Missing, Type.Missing);
            _Worksheet worksheet = (_Worksheet)workbook.ActiveSheet;
            Range range;

            for (int i = 1; i < worksheet.Rows.Count + 1; i++)
            {
                ExcelModels oExcelFolder = new ExcelModels();
                // A la première cellule vide, sortir de la boucle :
                if ((worksheet.Cells[i, 1] == null || ((Range)worksheet.Cells[i, 1]).Value2 == null) &&
                    (worksheet.Cells[i, 2] == null || ((Range)worksheet.Cells[i, 2]).Value2 == null) &&
                    (worksheet.Cells[i, 3] == null || ((Range)worksheet.Cells[i, 3]).Value2 == null))
                {
                    break;
                }

                if (i > 1)
                {
                    oExcelFolder.FirstName = ((Range)worksheet.Cells[i, 1]).Value2.ToString();
                    oExcelFolder.LastName = ((Range)worksheet.Cells[i, 2]).Value2.ToString();
                    oExcelFolder.Phone = ((Range)worksheet.Cells[i, 3]).Value2.ToString();

                    oExcelModelsList.Add(oExcelFolder);
                }
                

                // Lire les cellules :
                Console.WriteLine("{0}\t{1}", ((Range)worksheet.Cells[i, 1]).Value2.ToString(), ((Range)worksheet.Cells[i, 2]).Value2.ToString());
            }

            workbook.Close(false, Type.Missing, Type.Missing);
            application.Quit();

        }
        #endregion



    }
}
