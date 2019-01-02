using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataTransformation.Models;
using Microsoft.Office.Interop.Excel;
using DataTransformation.Interface;

namespace DataTransformation.Librairie
{
    public class ExcelFolder : IDataInterface
    {


        #region prepreties
        public string FileName { get; set; }

        public List<ExcelModels> oExcelModelsList;
        #endregion

        public ExcelFolder()
        {

        }

        #region Folder Manipulation
        //public async Task<List<ExcelModels>> ReadFile(string pathFile)
        //{

        //    _Application application = new ApplicationClass();
        //    _Workbook workbook = application.Workbooks.Open(pathFile, Type.Missing, Type.Missing, Type.Missing,
        //                                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                                    Type.Missing, Type.Missing, Type.Missing);
        //    _Worksheet worksheet = (_Worksheet)workbook.ActiveSheet;

        //    for (int i = 1; i < worksheet.Rows.Count + 1; i++)
        //    {
        //        ExcelModels oExcelFolder = new ExcelModels();
        //        // A la première cellule vide, sortir de la boucle :
        //        if ((worksheet.Cells[i, 1] == null || ((Range)worksheet.Cells[i, 1]).Value2 == null) &&
        //            (worksheet.Cells[i, 2] == null || ((Range)worksheet.Cells[i, 2]).Value2 == null) &&
        //            (worksheet.Cells[i, 3] == null || ((Range)worksheet.Cells[i, 3]).Value2 == null))
        //        {
        //            break;
        //        }

        //        if (i > 1)
        //        {
        //            oExcelFolder.FirstName = ((Range)worksheet.Cells[i, 1]).Value2.ToString();
        //            oExcelFolder.LastName = ((Range)worksheet.Cells[i, 2]).Value2.ToString();
        //            oExcelFolder.Phone = ((Range)worksheet.Cells[i, 3]).Value2.ToString();

        //            oExcelModelsList.Add(oExcelFolder);
        //        }
        //    }

        //    workbook.Close(false, Type.Missing, Type.Missing);
        //    application.Quit();
        //    return oExcelModelsList;

        //}

        #endregion

        #region NewFunctionnality

        private Workbook pBook = null;
        private Microsoft.Office.Interop.Excel.Application pApp = null;
        private Worksheet pSheet = null;

        public ExcelFolder(string PathXLSFile)
        {
            pApp = new Microsoft.Office.Interop.Excel.Application();
            pApp.Visible = false;
            pBook = pApp.Workbooks.Open(PathXLSFile);

            oExcelModelsList = new List<ExcelModels>();
        }

        public async Task ReadXLSFile(string pathFile)
        {
            int pSheetCount = pBook.Sheets.Count;
            var tasks = new List<Task>();

            try
            {
                for (int i = 1; i <= pSheetCount; i++)
                {
                    tasks.Add(ReadXLSSheets(i));
                }

                var continuation = Task.WhenAll(tasks);
                continuation.Wait();
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                pBook.Close(false, Type.Missing, Type.Missing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pBook);

                pApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pApp);

                pBook = null;
                pApp = null;
            }
        }
        private async Task ReadXLSSheets(int sheetNumber)
        {
            pSheet = (Worksheet)pBook.Sheets[sheetNumber]; // Explicit cast is not required here
            Range range = pSheet.UsedRange;

            ExcelModels oExcelFolder = null;
            string str;
            string HeaderName = null;
            bool EmptyLine = false;

            if (IsValidHeader(range))
            {
                for (int rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
                {
                    oExcelFolder = new ExcelModels();
                    for (int cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                    {
                        str = (range.Cells[rCnt, cCnt] as Range).Value2.ToString();
                        if (string.IsNullOrEmpty(str))
                        {
                            EmptyLine = true;
                        }

                        //chercher le header de cette line 
                        HeaderName = (range.Cells[1, cCnt] as Range).Value2.ToString();
                        //remplir proprite equivalente
                        FIllExcelObject(ref oExcelFolder, HeaderName, str);

                        if (cCnt == range.Columns.Count)
                        {
                            if (!EmptyLine)
                            {
                                //inserer dans la list pour envoie database
                                oExcelModelsList.Add(oExcelFolder);
                            }
                            else
                            {
                                // Ecrire dans le fichier log
                            }
                        }
                    }
                }
            }
            else
            {
                // Ecrire dans le fichier log
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pSheet);
        }

        private void FIllExcelObject(ref ExcelModels excelFolder, string header, string value)
        {
            switch (header.ToLower())
            {
                case "firstname":
                    excelFolder.FirstName = value;
                    break;
                case "lastname":
                    excelFolder.LastName = value;
                    break;
                case "phone":
                    excelFolder.Phone = value;
                    break;
            }
        }
        private bool IsValidHeader(Range range)
        {
            bool ValidHeader = true;
            string str;

            for (int cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                str = (string)(range.Cells[1, cCnt] as Range).Value2;
                if (string.IsNullOrEmpty(str))
                {
                    ValidHeader = false;
                }
            }
            return ValidHeader;
        }
        #endregion NewFunctionnality
    }
}
