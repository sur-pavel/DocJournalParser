using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace DocJournalParser
{
    internal class ExcelHandler
    {
        private Excel.Application xlApp = new Excel.Application();
        private Excel.Workbook xlWorkBook;
        private Excel.Worksheet xlWorkSheet;
        private object misValue = Missing.Value;
        private int row = 2;

        public ExcelHandler()
        {
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Cells[1, 1] = "Фамилия";
            xlWorkSheet.Cells[1, 2] = "Инициалы";
            xlWorkSheet.Cells[1, 3] = "Заглавие";
            xlWorkSheet.Cells[1, 4] = "Сведения к заглавию";
            xlWorkSheet.Cells[1, 5] = "Год";
            xlWorkSheet.Cells[1, 6] = "Том";
            xlWorkSheet.Cells[1, 7] = "Номер";
            xlWorkSheet.Cells[1, 8] = "Страницы";
            xlWorkSheet.Cells[1, 9] = "Примечание";
            xlWorkSheet.Cells[1, 10] = "Год полн. публ.";
            xlWorkSheet.Cells[1, 11] = "Том полн. публ.";
            xlWorkSheet.Cells[1, 12] = "Номер полн. публ.";
            xlWorkSheet.Cells[1, 13] = "Стр. ст. полн. публ.";

            //xlWorkSheet.get_Range("A1", "G1").Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
        }

        internal void AddRow(JDiscription jDiscription)
        {
            xlWorkSheet.Cells[row, 1] = jDiscription.LastName;
            xlWorkSheet.Cells[row, 2] = jDiscription.Initials;
            xlWorkSheet.Cells[row, 3] = jDiscription.Title;
            xlWorkSheet.Cells[row, 4] = jDiscription.TitleInfo;
            xlWorkSheet.Cells[row, 5] = jDiscription.Year;
            xlWorkSheet.Cells[row, 6] = jDiscription.Volume;
            xlWorkSheet.Cells[row, 7] = jDiscription.Number;
            xlWorkSheet.Cells[row, 8] = jDiscription.Pages;
            xlWorkSheet.Cells[row, 9] = jDiscription.Notes;
            xlWorkSheet.Cells[row, 10] = jDiscription.FullPubYear;
            xlWorkSheet.Cells[row, 11] = jDiscription.FullPubVolume;
            xlWorkSheet.Cells[row, 12] = jDiscription.FullPubNumber;
            xlWorkSheet.Cells[row, 13] = jDiscription.FullPubPageRange;

            row++;
        }

        internal void SaveFile(string fileName)
        {
            string appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            xlWorkBook.SaveAs(appPath + @"\" + fileName, ConflictResolution: Excel.XlSaveConflictResolution.xlLocalSessionChanges);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}