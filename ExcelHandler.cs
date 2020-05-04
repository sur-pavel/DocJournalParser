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
        private Excel.Application xlApp;
        private Excel.Workbook xlWorkBook;
        private Excel.Worksheet xlWorkSheet;
        private int row = 2;

        internal ExcelHandler()
        {
            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Add();
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Excel.Range cells = xlWorkBook.Worksheets[1].Cells;
            cells.NumberFormat = "@";
            cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            xlWorkSheet.Cells[1, 1] = "Фамилия";
            xlWorkSheet.Cells[1, 2] = "Инициалы";
            xlWorkSheet.Cells[1, 3] = "Чин";
            xlWorkSheet.Cells[1, 4] = "Инвертирование";
            xlWorkSheet.Cells[1, 5] = "Заглавие";
            xlWorkSheet.Cells[1, 6] = "Сведения к заглавию";
            xlWorkSheet.Cells[1, 7] = "Год";
            xlWorkSheet.Cells[1, 8] = "Том";
            xlWorkSheet.Cells[1, 9] = "Номер";
            xlWorkSheet.Cells[1, 10] = "Страницы";
            xlWorkSheet.Cells[1, 11] = "Примечание";
            xlWorkSheet.Cells[1, 12] = "Год полн. публ.";
            xlWorkSheet.Cells[1, 13] = "Том полн. публ.";
            xlWorkSheet.Cells[1, 14] = "Номер полн. публ.";
            xlWorkSheet.Cells[1, 15] = "Стр. ст. полн. публ.";
        }

        internal void AddRow(JDiscription jDiscription)
        {
            xlWorkSheet.Cells[row, 1] = jDiscription.LastName;
            xlWorkSheet.Cells[row, 2] = jDiscription.Initials;
            xlWorkSheet.Cells[row, 3] = jDiscription.Rank;
            xlWorkSheet.Cells[row, 4] = jDiscription.Invertion;
            xlWorkSheet.Cells[row, 5] = jDiscription.Title;
            xlWorkSheet.Cells[row, 6] = jDiscription.TitleInfo;
            xlWorkSheet.Cells[row, 7] = jDiscription.Year;
            xlWorkSheet.Cells[row, 8] = jDiscription.Volume;
            xlWorkSheet.Cells[row, 9] = jDiscription.Number;
            xlWorkSheet.Cells[row, 10] = jDiscription.Pages;
            xlWorkSheet.Cells[row, 11] = jDiscription.Notes;
            xlWorkSheet.Cells[row, 12] = jDiscription.FullPubYear;
            xlWorkSheet.Cells[row, 13] = jDiscription.FullPubVolume;
            xlWorkSheet.Cells[row, 14] = jDiscription.FullPubNumber;
            xlWorkSheet.Cells[row, 15] = jDiscription.FullPubPageRange;

            row++;
        }

        internal void SaveFile(string fileName)
        {
            xlWorkBook.SaveAs(fileName, AccessMode: Excel.XlSaveAsAccessMode.xlNoChange,
                ConflictResolution: Excel.XlSaveConflictResolution.xlLocalSessionChanges);
            xlWorkBook.Close(SaveChanges: true);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            Console.WriteLine("Excel file created");
        }
    }
}