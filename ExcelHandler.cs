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

            xlWorkSheet.Cells[1, 1] = "Номер записи";
            xlWorkSheet.Cells[1, 2] = "Фамилия";
            xlWorkSheet.Cells[1, 3] = "Разночтение";
            xlWorkSheet.Cells[1, 4] = "Инициалы";
            xlWorkSheet.Cells[1, 5] = "Чин";
            xlWorkSheet.Cells[1, 6] = "Инвертирование";
            xlWorkSheet.Cells[1, 7] = "Заглавие";
            xlWorkSheet.Cells[1, 8] = "Сведения к заглавию";
            xlWorkSheet.Cells[1, 9] = "Свед. об отв-ти";
            xlWorkSheet.Cells[1, 10] = "Функция п. ред-ра";
            xlWorkSheet.Cells[1, 11] = "Фамилия п. ред-ра";
            xlWorkSheet.Cells[1, 12] = "Разночтение п. ред-ра";
            xlWorkSheet.Cells[1, 13] = "Инициалы п. ред-ра";
            xlWorkSheet.Cells[1, 14] = "Чин п. ред-ра";
            xlWorkSheet.Cells[1, 15] = "Инверт. п. ред-ра";
            xlWorkSheet.Cells[1, 16] = "Год";
            xlWorkSheet.Cells[1, 17] = "Том";
            xlWorkSheet.Cells[1, 18] = "Номер";
            xlWorkSheet.Cells[1, 19] = "Страницы";
            xlWorkSheet.Cells[1, 20] = "№ пагинация";
            xlWorkSheet.Cells[1, 21] = "Примечание";
            xlWorkSheet.Cells[1, 22] = "Полная публикация";
        }

        internal void AddRow(JDiscription jDiscription)
        {
            xlWorkSheet.Cells[row, 1] = jDiscription.DеscriptionNumber;
            xlWorkSheet.Cells[row, 2] = jDiscription.Autor.LastName;
            xlWorkSheet.Cells[row, 3] = jDiscription.Autor.LNameVariation;
            xlWorkSheet.Cells[row, 4] = jDiscription.Autor.Initials;
            xlWorkSheet.Cells[row, 5] = jDiscription.Autor.Rank;
            xlWorkSheet.Cells[row, 6] = jDiscription.Autor.Invertion;
            xlWorkSheet.Cells[row, 7] = jDiscription.Title;
            xlWorkSheet.Cells[row, 8] = jDiscription.TitleInfo;
            xlWorkSheet.Cells[row, 9] = jDiscription.Editors;
            xlWorkSheet.Cells[row, 10] = jDiscription.FirstEditor.Function;
            xlWorkSheet.Cells[row, 11] = jDiscription.FirstEditor.LastName;
            xlWorkSheet.Cells[row, 12] = jDiscription.FirstEditor.LNameVariation;
            xlWorkSheet.Cells[row, 13] = jDiscription.FirstEditor.Initials;
            xlWorkSheet.Cells[row, 14] = jDiscription.FirstEditor.Rank;
            xlWorkSheet.Cells[row, 15] = jDiscription.FirstEditor.Invertion;
            xlWorkSheet.Cells[row, 16] = jDiscription.Year;
            xlWorkSheet.Cells[row, 17] = jDiscription.JVolume;
            xlWorkSheet.Cells[row, 18] = jDiscription.JNumber;
            xlWorkSheet.Cells[row, 19] = jDiscription.Pages;
            xlWorkSheet.Cells[row, 20] = jDiscription.Pagination;            
            xlWorkSheet.Cells[row, 21] = jDiscription.Notes;
            xlWorkSheet.Cells[row, 22] = jDiscription.FullPublication;

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