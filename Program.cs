using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DocJournalParser
{
    internal class Program
    {
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int cmdShow);

        private static void Main(string[] args)
        {
            Console.WriteLine("App started");

            string excelFileName = "БВ в ОПАК. Роспись БВ.xlsx";
            string docFileName = "БВ в ОПАК. Роспись БВ.doc";
            //docFileName = "Test.doc";

            KillWordAndExcel();
            WordHandler wordHandler = new WordHandler(docFileName);
            ExcelHandler excelHandler = new ExcelHandler();
            AutorPatterns autorPatterns = new AutorPatterns();
            LineParser lineParser = new LineParser(autorPatterns);

            MaximizeWindow();

            foreach (string line in wordHandler.ReadLines())
            {
                JDiscription jDiscription = lineParser.Parse(line);
                excelHandler.AddRow(jDiscription);
            }

            wordHandler.Quit();
            excelHandler.SaveFile(excelFileName);
            Console.WriteLine("All tasks ended");
            Console.ReadKey();
        }

        private static void KillWordAndExcel()
        {
            string[] processNames = new string[] { "winword", "excel" };
            foreach (string processName in processNames)
            {
                foreach (Process process in Process.GetProcessesByName(processName))
                {
                    try
                    {
                        process.Kill();
                        process.WaitForExit();
                    }
                    catch (Win32Exception winException)
                    {
                        Console.WriteLine(winException);
                    }
                    catch (InvalidOperationException invalidException)
                    {
                        Console.WriteLine(invalidException);
                    }
                }
            }
        }

        private static void MaximizeWindow()
        {
            Process p = Process.GetCurrentProcess();
            ShowWindow(p.MainWindowHandle, 3);
        }
    }
}