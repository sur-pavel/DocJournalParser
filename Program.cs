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
            string excelFileName = "БВ в ОПАК. Роспись БВ.xlsx";
            string docFileName = "БВ в ОПАК. Роспись БВ.doc";
            KillWordAndExcel();
            WordHandler wordHandler = new WordHandler(docFileName);
            ExcelHandler excelHandler = new ExcelHandler();
            LineParser lineParser = new LineParser();

            MaximizeWindow();
            Console.WriteLine("App started");

            foreach (string line in wordHandler.ReadLines())
            {
                JDiscription jDiscription = lineParser.Parse(line);
                excelHandler.AddRow(jDiscription);
            }

            wordHandler.Quit();
            excelHandler.SaveFile(excelFileName);
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
                        process.WaitForExit(); // possibly with a timeout
                    }
                    catch (Win32Exception winException)
                    {
                        // process was terminating or can't be terminated - deal with it
                    }
                    catch (InvalidOperationException invalidException)
                    {
                        // process has already exited - might be able to let this one go
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