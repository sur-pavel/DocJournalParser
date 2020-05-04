using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
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

            KillWordAndExcel();

            MaximizeWindow();

            string currentDirPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            DirectoryInfo dirFilesInfo = new DirectoryInfo(currentDirPath);
            foreach (FileInfo fileInfo in dirFilesInfo.GetFiles("*.doc"))
            {
                Console.WriteLine(fileInfo.FullName);
                WordHandler wordHandler = new WordHandler(fileInfo.Name);
                ExcelHandler excelHandler = new ExcelHandler();
                Patterns autorPatterns = new Patterns();
                LineParser lineParser = new LineParser(autorPatterns);

                foreach (string line in wordHandler.ReadLines())
                {
                    JDiscription jDiscription = lineParser.Parse(line);
                    excelHandler.AddRow(jDiscription);
                }
                wordHandler.Quit();
                excelHandler.SaveFile(fileInfo.Name.Replace(".doc", ".xlsx"));
            }

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