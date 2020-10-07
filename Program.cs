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
    internal static class Program
    {
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int cmdShow);

        private static void Main(string[] args)
        {
            Console.WriteLine("App started");
            KillProcesses("winword", "excel");
            MaximizeConsoleWindow();

            Patterns patterns = new Patterns();
            LineParser lineParser = new LineParser(patterns);

            string currentDirPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            DirectoryInfo dirFilesInfo = new DirectoryInfo(currentDirPath);

            foreach (FileInfo fileInfo in dirFilesInfo.GetFiles("*.doc"))
            {
                Console.WriteLine("Parse file: " + fileInfo.FullName);
                WordHandler wordHandler = new WordHandler(patterns);
                ExcelHandler excelHandler = new ExcelHandler();

                foreach (string line in wordHandler.ReadLines(fileInfo.Name))
                {
                    JDiscription jDiscription = lineParser.Parse(line);
                    excelHandler.AddRow(jDiscription);
                }
                wordHandler.Quit();
                excelHandler.SaveFile(fileInfo.FullName.Replace(".docx", ".xlsx"));
            }

            Console.WriteLine("All tasks ended");
//            Console.WriteLine("\n\nPress any key...");
//            Console.ReadKey();
        }

        private static void KillProcesses(params string[] processNames)
        {
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

        private static void MaximizeConsoleWindow()
        {
            Process process = Process.GetCurrentProcess();
            ShowWindow(process.MainWindowHandle, 3);
        }
    }
}