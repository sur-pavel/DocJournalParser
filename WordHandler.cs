using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace DocJournalParser
{
    public class WordHandler
    {
        private Document objDoc;
        private Application wordApp;

        public WordHandler(string fileName)
        {
            wordApp = new Application();
            string appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            objDoc = wordApp.Documents.Open(appPath + @"\" + fileName, ReadOnly: false);
        }

        internal List<string> ReadLines()

        {
            Console.WriteLine("Reading of lines in doc begins");
            List<string> data = new List<string>();
            for (int i = 0; i < objDoc.Paragraphs.Count; i++)
            {
                Range parRange = objDoc.Paragraphs[i + 1].Range;
                string line = parRange.Text.Trim();

                Match yearNumber = Regex.Match(line, @"(\d{4})*\/?\d{4}-\d\/?\d?");

                if (line != string.Empty && !yearNumber.Success)
                {
                    data.Add(line);
                }
            }
            Console.WriteLine("Reading of lines in doc ends");
            return data;
        }

        private string GetItalicText(Range parRange)
        {
            Range tempRange = parRange;
            tempRange.Find.Forward = true;
            tempRange.Find.Format = true;
            tempRange.Find.Font.Italic = 1;
            tempRange.Find.Execute();
            return tempRange.Text.Trim();
        }

        internal void Quit()
        {
            objDoc.Close();
            wordApp.Quit();
        }
    }
}