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
    internal class WordHandler
    {
        private Document objDoc;
        private Application wordApp;
        private Patterns patterns;
        private string appPath;

        internal WordHandler(Patterns patterns)
        {
            this.patterns = patterns;
            wordApp = new Application();
            appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }

        internal List<string> ReadLines(string fileName)
        {
            objDoc = wordApp.Documents.Open(appPath + @"\" + fileName, ReadOnly: false);
            Console.WriteLine("Reading of lines in doc begins");
            List<string> data = new List<string>();
            for (int i = 0; i < objDoc.Paragraphs.Count; i++)
            {
                Range parRange = objDoc.Paragraphs[i + 1].Range;
                string line = parRange.Text.Trim();

                Match yearNumberMatch = Regex.Match(line, patterns.yearNumberPattern);
                Match oddPagesMatch = Regex.Match(line, patterns.oddPagesPattern);
                Match lineMatch = Regex.Match(line, patterns.linePattern);

                if (!yearNumberMatch.Success && !oddPagesMatch.Success)
                {
                    if (lineMatch.Success)
                    {
                        data.Add(line);
                    }
                    else
                    {
                        //i--;
                    }
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