﻿using System;
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
            //RewriteParagraphs();

            Console.WriteLine("Reading of lines in doc begins");
            List<string> data = new List<string>();
            int linesCount = 0;
            int i = 0;
            int index = 0;
            for (; i < objDoc.Paragraphs.Count; i++)
            {
                linesCount++;
                Range parRange = objDoc.Paragraphs[i + 1].Range;
                string line = parRange.Text.Trim();

                if (patterns.yearNumberPattern(line).Success && !patterns.yearPattern(line).Success)
                {
                    index = 0;
                }
                else if (patterns.MatchOddPages(line).Success)
                {
                    index++;
                }
                else if (!string.IsNullOrEmpty(line) && patterns.linePattern(line).Success)
                {
                    int recordIndex = int.Parse(Regex.Match(line, @"^\d+").Value) - index;
                    line = Regex.Replace(line, @"^\d+.\s?", recordIndex + ". ");
                    data.Add(line);
                    Console.WriteLine($"index = {recordIndex} \n {line}");
                }
                else
                {
                    Console.WriteLine(line);
                }
            }
            Console.WriteLine($"linesCount: {linesCount} Paragraphs.Count = {objDoc.Paragraphs.Count}");
            Console.WriteLine("Reading of lines in doc ends");
            return data;
        }

        private void RewriteParagraphs()
        {
            wordApp.Selection.Find.ClearFormatting();
            wordApp.Selection.Find.Execute(FindText: "^p", ReplaceWith: "");
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