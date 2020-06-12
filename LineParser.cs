using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace DocJournalParser
{
    public class LineParser
    {
        private Patterns patterns;

        public LineParser(Patterns patterns)
        {
            this.patterns = patterns;
        }

        public JDiscription Parse(string line)
        {
            JDiscription jDiscription = new JDiscription();
            jDiscription.DеscriptionNumber = int.Parse(Regex.Match(line, @"^\d+").Value);
            line = Regex.Replace(line, @"^\d+. ", "");
            line = line.Replace("//", " //").Replace("  ", " ").Replace(". ,", ".,");

            try
            {
                string articleData = line.Split(new string[] { " // " }, StringSplitOptions.None)[0];
                GetAutor(jDiscription, ref articleData);

                GetEditors(jDiscription, ref articleData);

                GetTitleAndTitleInfo(jDiscription, ref articleData);

                string journalData = line.Split(new string[] { " // " }, StringSplitOptions.None)[1];
                string fullPubInfo = string.Empty;
                ExtractFullPubInfo(ref journalData, ref fullPubInfo);

                jDiscription.Year = ExtractProp(journalData, patterns.yearPattern);
                jDiscription.JVolume = ExtractProp(journalData, patterns.volumePattern, "Т. ");
                jDiscription.JNumber = ExtractProp(journalData, patterns.numberPattern, "№ ");
                jDiscription.Pages = GetPages(journalData);
                jDiscription.Notes = ExtractProp(journalData, patterns.notesPattern, "(", ")", ".");

                jDiscription.FullPubYear = ExtractProp(fullPubInfo, patterns.yearPattern, " ", ".");
                jDiscription.FullPubVolume = ExtractProp(fullPubInfo, patterns.volumePattern, "Т. ");
                jDiscription.FullPubNumber = ExtractProp(fullPubInfo, patterns.numberPattern, "№ ");
                jDiscription.FullPubPageRange = ExtractProp(fullPubInfo, patterns.pagesPattern, "С. ");
                jDiscription.FullPubPageRange = CleanPages(jDiscription.FullPubPageRange);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.WriteLine(line);
                Console.WriteLine(jDiscription.ToString() + "\n");
            }
            return jDiscription;
        }

        private void GetEditors(JDiscription jDiscription, ref string articleData)
        {
            Match editorsMatch = Regex.Match(articleData, patterns.editorsPattern);
            string necrologue = Regex.Match(articleData, patterns.necrologuePattern).Value;
            if (editorsMatch.Success && !Regex.IsMatch(articleData, patterns.exclusionPattern))
            {
                jDiscription.Editors = editorsMatch.Value;
                ReplaceIfNotNull(jDiscription.Editors, necrologue);
                jDiscription.Editors = Regex.Replace(jDiscription.Editors, @"\s?\/\s?", "");
                articleData = ReplaceIfNotNull(articleData, jDiscription.Editors);
                articleData += necrologue;
                GetFirstEditor(jDiscription);
                jDiscription.Editors = CleanUpString(jDiscription.Editors);
            }
        }

        private void GetFirstEditor(JDiscription jDiscription)
        {
            MatchCollection matchCollection = patterns.editorsCountPattern(jDiscription.Editors);
            int edCount = matchCollection.Count;
            if (edCount > 1)
            {
                jDiscription.FirstEdLastName = matchCollection[0].Value;
            }
            else
            {
                jDiscription.FirstEdLastName = jDiscription.Editors;
            }

            jDiscription.FirstEdFunc = Regex.Match(jDiscription.FirstEdLastName, patterns.editorFunc).Value;
            jDiscription.FirstEdLastName = ReplaceIfNotNull(jDiscription.FirstEdLastName, jDiscription.FirstEdFunc);
            foreach (Match match in patterns.InvertMathches(jDiscription.FirstEdLastName))
            {
                if (match.Success)
                {
                    jDiscription.FirstEdInvertion = "1";
                }
            }

            //from here
            jDiscription.FirstEdInitials = ExtractProp(jDiscription.FirstEdLastName, patterns.initialsPattern);
            jDiscription.FirstEdLastName = ReplaceIfNotNull(jDiscription.FirstEdLastName, jDiscription.FirstEdInitials);

            string[] lastNameSplit = Regex.Split(jDiscription.FirstEdLastName, @"(,\s)");

            if (lastNameSplit.Length > 1)
            {
                for (int i = 1; i < lastNameSplit.Length; i++)
                {
                    jDiscription.FirstEdRank += lastNameSplit[i];
                    jDiscription.FirstEdLastName = ReplaceIfNotNull(jDiscription.FirstEdLastName, lastNameSplit[i]);
                }
                if (Regex.IsMatch(jDiscription.FirstEdRank, @"\.\s*$"))
                {
                    jDiscription.FirstEdRank = CleanUpString(jDiscription.FirstEdRank) + ".";
                }
                else
                {
                    jDiscription.FirstEdRank = CleanUpString(jDiscription.FirstEdRank);
                }
            }
            jDiscription.FirstEdLastName = CleanUpString(jDiscription.FirstEdLastName);
            jDiscription.FirstEdLastName = patterns.DeclineLastName(jDiscription.FirstEdLastName);
        }

        private void ExtractFullPubInfo(ref string journalData, ref string fullPubInfo)
        {
            string[] fullPubSplit = Regex.Split(journalData, "Полная публикация:");
            if (fullPubSplit.Length > 1)
            {
                fullPubInfo = fullPubSplit[1];
                journalData = journalData.Replace(fullPubInfo, "");
            }
        }

        private void GetTitleAndTitleInfo(JDiscription jDiscription, ref string articleData)
        {
            articleData = articleData.Trim();
            articleData = Regex.Replace(articleData, @"^\.+", "");
            Match reviewMatch = Regex.Match(articleData, patterns.reviewPattern);
            if (articleData.StartsWith("["))
            {
                jDiscription.Title = articleData;
            }
            else if (reviewMatch.Success && !articleData.StartsWith("[Рец. на"))
            {
                jDiscription.Title = articleData.Replace(reviewMatch.Value, "");
                jDiscription.TitleInfo = reviewMatch.Value;
            }
            else if (articleData.Split(new[] { ':' }, 2).Length > 1 && !articleData.StartsWith("[Рец. на"))
            {
                jDiscription.Title = articleData.Split(new[] { ':' }, 2)[0];
                jDiscription.TitleInfo = articleData.Split(new[] { ':' }, 2)[1];
                jDiscription.TitleInfo = Regex.Replace(jDiscription.TitleInfo, @"^\.? ", "");
            }
            else
            {
                jDiscription.Title = articleData;
            }
            jDiscription.Title = CleanUpString(jDiscription.Title);
            jDiscription.TitleInfo = CleanUpString(jDiscription.TitleInfo);
        }

        private string GetPages(string journalData)
        {
            string pagesPattern = patterns.pagesPattern;
            string pages = ExtractProp(journalData, pagesPattern, "С. ");

            if (string.IsNullOrEmpty(pages))
            {
                pages = ExtractProp(journalData, @"\d+ с\.");
            }
            pages = CleanPages(pages);

            return pages;
        }

        private static string CleanPages(string pages)
        {
            pages = Regex.Replace(pages, @"\.? $", "");
            pages = Regex.Replace(pages, @"\)\. ?$", ")");
            return pages;
        }

        private void GetAutor(JDiscription jDiscription, ref string articleData)
        {
            articleData = articleData.Trim();
            foreach (Match mPattern in patterns.AutorMatches(articleData))
            {
                if (mPattern.Success)
                {
                    jDiscription.LastName = mPattern.Value;
                    if (!string.IsNullOrEmpty(jDiscription.LastName))
                    {
                        articleData = ReplaceIfNotNull(articleData, jDiscription.LastName);
                        ParseAutor(jDiscription, mPattern);
                    }
                }
            }
        }

        private void ParseAutor(JDiscription jDiscription, Match mPattern)
        {
            foreach (Match match in patterns.InvertMathches(mPattern.Value))
            {
                if (match.Success)
                {
                    jDiscription.Invertion = "1";
                }
            }
            if (!patterns.DetectedMatches(jDiscription.LastName).Contains(mPattern))
            {
                jDiscription.Initials = ExtractProp(jDiscription.LastName, patterns.initialsPattern);
                jDiscription.LastName = ReplaceIfNotNull(jDiscription.LastName, jDiscription.Initials);
            }
            string[] lastNameSplit = Regex.Split(jDiscription.LastName, @"(,\s)");

            if (lastNameSplit.Length > 1)
            {
                for (int i = 1; i < lastNameSplit.Length; i++)
                {
                    jDiscription.Rank += lastNameSplit[i];
                    jDiscription.LastName = ReplaceIfNotNull(jDiscription.LastName, lastNameSplit[i]);
                }
                if (Regex.IsMatch(jDiscription.Rank, @"\.\s*$"))
                {
                    jDiscription.Rank = CleanUpString(jDiscription.Rank) + ".";
                }
                else
                {
                    jDiscription.Rank = CleanUpString(jDiscription.Rank);
                }
            }
            jDiscription.LastName = CleanUpString(jDiscription.LastName);
        }

        private string CleanUpString(string str)
        {
            return Regex.Replace(str, patterns.cleanUpPattern, "").Trim();
        }

        private string ReplaceIfNotNull(string inputString, string replaceString)
        {
            if (!string.IsNullOrEmpty(replaceString))
            {
                inputString = inputString.Replace(replaceString, "");
            }
            return inputString;
        }

        private string ExtractProp(string inputString, string matchPattern, params string[] replaceStrings)
        {
            Match match = Regex.Match(inputString, matchPattern);
            string returnValue = match.Value;
            if (!string.IsNullOrEmpty(returnValue))
            {
                foreach (string repStr in replaceStrings)
                {
                    returnValue = returnValue.Replace(repStr, "");
                }
            }
            return returnValue.Trim();
        }
    }
}