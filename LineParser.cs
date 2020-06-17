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
                string articleData = line.Split(new string[] {" // "}, StringSplitOptions.None)[0];
                GetAutor(jDiscription, ref articleData);
                GetEditors(jDiscription, ref articleData);
                GetTitleAndTitleInfo(jDiscription, ref articleData);

                string journalData = line.Split(new string[] {" // "}, StringSplitOptions.None)[1];
                string fullPubInfo = string.Empty;
                ExtractFullPubInfo(ref journalData, ref fullPubInfo);

                jDiscription.Year = ExtractProp(journalData, Patterns.YearPattern);
                jDiscription.JVolume = ExtractProp(journalData, Patterns.VolumePattern, "Т. ");
                jDiscription.JNumber = ExtractProp(journalData, Patterns.NumberPattern, "№ ");
                jDiscription.Pages = GetPages(journalData);
                jDiscription.Notes = ExtractProp(journalData, Patterns.NotesPattern, "(", ")", ".");

                jDiscription.FullPubYear = ExtractProp(fullPubInfo, Patterns.YearPattern, " ", ".");
                jDiscription.FullPubVolume = ExtractProp(fullPubInfo, Patterns.VolumePattern, "Т. ");
                jDiscription.FullPubNumber = ExtractProp(fullPubInfo, Patterns.NumberPattern, "№ ");
                jDiscription.FullPubPageRange = ExtractProp(fullPubInfo, Patterns.PagesPattern, "С. ");
                jDiscription.FullPubPageRange = CleanPages(jDiscription.FullPubPageRange);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.WriteLine(line);
                Console.WriteLine(jDiscription + "\n");
            }

            return jDiscription;
        }

        private void GetEditors(JDiscription jDiscription, ref string articleData)
        {
            if (!articleData.Contains("[Рец. на"))
            {
                Match editorsMatch = Regex.Match(articleData, Patterns.EditorsPattern);
                string necrologue = Regex.Match(articleData, Patterns.NecrologuePattern).Value;
                if (editorsMatch.Success && !Regex.IsMatch(articleData, Patterns.ExclusionPattern))
                {
                    jDiscription.Editors = editorsMatch.Value;
                    articleData = ReplaceIfNotNull(articleData, jDiscription.Editors);
                    jDiscription.Editors = ReplaceIfNotNull(jDiscription.Editors, necrologue);
                    if (!string.IsNullOrEmpty(necrologue))
                    {
                        articleData += " : " + necrologue;
                    }

                    jDiscription.Editors = Regex.Replace(jDiscription.Editors, @"\s?\/\s?", "");
                    GetFirstEditor(jDiscription);
                    jDiscription.Editors = CleanUpString(jDiscription.Editors);
                }
            }
        }

        private void GetFirstEditor(JDiscription jDiscription)
        {
            MatchCollection matchCollection = patterns.EditorsCountPattern(jDiscription.Editors);
            int edCount = matchCollection.Count;
            if (edCount > 1)
            {
                jDiscription.FirstEditor.LastName = matchCollection[0].Value;
            }
            else
            {
                jDiscription.FirstEditor.LastName = jDiscription.Editors;
            }

            jDiscription.FirstEditor.Function =
                jDiscription.Editors.Split(new string[] {matchCollection[0].Value}, StringSplitOptions.None)[0];
            jDiscription.FirstEditor.LastName =
                ReplaceIfNotNull(jDiscription.FirstEditor.LastName, jDiscription.FirstEditor.Function);

            foreach (Match match in patterns.InvertMathches(jDiscription.FirstEditor.LastName))
            {
                if (match.Success)
                {
                    jDiscription.FirstEditor.Invertion = "1";
                }
            }

            jDiscription.FirstEditor.Initials =
                ExtractProp(jDiscription.FirstEditor.LastName, Patterns.InitialsPattern);
            jDiscription.FirstEditor.LastName =
                ReplaceIfNotNull(jDiscription.FirstEditor.LastName, jDiscription.FirstEditor.Initials);

            jDiscription.FirstEditor.LastName = CleanUpString(jDiscription.FirstEditor.LastName);
            jDiscription.FirstEditor.LastName = patterns.DeclineEditorNames(jDiscription.FirstEditor.LastName);

            jDiscription.FirstEditor.Rank = ExtractProp(jDiscription.FirstEditor.LastName, Patterns.RankPattern);
            jDiscription.FirstEditor.LastName =
                ReplaceIfNotNull(jDiscription.FirstEditor.LastName, jDiscription.FirstEditor.Rank);
            jDiscription.FirstEditor.LastName = jDiscription.FirstEditor.LastName.Trim();
        }

        private void ExtractFullPubInfo(ref string journalData, ref string fullPubInfo)
        {
            string[] fullPubSplit = journalData.Split(new string[] {"Полная публикация:"}, StringSplitOptions.None);
            if (fullPubSplit.Length <= 1) return;
            fullPubInfo = fullPubSplit[1];
            journalData = journalData.Replace(fullPubInfo, "");
        }

        private void GetTitleAndTitleInfo(JDiscription jDiscription, ref string articleData)
        {
            articleData = articleData.Trim();
            articleData = Regex.Replace(articleData, @"^\.+", "");
            Match reviewMatch = Regex.Match(articleData, Patterns.ReviewPattern);
            if (articleData.StartsWith("["))
            {
                jDiscription.Title = articleData;
            }
            else if (reviewMatch.Success && !articleData.StartsWith("[Рец. на"))
            {
                jDiscription.Title = articleData.Replace(reviewMatch.Value, "");
                jDiscription.TitleInfo = reviewMatch.Value;
            }
            else if (articleData.Split(new[] {':'}, 2).Length > 1 && !articleData.StartsWith("[Рец. на"))
            {
                jDiscription.Title = articleData.Split(new[] {':'}, 2)[0];
                jDiscription.TitleInfo = articleData.Split(new[] {':'}, 2)[1];
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
            string pagesPattern = Patterns.PagesPattern;
            string pages = ExtractProp(journalData, pagesPattern, "С. ");

            if (string.IsNullOrEmpty(pages))
            {
                pages = ExtractProp(journalData, @"\d+ с\.");
            }

            pages = CleanPages(pages);

            return pages;
        }

        private string CleanPages(string pages)
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
                    jDiscription.Autor.LastName = mPattern.Value;
                    if (!string.IsNullOrEmpty(jDiscription.Autor.LastName))
                    {
                        articleData = ReplaceIfNotNull(articleData, jDiscription.Autor.LastName);
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
                    jDiscription.Autor.Invertion = "1";
                }
            }

            if (!patterns.DetectedMatches(jDiscription.Autor.LastName).Contains(mPattern))
            {
                jDiscription.Autor.Initials = ExtractProp(jDiscription.Autor.LastName, Patterns.InitialsPattern);
                jDiscription.Autor.LastName =
                    ReplaceIfNotNull(jDiscription.Autor.LastName, jDiscription.Autor.Initials);
            }

            string[] lastNameSplit =
                jDiscription.Autor.LastName.Split(new string[] {@"(,\s)"}, StringSplitOptions.None);

            if (lastNameSplit.Length > 1)
            {
                for (int i = 1; i < lastNameSplit.Length; i++)
                {
                    jDiscription.Autor.Rank += lastNameSplit[i];
                    jDiscription.Autor.LastName = ReplaceIfNotNull(jDiscription.Autor.LastName, lastNameSplit[i]);
                }

                if (Regex.IsMatch(jDiscription.Autor.Rank, @"\.\s*$"))
                {
                    jDiscription.Autor.Rank = CleanUpString(jDiscription.Autor.Rank) + ".";
                }
                else
                {
                    jDiscription.Autor.Rank = CleanUpString(jDiscription.Autor.Rank);
                }
            }

            jDiscription.Autor.LastName = CleanUpString(jDiscription.Autor.LastName);
        }

        private string CleanUpString(string str)
        {
            if (!Regex.IsMatch(str, Patterns.RankPattern + @"$"))
            {
                str = Regex.Replace(str, @"\.$", "");
            }
            return Regex.Replace(str, Patterns.CleanUpPattern, "").Trim();
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

        private string ReplaceIfNotNull(string inputString, string replaceString)
        {
            if (!string.IsNullOrEmpty(replaceString))
            {
                inputString = inputString.Replace(replaceString, "");
            }

            return inputString;
        }
    }
}