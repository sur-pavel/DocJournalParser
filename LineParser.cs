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
            jDiscription.DiscriptionNumber = int.Parse(Regex.Match(line, @"^\d+").Value);
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

                jDiscription.Year = ExtractProp(journalData, patterns.yearPattern, " ", ".");
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
                if (!string.IsNullOrEmpty(necrologue))
                {
                    jDiscription.Editors.Replace(necrologue, "");
                }
                articleData = articleData.Replace(jDiscription.Editors, "");
                articleData += necrologue;

                if (jDiscription.Editors.Contains(" и ") || jDiscription.Editors.Contains("; "))
                {
                    Match firstEditorM = patterns.MatchSplitEditor(jDiscription.Editors);
                    if (firstEditorM.Success)
                    {
                        foreach (string mPattern in patterns.matchPatterns)
                        {
                            string firstEditor = jDiscription.Editors.Split(new string[] { firstEditorM.Value }, StringSplitOptions.None)[0];
                            Match match = Regex.Match(firstEditor, mPattern);
                            if (match.Success)
                            {
                                jDiscription.FirstEdLastName = ExtractProp(articleData, mPattern);
                                if (!string.IsNullOrEmpty(jDiscription.LastName))
                                {
                                    articleData = articleData.Replace(jDiscription.LastName, "");
                                    if (!patterns.detectedPatterns.Contains(mPattern))
                                    {
                                        jDiscription.Initials = ExtractProp(jDiscription.LastName, patterns.initialsPattern);
                                        jDiscription.LastName = ReplaceIfNotNull(jDiscription.LastName, jDiscription.Initials);
                                    }

                                    if (jDiscription.LastName.Split(new[] { ',' }, 2).Length > 1)
                                    {
                                        jDiscription.Rank = jDiscription.LastName.Split(new[] { ',' }, 2)[1].Trim();
                                        jDiscription.LastName = ReplaceIfNotNull(jDiscription.LastName, jDiscription.Rank);
                                    }
                                    if (patterns.invertPatterns.Contains(mPattern))
                                    {
                                        jDiscription.Invertion = "1";
                                    }
                                }
                                jDiscription.LastName = CleanUpString(jDiscription.LastName);
                            }
                        }
                    }
                    jDiscription.Editors = CleanUpString(jDiscription.Editors.Replace(" / ", ""));
                }
            }
        }

        private void ExtractFullPubInfo(ref string journalData, ref string fullPubInfo)
        {
            if (journalData.Split(new string[] { "Полная публикация:" }, StringSplitOptions.None).Length > 1)
            {
                fullPubInfo = journalData.Split(new string[] { "Полная публикация:" }, StringSplitOptions.None)[1];
                journalData = journalData.Replace(fullPubInfo, "");
            }
        }

        private void GetTitleAndTitleInfo(JDiscription jDiscription, ref string articleData)
        {
            articleData = CleanUpString(articleData);
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
            foreach (string mPattern in patterns.matchPatterns)
            {
                Match match = Regex.Match(articleData, mPattern);
                if (match.Success)
                {
                    jDiscription.LastName = ExtractProp(articleData, mPattern);
                    if (!string.IsNullOrEmpty(jDiscription.LastName))
                    {
                        articleData = articleData.Replace(jDiscription.LastName, "");
                        ParseAutor(jDiscription, mPattern);
                    }
                    jDiscription.LastName = CleanUpString(jDiscription.LastName);
                }
            }
        }

        private void ParseAutor(JDiscription jDiscription, string mPattern)
        {
            if (!patterns.detectedPatterns.Contains(mPattern))
            {
                jDiscription.Initials = ExtractProp(jDiscription.LastName, patterns.initialsPattern);
                jDiscription.LastName = ReplaceIfNotNull(jDiscription.LastName, jDiscription.Initials);
            }

            if (jDiscription.LastName.Split(new[] { ',' }, 2).Length > 1)
            {
                jDiscription.Rank = jDiscription.LastName.Split(new[] { ',' }, 2)[1].Trim();
                jDiscription.LastName = ReplaceIfNotNull(jDiscription.LastName, jDiscription.Rank);
            }
            if (patterns.invertPatterns.Contains(mPattern))
            {
                jDiscription.Invertion = "1";
            }
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