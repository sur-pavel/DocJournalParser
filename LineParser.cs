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

                jDiscription.Year = ExtractProp(journalData, patterns.yearPattern(journalData).Value);
                jDiscription.JVolume = ExtractProp(journalData, patterns.volumePattern(journalData).Value, "Т. ");
                jDiscription.JNumber = ExtractProp(journalData, patterns.numberPattern(journalData).Value, "№ ");
                jDiscription.Pages = GetPages(journalData);
                jDiscription.Notes = ExtractProp(journalData, patterns.notesPattern(journalData).Value, "(", ")", ".");

                jDiscription.FullPubYear = ExtractProp(fullPubInfo, patterns.yearPattern(fullPubInfo).Value, " ", ".");
                jDiscription.FullPubVolume = ExtractProp(fullPubInfo, patterns.volumePattern(fullPubInfo).Value, "Т. ");
                jDiscription.FullPubNumber = ExtractProp(fullPubInfo, patterns.numberPattern(fullPubInfo).Value, "№ ");
                jDiscription.FullPubPageRange = ExtractProp(fullPubInfo, patterns.pagesPattern(fullPubInfo).Value, "С. ");
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
            Match editorsMatch = Regex.Match(articleData, patterns.editorsPattern(articleData).Value);
            string necrologue = Regex.Match(articleData, patterns.necrologuePattern(articleData).Value);
            if (editorsMatch.Success && !Regex.IsMatch(articleData, patterns.exclusionPattern(articleData).Value))
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
                        foreach (Match mPattern in patterns.AutorMatches(firstEditorM.Value))
                        {
                            string firstEditor = jDiscription.Editors.Split(new string[] { firstEditorM.Value }, StringSplitOptions.None)[0];
                            Match match = Regex.Match(firstEditor, mPattern.Value);
                            if (match.Success)
                            {
                                jDiscription.FirstEdLastName = ExtractProp(articleData, mPattern.Value);
                                if (!string.IsNullOrEmpty(jDiscription.LastName))
                                {
                                    articleData = articleData.Replace(jDiscription.LastName, "");
                                    foreach (Match detected in DetectedMatches(mPattern.Value))
                                    {
                                    }
                                    if (!patterns.DetectedMatches.Contains(mPattern))
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
            Match reviewMatch = Regex.Match(articleData, patterns.reviewPattern(articleData).Value);
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
            string pagesPattern = patterns.pagesPattern(journalData).Value;
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
            return Regex.Replace(str, patterns.cleanUpPattern(str).Value, "").Trim();
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