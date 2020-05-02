using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace DocJournalParser
{
    internal class LineParser
    {
        public LineParser()
        {
        }

        internal JDiscription Parse(string line)
        {
            line = Regex.Replace(line, @"^\d+. ", "");
            line = line.Replace("//", " //").Replace("  ", " ");

            JDiscription jDiscription = new JDiscription();
            try
            {
                string articleData = line.Split(new string[] { " // " }, StringSplitOptions.None)[0];
                ParseAutor(jDiscription, articleData);

                GetTitleAndTitleInfo(jDiscription, articleData);
                if (Regex.IsMatch(jDiscription.Title, @"^[^А-Я|\[|\<|\d|\w]"))
                {
                    throw new InvalidCastException("WRONG START OF TITLE");
                }

                string journalData = line.Split(new string[] { " // " }, StringSplitOptions.None)[1];
                string fullPubInfo = string.Empty;
                if (journalData.Split(new string[] { "Полная публикация: " }, StringSplitOptions.None).Length > 1)
                {
                    fullPubInfo = journalData.Split(new string[] { "Полная публикация: " }, StringSplitOptions.None)[1];
                    journalData = journalData.Replace(fullPubInfo, "");
                }

                jDiscription.Year = ExtractProp(journalData, @" \d{4}.", " ", ".");
                jDiscription.Volume = ExtractProp(journalData, @"Т. \d", "Т. ");
                jDiscription.Number = ExtractProp(journalData, @"№ \d", "№ ");
                jDiscription.Pages = GetPages(journalData);

                jDiscription.FullPubYear = ExtractProp(fullPubInfo, @" \d{4}.", " ", ".");
                jDiscription.FullPubVolume = ExtractProp(fullPubInfo, @"Т. \d", "Т. ");
                jDiscription.FullPubNumber = ExtractProp(fullPubInfo, @"№ \d", "№ ");
                jDiscription.FullPubPageRange = ExtractProp(fullPubInfo, @"С. \d+–\d+ \(\d-([а-я])+ пагин.\)", "С. ");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.WriteLine(line);
                Console.WriteLine(jDiscription.ToString() + "\n");
            }
            return jDiscription;
        }

        private static void GetTitleAndTitleInfo(JDiscription jDiscription, string articleData)
        {
            articleData = Regex.Replace(articleData, @"^\.? ", "");
            if (articleData.StartsWith("["))
            {
                jDiscription.Title = articleData;
            }
            else if (articleData.Split(new string[] { ". [" }, StringSplitOptions.None).Length > 1)
            {
                jDiscription.Title = articleData.Split(new string[] { ". [" }, StringSplitOptions.None)[0];
                jDiscription.TitleInfo = "[" + articleData.Split(new string[] { ". [" }, StringSplitOptions.None)[1];
            }
            else if (articleData.Split(new[] { ':' }, 2).Length > 1)
            {
                jDiscription.Title = articleData.Split(new[] { ':' }, 2)[0];
                jDiscription.TitleInfo = articleData.Split(new[] { ':' }, 2)[1];
                jDiscription.TitleInfo = Regex.Replace(jDiscription.TitleInfo, @"^\.? ", "");
            }
            else
            {
                jDiscription.Title = articleData;
            }
        }

        private string GetPages(string journalData)
        {
            string pagesPattern = @"С. (\d+|(I{0,3}|XC|XL|L?X{0,3}))–(\d+|(I{0,3}|XC|XL|L?X{0,3})) " +
                @"\(\d-([а-я])+ пагин.\)\.? ?\(?(Начало.|Продолжение.|Окончание.)?\)?";
            string pages = ExtractProp(journalData, pagesPattern, "С. ");

            if (string.IsNullOrEmpty(pages))
            {
                pages = ExtractProp(journalData, @"\d+ с\.");
            }
            pages = Regex.Replace(pages, @"\.? $", "");
            pages = Regex.Replace(pages, @"\)\. ?$", ")");

            return pages;
        }

        private void ParseAutor(JDiscription jDiscription, string articleData)
        {
            string initialsPattern = @"[А-Я].\s[А-Я].";
            string unknownPattern = @"^([А-Я]|\w)*\.? ?([А-Я]|\w)*\.? ?\[Автор не установлен.\]";
            string knownPattern = @"^[А-Я]*[а-я]* ?([А-Я]|\w)*\.? ?([А-Я]|\w|\*)\.? \[= ?[А-я]*-?[А-Я][а-я]+ [А-Я]\. ?[А-Я]?\.?\]";
            string knownMonachPattern = @"^[А-Я]*[а-я]* ?([А-Я]|\w)\. ([А-Я]|\w)\. \*?\[= ?[А-я]+ \([А-я]+\), [а-я]+\.\]";
            string hiddenManPattern = @"^\[([А-Я])([а-я])+ ([А-Я]). ([А-Я]).\]";
            string manPattern = @"^([А-я])*-?([А-Я])([а-я])+ ([А-Я])\. ([А-Я])\.,?( диак| свящ| прот| граф)?\.?( проф.)?";
            string monachPattern = @"^([А-я])+ \(([А-я])+\), ([а-я])+\.";
            string bishopPattern = @"^([А-Я])([а-я])+ \(([А-Я])([а-я])+\), ([а-я])+ ([А-Я])([а-я])+ий ?и? ?([А-Я])?([а-я])*\.?";
            string saintPattern = @"^([А-Я])([а-я])+ ([А-Я])([а-я])+, ([а-я])+\.";
            string saintBishopPattern = @"^([А-Я])([а-я])+, ([а-я])+\. ([А-Я])([а-я])+ий, ([а-я])+\.";
            string[] invertPatterns = new string[] {
                monachPattern,
                bishopPattern,
                saintPattern,
                saintBishopPattern
            };

            string[] mPatterns = new string[] {
                unknownPattern,
                knownPattern,
                knownMonachPattern,
                hiddenManPattern,
                manPattern,
                monachPattern,
                bishopPattern,
                saintPattern,
                saintBishopPattern
            };
            foreach (string mPattern in mPatterns)
            {
                Match match = Regex.Match(articleData, mPattern);
                if (match.Success)
                {
                    jDiscription.LastName = ExtractProp(articleData, mPattern);
                    if (!string.IsNullOrEmpty(jDiscription.LastName))
                    {
                        articleData = articleData.Replace(jDiscription.LastName, "");
                        jDiscription.Initials = ExtractProp(jDiscription.LastName, initialsPattern);
                        jDiscription.LastName = jDiscription.LastName.Replace(jDiscription.Initials, "");

                        if (jDiscription.LastName.Split(new[] { ',' }, 2).Length > 1)
                        {
                            jDiscription.Rank = jDiscription.LastName.Split(new[] { ',' }, 2)[1];
                        }
                        if (invertPatterns.Contains(mPattern))
                        {
                            jDiscription.Invertion = "1";
                        }
                        jDiscription.LastName = Regex.Replace(jDiscription.LastName, @"\.\s?$", "");
                        jDiscription.LastName = Regex.Replace(jDiscription.LastName, @"\s$", "");
                    }
                }
            }
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
            return returnValue;
        }
    }
}