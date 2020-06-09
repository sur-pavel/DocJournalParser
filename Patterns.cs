using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DocJournalParser
{
    public class Patterns

    {
        internal Match cleanUpPattern(string str)
        {
            return Regex.Match(str, @"^\s*(\.|\,|\:|\;)|(\.|\,|\:|\;)\s*$");
        }

        internal Match yearNumberPattern(string str)
        {
            return Regex.Match(str, @"^(\d{4})*\/?\d{4}-\d\/?\d?");
        }

        internal Match oddPagesPattern(string str)
        {
            return Regex.Match(str, @"^\d+\.\s(Передняя обложка|Объявления|Задняя обложка).+");
        }

        internal Match linePattern(string str)
        {
            return Regex.Match(str, @"^\d\. .+\/\/.+");
        }

        internal Match reviewPattern(string str)
        {
            return Regex.Match(str, @"\[Рец. на:.+$");
        }

        internal Match yearPattern(string str)
        {
            return Regex.Match(str, @"\d{4}(\–|\-)?\d{0,4}");
        }

        internal Match volumePattern(string str)
        {
            return Regex.Match(str, @"Т. \d+(\–\d+)?");
        }

        internal Match numberPattern(string str)
        {
            return Regex.Match(str, @"№ \d+\/?\d*\/?\d*");
        }

        internal Match pagesPattern(string str)
        {
            return Regex.Match(str, @"С\.\s(.*\(\d-([а-я|ё])+ пагин\.\)|\d+\–?\d*)");
        }

        internal Match notesPattern(string str)
        {
            return Regex.Match(str, @"\((Начало.|Продолжение.|Окончание.)\)");
        }

        internal Match initialsPattern(string str)
        {
            return Regex.Match(str, @"[А-Я]\.(\s[А-Я]\.)?");
        }

        internal Match necrologuePattern(string str)
        {
            return Regex.Match(str, @"\[Некролог\.?\]");
        }

        internal Match exclusionPattern(string str)
        {
            return Regex.Match(str, @"\[.*(по поводу кн.|рец. на|о книге|по поводу .+ст.).*\]");
        }

        internal Match editorsPattern(string str)
        {
            return Regex.Match(str, @"\s\/\s(Сообщ|Пер|Под|Вступ|Примеч|Предисл|С портр|Сост).+");
        }

        internal Match editorsCountPattern(string str)
        {
            return Regex.Match(str, @"(\[?[А-Я]\.(\s[А-Я]\.)?\]?\s[А-Я][а-я|ё]+|(архим|иг|прот|свящ|иером)[а-я|ё]*\.?\s[А-Я][а-я|ё]+\s?\(?[А-Я][а-я|ё]+\)?)\,?");
        }

        internal Match unknownPattern(string str)
        {
            return Regex.Match(str, @"^.*\[Автор не установлен.\]");
        }

        internal Match detectedAutorPattern(string str)
        {
            return Regex.Match(str, @"^[А-Я]*[а-я|ё]* ?([А-Я]|\w)*\.? ?([А-Я]|\w|\*)\.? \[= ?[А-я|ё]*-?[А-Я][а-я|ё]+ [А-Я]\. ?[А-Я]?\.?\]");
        }

        internal Match detectedMonachPattern(string str)
        {
            return Regex.Match(str, @"^[А-Я]*[а-я|ё]* ?([А-Я]|\w)\. ([А-Я]|\w)\. \*?\[= ?[А-я|ё]+ \([А-я|ё]+\), [а-я|ё]+\.\]");
        }

        internal Match hiddenManPattern(string str)
        {
            return Regex.Match(str, @"^\[([А-Я])([а-я|ё])+ ([А-Я]). ([А-Я]).\]");
        }

        internal Match manPattern(string str)
        {
            return Regex.Match(str, @"^[А-я|ё]*-?[А-Я][а-я|ё]+\s[А-Я]\.(\s[А-Я]\.)?,?\s?(диак|свящ|прот|граф)?\.?\,?\s?(проф.)?");
        }

        internal Match monachPattern(string str)
        {
            return Regex.Match(str, @"^[А-я|ё]+\s\([А-я|ё]+\),\s[а-я|ё]+\.\,?\s?(наместник.+пустыни|наместник.+монастыря)?");
        }

        internal Match bishopPattern(string str)
        {
            return Regex.Match(str, @"^[А-Я][а-я|ё]+\s\([А-Я][а-я|ё]+\),\s(еп|архиеп|митр|патр)[а-я|ё]*\.?\s[А-Я][а-я|ё]+(ий|ой)\s?и?\s?[А-Я]?[а-я|ё]*");
        }

        internal Match saintPattern(string str)
        {
            return Regex.Match(str, @"^([А-Я])([а-я|ё])+ ([А-Я])([а-я|ё])+, ([а-я|ё])+\.");
        }

        internal Match saintBishopPattern(string str)
        {
            return Regex.Match(str, @"^([А-Я])([а-я|ё])+, ([а-я|ё])+\. ([А-Я])([а-я|ё])+ий, ([а-я|ё])+\.");
        }

        internal List<Match> InvertMathches(string str)
        {
            Match monachMatch = monachPattern(str);
            Match bishopMatch = bishopPattern(str);
            Match saintMatch = saintPattern(str);
            Match saintBishopMatch = saintBishopPattern(str);
            List<Match> invertMathces = new List<Match>();
            invertMathces.Add(monachMatch);
            invertMathces.Add(bishopMatch);
            invertMathces.Add(saintMatch);
            invertMathces.Add(saintBishopMatch);
            return invertMathces;
        }

        internal List<Match> AutorMatches(string str)
        {
            Match unknownMatch = unknownPattern(str);
            Match detectedAutorMatch = detectedAutorPattern(str);
            Match detectedMonachMatch = detectedMonachPattern(str);
            Match hiddenManMatch = hiddenManPattern(str);
            Match manMatch = manPattern(str);
            Match monachMatch = monachPattern(str);
            Match bishopMatch = bishopPattern(str);
            Match saintMatch = saintPattern(str);
            Match saintBishopMatch = saintBishopPattern(str);

            List<Match> autorMatches = new List<Match>();
            autorMatches.Add(unknownMatch);
            autorMatches.Add(detectedAutorMatch);
            autorMatches.Add(detectedMonachMatch);
            autorMatches.Add(hiddenManMatch);
            autorMatches.Add(manMatch);
            autorMatches.Add(monachMatch);
            autorMatches.Add(bishopMatch);
            autorMatches.Add(saintMatch);
            autorMatches.Add(saintBishopMatch);
            return autorMatches;
        }

        public List<Match> DetectedMatches(string str)
        {
            Match detectedAutorMatch = detectedAutorPattern(str);
            Match detectedMonachMatch = detectedMonachPattern(str);
            List<Match> detectedMatches = new List<Match>();
            detectedMatches.Add(detectedAutorMatch);
            detectedMatches.Add(detectedMonachMatch);
            return detectedMatches;
        }

        internal Match MatchOddPages(string str)
        {
            return Regex.Match(str, @"^\d+\.\s?(Передняя обложка|Объявления|Задняя обложка|Список сокращений).+");
        }

        internal Match MatchYear(string str)
        {
            return Regex.Match(str, @"^(\d{4})*\/?\d{4}-\d\/?\d?");
        }

        internal Match MatchLastName(string str)
        {
            return Regex.Match(str, @"(\sи|;)\s");
        }

        internal Match MatchRank(string str)
        {
            return Regex.Match(str, @"(\sи|;)\s");
        }

        internal Match MatchInitials(string str)
        {
            return Regex.Match(str, @"[А-Я]\.(\s[А-Я]\.)?");
        }

        public Match MatchLine(string str)
        {
            return Regex.Match(str, @"^\d+\.\s?.+\/\/.+");
        }

        internal Match MatchSplitEditor(string str)
        {
            return Regex.Match(str, @"(\sи|;)\s");
        }
    }
}