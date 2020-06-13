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
        internal string cleanUpPattern = @"^\s*(\.|\,|\:|\;|\/)|(\.|\,|\:|\;|\/)\s*$";

        internal string oddPagesPattern = "^.+(Передняя обложка|Объявления|Задняя обложка|Последняя страница).+";

        internal string linePattern = @"^\d\. .+\/\/.+";

        internal string reviewPattern = @"\[Рец. на:.+$";

        internal string volumePattern = @"Т. \d+(\–\d+)?";

        internal string numberPattern = @"№ \d+\/?\d*\/?\d*";

        internal string pagesPattern = @"С\.\s(.*\(\d-([а-я|ё])+ пагин\.\)|\d+\–?\d*)";

        internal string notesPattern = @"\((Начало.|Продолжение.|Окончание.)\)";

        internal string initialsPattern = @"[А-Я]\.(\s[А-Я]\.)?";

        internal string necrologuePattern = @"\[Некролог\.?\]";

        internal string exclusionPattern = @"\[.*(по поводу кн.|рец. на|о книге|по поводу .+ст.).*\]";

        internal string editorsPattern = @"\s\/\s(Сообщ|Пер|Под ред|Вступ|Примеч|Публ|Предисл|С портр|Сост).+";

        internal string editorFunc = @"^\s?(Сообщ|Пер|Под ред|Вступ|Примеч|Публ|Предисл|С портр|Сост)\.?(\sс\s[а-я]*\.)?";

        internal string lastName = @"(\sи|;)\s";

        internal string rankPattern = @"\s[а-я]+\.?";

        internal string MatchInitials = @"[А-Я]\.(\s[А-Я]\.)?";

        internal string yearPattern = @"\d{4}(\–|\-)?\d{0,4}";

        internal string yearNumberPattern = @"^\d{4}(-|\/\d{4})*(-|\/)*\d*(\/\d)?$";

        internal Match MatchLine(string str)
        {
            return Regex.Match(str, @"^\d+\.\s?.+\/\/.+");
        }

        internal Match MatchOddPages(string str)
        {
            return Regex.Match(str, @".+(Передняя обложка|Объявления|Задняя обложка|Список сокращений|Последняя страница).+");
        }

        internal Match MatchSplitEditor(string str)
        {
            return Regex.Match(str, @"(\sи|;)\s");
        }

        internal MatchCollection EditorsCountPattern(string str)
        {
            return Regex.Matches(str, @"(\[?[А-Я]\.(\s[А-Я]\.)?\]?\s[А-Я][а-я|ё]+|(архим|иг|прот|свящ|иером)[а-я|ё]*\.?\s[А-Я][а-я|ё]+\s?\(?[А-Я][а-я|ё]+\)?)");
        }

        internal Match MatchUnknownPattern(string str)
        {
            return Regex.Match(str, @"^.*\[Автор не установлен.\]");
        }

        internal Match DetectedAutorPattern(string str)
        {
            return Regex.Match(str, @"^[А-Я]*[а-я|ё]* ?([А-Я]|\w)*\.? ?([А-Я]|\w|\*)\.? \[= ?[А-я|ё]*-?[А-Я][а-я|ё]+ [А-Я]\. ?[А-Я]?\.?\]");
        }

        internal Match DetectedMonachPattern(string str)
        {
            return Regex.Match(str, @"^[А-Я]*[а-я|ё]* ?([А-Я]|\w)\. ([А-Я]|\w)\. \*?\[= ?[А-я|ё]+ \([А-я|ё]+\), [а-я|ё]+\.\]");
        }

        internal Match HiddenManPattern(string str)
        {
            return Regex.Match(str, @"^\[([А-Я])([а-я|ё])+ ([А-Я]). ([А-Я]).\]");
        }

        internal Match MatchManPattern(string str)
        {
            return Regex.Match(str, @"^[А-я|ё]*-?[А-Я][а-я|ё]+\s[А-Я]\.(\s[А-Я]\.)?,?\s?(диак|свящ|прот|граф)?\.?\,?\s?(проф.)?");
        }

        internal Match MatchMonachPattern(string str)
        {
            return Regex.Match(str, @"^[А-я|ё]+\s\([А-я|ё]+\),\s[а-я|ё]+\.\,?\s?(наместник.+пустыни|наместник.+монастыря)?");
        }

        private Match ReversedMonachPattern(string str)
        {
            return Regex.Match(str, @"[а-я|ё]+\.?\s[А-Я][а-я|ё]+\s?\([А-Я][а-я|ё]+\)");
        }

        internal Match MatchBishopPattern(string str)
        {
            return Regex.Match(str, @"^[А-Я][а-я|ё]+\s\([А-Я][а-я|ё]+\),\s(еп|архиеп|митр|патр)[а-я|ё]*\.?\s[А-Я][а-я|ё]+(ий|ой)\s?и?\s?[А-Я]?[а-я|ё]*");
        }

        internal Match MatchSaintPattern(string str)
        {
            return Regex.Match(str, @"^([А-Я])([а-я|ё])+ ([А-Я])([а-я|ё])+, ([а-я|ё])+\.");
        }

        internal Match SaintBishopPattern(string str)
        {
            return Regex.Match(str, @"^([А-Я])([а-я|ё])+, ([а-я|ё])+\. ([А-Я])([а-я|ё])+ий, ([а-я|ё])+\.");
        }

        internal List<Match> InvertMathches(string str)
        {
            return new List<Match>() {
            MatchMonachPattern(str),
            ReversedMonachPattern(str),
            MatchBishopPattern(str),
            MatchSaintPattern(str),
            SaintBishopPattern(str)};
        }

        internal List<Match> AutorMatches(string str)
        {
            return new List<Match>() {
                MatchUnknownPattern(str),
                DetectedAutorPattern(str),
                DetectedMonachPattern(str),
                HiddenManPattern(str),
                MatchManPattern(str),
                MatchMonachPattern(str),
                MatchBishopPattern(str),
                MatchSaintPattern(str),
                SaintBishopPattern(str)
            };
        }

        public List<Match> DetectedMatches(string str)
        {
            return new List<Match>() {
                DetectedAutorPattern(str),
                DetectedMonachPattern(str)};
        }

        internal string DeclineLastName(string firstEdLastName)
        {
            Match monachLNameMatch = Regex.Match(firstEdLastName, @"\(.+\)");
            if (monachLNameMatch.Success)
            {
                string lastName = GetEdLastName(monachLNameMatch.Value);
                firstEdLastName = Regex.Replace(firstEdLastName, lastName, "");
                string firstName = GetFirstName(firstEdLastName.Trim());
                firstEdLastName = Regex.Replace(firstEdLastName, firstName, "");
                firstEdLastName = firstEdLastName + firstName + lastName;
            }
            return firstEdLastName;
        }

        private string GetFirstName(string firstName)
        {
            return firstName + ' ';
        }

        private static string GetEdLastName(string lastName)
        {
            lastName = Regex.Replace(lastName, @"ского\)$", "ский)");
            lastName = Regex.Replace(lastName, @"ова\)$", "ов)");
            lastName = Regex.Replace(lastName, @"ева\)$", "ев)");
            lastName = Regex.Replace(lastName, @"ына\)$", "ын)");
            lastName = Regex.Replace(lastName, @"ина\)$", "ин)");
            lastName = Regex.Replace(lastName, @"ыка\)$", "ык)");
            lastName = Regex.Replace(lastName, @"ика\)$", "ик)");
            return lastName;
        }
    }
}