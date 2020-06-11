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

        internal string editorsPattern = @"\s\/\s(Сообщ|Пер|Под|Вступ|Примеч|Публ|Предисл|С портр|Сост).+";

        internal string editorFunc = @"\s?(Сообщ|Пер|Под|Вступ|Примеч|Публ|Предисл|С портр|Сост)\.?\s?";

        internal string lastName = @"(\sи|;)\s";

        internal string rank = @"(\sи|;)\s";

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
            return new List<Match>() {
            monachPattern(str),
            bishopPattern(str),
            saintPattern(str),
            saintBishopPattern(str)};
        }

        internal List<Match> AutorMatches(string str)
        {
            return new List<Match>() {
                unknownPattern(str),
                detectedAutorPattern(str),
                detectedMonachPattern(str),
                hiddenManPattern(str),
                manPattern(str),
                monachPattern(str),
                bishopPattern(str),
                saintPattern(str),
                saintBishopPattern(str)
            };
        }

        public List<Match> DetectedMatches(string str)
        {
            return new List<Match>() {
                detectedAutorPattern(str),
                detectedMonachPattern(str)};
        }
    }
}