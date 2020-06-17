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
        internal const string CleanUpPattern = @"^\s*(\.|\,|\:|\;|\/)|(\,|\:|\;|\/)\s*$";

        internal const string ReviewPattern = @"\[Рец. на:.+$";

        internal const string VolumePattern = @"Т.\s\d+(\–\d+)?";

        internal const string NumberPattern = @"№\s\d+\/?\d*\/?\d*";

        internal const string PagesPattern = @"С\.\s(.*\(\d-([а-я|ё])+\sпагин\.\)|\d+\–?\d*)";

        internal const string NotesPattern = @"\((Начало.|Продолжение.|Окончание.)\)";

        internal const string InitialsPattern = @"\[?[А-Я]\.(\s[А-Я]\.\]?)?";

        internal const string NecrologuePattern = @"\[Некролог\.?\]";

        internal const string ExclusionPattern = @"\[.*(по поводу кн.|рец. на|о книге|по поводу .+ст.).*\]";

        internal const string EditorsPattern = @"\s\/\s(Сообщ|Пер|Под ред|Вступ|Примеч|Публ|Предисл|С портр|Сост).+";

        internal string editorFunc =
            @"^\s(Сообщ|Пер|Под ред|Вступ|Примеч|Публ|Предисл|С портр|Сост)\.?(\sс\s[а-я]+\.?)?";

        internal string lastNamePattern = @"(\sи|;)\s";

        internal const string RankPattern = @"\s[а-я]{3,10}\.";

        internal const string YearPattern = @"\d{4}(\–|\-)?\d{0,4}";

        internal const string YearNumberPattern = @"^\d{4}(-|\/\d{4})*(-|\/)*\d*(\/\d)?$";
        
        private const string FullEditorName = @"[а-я|ё]+\.?\s[А-Я][а-я|ё]+(\s\(?[А-Я][а-я|ё]+\)?)?";
        private readonly Dictionary<string, string> _firstNameFinals;
        private readonly Dictionary<string, string> _lastNameFinals;

        public Patterns()
        {
            _firstNameFinals = new Dictionary<string, string>()
            {
                {"ия$", "ий"},
                {"ием$", "ий"},
                {"ида$", "ид"},
                {"идом$", "ид"},
                {"ита$", "ит"},
                {"итом$", "ит"},
                {"ана$", "ан"},
                {"аном$", "ан"},
                {"ила$", "ил"},
                {"илом$", "ил"},
                {"ника$", "ник"},
                {"ником$", "ник"},
                {"ея$", "ей"},
                {"еем$", "ей"},
                {"слава$", "слав"},
                {"славом$", "слав"},
                {"ра$", "р"},
                {"ром$", "р"},
                {"ры$", "ра"},
                {"рой$", "ра"},
            };
            _lastNameFinals = new Dictionary<string, string>()
            {
                {"ского$", "ский"},
                {"ским$", "ский"},
                {"ской$", "ская"},
                {"ова$", "ов"},
                {"овом$", "ов"},
                {"овой$", "ова"},
                {"ева$", "ев"},
                {"евым$", "ев"},
                {"евой$", "ева"},
                {"ына$", "ын"},
                {"ынымм$", "ын"},
                {"ина$", "ин"},
                {"иным$", "ин"},
                {"ыка$", "ык"},
                {"ыком$", "ык"},
                {"ихеса$", "ихес"},
                {"ихесом$", "ихес"},
                {"селя$", "сель"},
                {"селем$", "сель"},
                {"ика$", "ик"},
                {"иком$", "ик"},
                {"ича$", "ич"},
                {"ичом$", "ич"},
                {"юка$", "юк"},
                {"юком$", "юк"}
            };
        }

        internal Match MatchLine(string str)
        {
            return Regex.Match(str, @"^\d+\.\s?.+\/\/.+");
        }

        internal Match MatchOddPages(string str)
        {
            return Regex.Match(str,
                @".+(Передняя обложка|Объявления|Задняя обложка|Список сокращений|Последняя страница).+");
        }

        internal Match MatchSplitEditor(string str)
        {
            return Regex.Match(str, @"(\sи|;)\s");
        }

        internal MatchCollection EditorsCountPattern(string str)
        {
            return Regex.Matches(str,
                @"(\[?[А-Я]\.(\s[А-Я]\.)?\]?\s[А-Я][а-я|ё]+|(архим|иг|прот|свящ|иером)[а-я|ё]*\.?\s[А-Я][а-я|ё]+(\s\(?[А-Я][а-я|ё]+\)?)?)");
        }

        internal Match MatchUnknownPattern(string str)
        {
            return Regex.Match(str, @"^.*\[Автор не установлен.\]");
        }

        internal Match DetectedAutorPattern(string str)
        {
            return Regex.Match(str,
                @"^[А-Я]*[а-я|ё]*\s?([А-Я]|\w)*\.?\s?([А-Я]|\w|\*)\.?\s\[=\s?[А-я|ё]*-?[А-Я][а-я|ё]+\s[А-Я]\.\s?[А-Я]?\.?\]");
        }

        internal Match DetectedMonachPattern(string str)
        {
            return Regex.Match(str,
                @"^[А-Я]*[а-я|ё]*\s?([А-Я]|\w)\.\s([А-Я]|\w)\.\s\*?\[=\s?[А-я|ё]+\s\([А-я|ё]+\),\s[а-я|ё]+\.\]");
        }

        internal Match HiddenManPattern(string str)
        {
            return Regex.Match(str, @"^\[([А-Я])([а-я|ё])+\s([А-Я]).\s([А-Я]).\]");
        }

        internal Match MatchManPattern(string str)
        {
            return Regex.Match(str,
                @"^[А-я|ё]*-?[А-Я][а-я|ё]+\s[А-Я]\.(\s[А-Я]\.)?,?\s?(диак|свящ|прот|граф)?\.?\,?\s?(проф.)?");
        }

        internal Match MatchMonachPattern(string str)
        {
            return Regex.Match(str,
                @"^[А-я|ё]+\s\([А-я|ё]+\),\s[а-я|ё]+\.\,?\s?(наместник.+пустыни|наместник.+монастыря)?");
        }

        private Match ReversedMonachPattern(string str)
        {
            return Regex.Match(str, @"[а-я|ё]+\.?\s[А-Я][а-я|ё]+\s?\([А-Я][а-я|ё]+\)");
        }

        internal Match MatchBishopPattern(string str)
        {
            return Regex.Match(str,
                @"^[А-Я][а-я|ё]+\s\([А-Я][а-я|ё]+\),\s(еп|архиеп|митр|патр)[а-я|ё]*\.?\s[А-Я][а-я|ё]+(ий|ой)\s?и?\s?[А-Я]?[а-я|ё]*");
        }

        internal Match MatchSaintPattern(string str)
        {
            return Regex.Match(str, @"^([А-Я])([а-я|ё])+\s([А-Я])([а-я|ё])+,\s([а-я|ё])+\.");
        }

        internal Match SaintBishopPattern(string str)
        {
            return Regex.Match(str, @"^([А-Я])([а-я|ё])+,\s([а-я|ё])+\.\s([А-Я])([а-я|ё])+ий,\s([а-я|ё])+\.");
        }

        internal List<Match> InvertMathches(string str)
        {
            return new List<Match>()
            {
                MatchMonachPattern(str),
                ReversedMonachPattern(str),
                MatchBishopPattern(str),
                MatchSaintPattern(str),
                SaintBishopPattern(str)
            };
        }

        internal List<Match> AutorMatches(string str)
        {
            return new List<Match>()
            {
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
            return new List<Match>()
            {
                DetectedAutorPattern(str),
                DetectedMonachPattern(str)
            };
        }

        internal string DeclineEditorNames(string firstEditor)
        {
            Match fullNameMatch = Regex.Match(firstEditor, FullEditorName);
            if (fullNameMatch.Success)
            {
                string lastName = string.Empty;
                string[] words = firstEditor.Trim().Split(' ');
                string firstName = words[1];
                firstEditor = firstEditor.Replace(firstName, "");
                if (words.Length > 2)
                {
                    lastName = words[2];
                    firstEditor = firstEditor.Replace(lastName, "");
                    lastName = " " + DeclineEdLastName(lastName);
                }

                firstName = DeclineFirstName(firstName.Trim());
                firstEditor = Regex.Replace(firstEditor, @"\s\s", " ") + firstName + lastName;
            }
            else
            {
                firstEditor = DeclineEdLastName(firstEditor);
            }

            return firstEditor;
        }

        private string DeclineFirstName(string firstName)
        {
            foreach (var final in _firstNameFinals)
            {
                firstName = Regex.Replace(firstName, final.Key, final.Value);
            }

            return firstName;
        }

        private string DeclineEdLastName(string lastName)
        {
            string bracket = "";
            if (lastName.Contains(")"))
            {
                bracket = ")";
                lastName = lastName.Replace(")", "").Trim();
            }

            foreach (var final in _lastNameFinals)
            {
                lastName = Regex.Replace(lastName, final.Key, final.Value);
            }

            return lastName + bracket;
        }
    }
}