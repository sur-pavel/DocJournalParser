using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocJournalParser
{
    public class Patterns

    {
        internal string yearNumberPattern = @"^(\d{4})*\/?\d{4}-\d\/?\d?";
        internal string oddPagesPattern = @"^\d+\.\s(Передняя обложка|Объявления|Задняя обложка).+";
        internal string linePattern = @"^\d\. .+\/\/.+";
        internal string reviewPattern = @"\[Рец. на:.+$";
        internal string yearPattern = @"\d{4}(\–|\-)?\d{0,4}";
        internal string volumePattern = @"Т. \d+(\–\d+)?";
        internal string numberPattern = @"№ \d+\/?\d*\/?\d*";
        internal string pagesPattern = @"С\.\s(.*\(\d-([а-я])+ пагин\.\)|\d+\–?\d*)";
        internal string notesPattern = @"\((Начало.|Продолжение.|Окончание.)\)";
        internal string initialsPattern = @"[А-Я].\s[А-Я].";

        internal string unknownPattern = @"^.*\[Автор не установлен.\]";
        internal string detectedAutorPattern = @"^[А-Я]*[а-я]* ?([А-Я]|\w)*\.? ?([А-Я]|\w|\*)\.? \[= ?[А-я]*-?[А-Я][а-я]+ [А-Я]\. ?[А-Я]?\.?\]";
        internal string detectedMonachPattern = @"^[А-Я]*[а-я]* ?([А-Я]|\w)\. ([А-Я]|\w)\. \*?\[= ?[А-я]+ \([А-я]+\), [а-я]+\.\]";
        internal string hiddenManPattern = @"^\[([А-Я])([а-я])+ ([А-Я]). ([А-Я]).\]";
        internal string manPattern = @"^([А-я])*-?([А-Я])([а-я])+ ([А-Я])\. ([А-Я])\.,?( диак| свящ| прот| граф)?\.?\,?( проф.)?";
        internal string monachPattern = @"^[А-я]+\s\([А-я]+\),\s[а-я]+\.\,?\s?(наместник.+пустыни|наместник.+монастыря)?";
        internal string bishopPattern = @"^[А-Я][а-я]+\s\([А-Я][а-я]+\),\s(еп|архиеп|митр|патр)[а-я]*\.?\s[А-Я][а-я]+(ий|ой)\s?и?\s?[А-Я]?[а-я]*";
        internal string saintPattern = @"^([А-Я])([а-я])+ ([А-Я])([а-я])+, ([а-я])+\.";
        internal string saintBishopPattern = @"^([А-Я])([а-я])+, ([а-я])+\. ([А-Я])([а-я])+ий, ([а-я])+\.";
        internal string[] invertPatterns;
        internal string[] matchPatterns;
        internal string[] detectedPatterns;

        public Patterns()
        {
            invertPatterns = new string[] {
                bishopPattern,
                monachPattern,
                saintPattern,
                saintBishopPattern
            };
            matchPatterns = new string[] {
                unknownPattern,
                detectedAutorPattern,
                detectedMonachPattern,
                hiddenManPattern,
                manPattern,
                bishopPattern,
                monachPattern,
                saintPattern,
                saintBishopPattern
            };

            detectedPatterns = new string[]
            {
                detectedAutorPattern,
                detectedMonachPattern,
            };
        }
    }
}