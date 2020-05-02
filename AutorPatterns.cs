using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocJournalParser
{
    internal class AutorPatterns
    {
        internal string initialsPattern = @"[А-Я].\s[А-Я].";
        internal string unknownPattern = @"^.+\[Автор не установлен.\]";
        internal string detectedAutorPattern = @"^[А-Я]*[а-я]* ?([А-Я]|\w)*\.? ?([А-Я]|\w|\*)\.? \[= ?[А-я]*-?[А-Я][а-я]+ [А-Я]\. ?[А-Я]?\.?\]";
        internal string detectedMonachPattern = @"^[А-Я]*[а-я]* ?([А-Я]|\w)\. ([А-Я]|\w)\. \*?\[= ?[А-я]+ \([А-я]+\), [а-я]+\.\]";
        internal string hiddenManPattern = @"^\[([А-Я])([а-я])+ ([А-Я]). ([А-Я]).\]";
        internal string manPattern = @"^([А-я])*-?([А-Я])([а-я])+ ([А-Я])\. ([А-Я])\.,?( диак| свящ| прот| граф)?\.?( проф.)?";
        internal string monachPattern = @"^([А-я])+ \(([А-я])+\), ([а-я])+\.";
        internal string bishopPattern = @"^([А-Я])([а-я])+ \(([А-Я])([а-я])+\), ([а-я])+ ([А-Я])([а-я])+ий ?и? ?([А-Я])?([а-я])*\.?";
        internal string saintPattern = @"^([А-Я])([а-я])+ ([А-Я])([а-я])+, ([а-я])+\.";
        internal string saintBishopPattern = @"^([А-Я])([а-я])+, ([а-я])+\. ([А-Я])([а-я])+ий, ([а-я])+\.";
        internal string[] invertPatterns;
        internal string[] matchPatterns;
        internal string[] detectedPatterns;

        public AutorPatterns()
        {
            invertPatterns = new string[] {
                monachPattern,
                bishopPattern,
                saintPattern,
                saintBishopPattern
            };
            matchPatterns = new string[] {
                unknownPattern,
                detectedAutorPattern,
                detectedMonachPattern,
                hiddenManPattern,
                manPattern,
                monachPattern,
                bishopPattern,
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