using DocJournalParser;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DocJournalParserTEST
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMan()
        {
            Patterns autorPatterns = new Patterns();
            LineParser lineParser = new LineParser(autorPatterns);
            string line = @"Горский-Платонов В. И., прот., проф. Об употреблении печатного слова: " +
                "[Речь : Прения о вере. М., 1891] " +
                "// Богословский вестник 2005-2006. Т. 1. № 1. С. 115 (2-я пагин.). (Начало.) " +
                "Полная публикация: БВ 1892. Т. 1. № 3. С. 532–544 (2-я пагин.). ";
            JDiscription actualJD = lineParser.Parse(line);
            JDiscription expectedJD = new JDiscription();
            expectedJD.LastName = "Горский-Платонов";
            expectedJD.Initials = "В. И.";
            expectedJD.Rank = "прот., проф.";
            expectedJD.Title = "Об употреблении печатного слова";
            expectedJD.TitleInfo = "[Речь : Прения о вере. М., 1891]";
            expectedJD.Year = "2005-2006";
            expectedJD.Volume = "1";
            expectedJD.Number = "1";
            expectedJD.Pages = "115 (2-я пагин.)";
            expectedJD.Notes = "Начало";
            expectedJD.FullPubYear = "1892";
            expectedJD.FullPubVolume = "1";
            expectedJD.FullPubNumber = "3";
            expectedJD.FullPubPageRange = "532–544 (2-я пагин.)";

            TestEquals(actualJD, expectedJD);
        }

        [TestMethod]
        public void TestMonach()
        {
            Patterns autorPatterns = new Patterns();
            LineParser lineParser = new LineParser(autorPatterns);
            string line = @"Антоний (Храповицкий), архим. Отрадное явление: " +
                "[Рец. на:] Летопись историко-филологического общества при Новороссийском университете" +
                "// Богословский вестник 1892. Т. 3. № 10/11/12. С. I + I–VIII (4-я пагин.). (Начало.) ";
            JDiscription actualJD = lineParser.Parse(line);
            JDiscription expectedJD = new JDiscription();
            expectedJD.LastName = "Антоний (Храповицкий)";
            expectedJD.Rank = "архим.";
            expectedJD.Invertion = "1";
            expectedJD.Title = "Отрадное явление";
            expectedJD.TitleInfo = "[Рец. на:] Летопись историко-филологического общества при Новороссийском университете";
            expectedJD.Year = "1892";
            expectedJD.Volume = "3";
            expectedJD.Number = "10/11/12";
            expectedJD.Pages = "I + I–VIII (4-я пагин.)";
            expectedJD.Notes = "Начало";

            TestEquals(actualJD, expectedJD);
        }

        [TestMethod]
        public void TestSaint1()
        {
            Patterns autorPatterns = new Patterns();
            LineParser lineParser = new LineParser(autorPatterns);
            string line = @"Кирилл, архиеп. Александрийский, свт. Толкование на пророка Михея / Пер. и примеч. М. Д. Муретова" +
                "// Богословский вестник 1894. Т. 2. № 4. С. 131–162 (1-я пагин.). (Продолжение.) ";
            JDiscription actualJD = lineParser.Parse(line);
            JDiscription expectedJD = new JDiscription();
            expectedJD.LastName = "Кирилл";
            expectedJD.Rank = "архиеп. Александрийский, свт.";
            expectedJD.Invertion = "1";
            expectedJD.Title = "Толкование на пророка Михея / Пер. и примеч. М. Д. Муретова";
            expectedJD.Year = "1894";
            expectedJD.Volume = "2";
            expectedJD.Number = "4";
            expectedJD.Pages = "131–162 (1-я пагин.)";
            expectedJD.Notes = "Продолжение";

            TestEquals(actualJD, expectedJD);
        }

        [TestMethod]
        public void TestSaint2()
        {
            Patterns autorPatterns = new Patterns();
            LineParser lineParser = new LineParser(autorPatterns);
            string line = @"Астерий Амасийский, св. Толкование на пророка Михея [Рец. на: Die Nachtwache oder] " +
                "// Богословский вестник 1894. Т. 2. № 4. С. 131–162 (1-я пагин.). (Продолжение.) ";
            JDiscription actualJD = lineParser.Parse(line);
            JDiscription expectedJD = new JDiscription();
            expectedJD.LastName = "Астерий Амасийский";
            expectedJD.Rank = "св.";
            expectedJD.Invertion = "1";
            expectedJD.Title = "Толкование на пророка Михея";
            expectedJD.TitleInfo = "[Рец. на: Die Nachtwache oder]";
            expectedJD.Year = "1894";
            expectedJD.Volume = "2";
            expectedJD.Number = "4";
            expectedJD.Pages = "131–162 (1-я пагин.)";
            expectedJD.Notes = "Продолжение";

            TestEquals(actualJD, expectedJD);
        }

        [TestMethod]
        public void TestBishop()
        {
            Patterns autorPatterns = new Patterns();
            LineParser lineParser = new LineParser(autorPatterns);
            string line = @"Леонтий (Лебединский), митрополит Московский и Коломенский.[Рец. на:] Покровский Н. " +
                "Евангелие в памятниках иконографии, преимущественно византийских и русских. СПб., 1892" +
                "// Богословский вестник 2005–2006. Т. 5–6. № 5/6. С. 791";
            JDiscription actualJD = lineParser.Parse(line);
            JDiscription expectedJD = new JDiscription();
            expectedJD.LastName = "Леонтий (Лебединский)";
            expectedJD.Rank = "митрополит Московский и Коломенский";
            expectedJD.Invertion = "1";
            expectedJD.Title = "[Рец. на:] Покровский Н. Евангелие в памятниках иконографии, преимущественно византийских и русских. СПб., 1892";
            expectedJD.Year = "2005–2006";
            expectedJD.Volume = "5–6";
            expectedJD.Number = "5/6";
            expectedJD.Pages = "791";

            TestEquals(actualJD, expectedJD);
        }

        [TestMethod]
        public void TestUnknown()
        {
            Patterns autorPatterns = new Patterns();
            LineParser lineParser = new LineParser(autorPatterns);
            string line = @"11. [Автор не установлен.] [Рец. на:] Лебедев А., проф. " +
                @"Очерки истории Византийско-Восточной Церкви от конца XI до половины XV века. М., 1892" +
                @"// Богословский вестник 1892. Т. 1. № 2. С. 443–445 (2-я пагин.). ";
            JDiscription actualJD = lineParser.Parse(line);
            JDiscription expectedJD = new JDiscription();
            expectedJD.LastName = "[Автор не установлен.]";
            expectedJD.Title = "[Рец. на:] Лебедев А., проф. Очерки истории Византийско-Восточной Церкви от конца XI до половины XV века. М., 1892";
            expectedJD.Year = "1892";
            expectedJD.Volume = "1";
            expectedJD.Number = "2";
            expectedJD.Pages = "443–445 (2-я пагин.)";

            TestEquals(actualJD, expectedJD);
        }

        private void TestEquals(JDiscription actualJD, JDiscription expectedJD)
        {
            Assert.AreEqual(expectedJD.LastName, actualJD.LastName);
            Assert.AreEqual(expectedJD.Initials, actualJD.Initials);
            Assert.AreEqual(expectedJD.Rank, actualJD.Rank);
            Assert.AreEqual(expectedJD.Invertion, actualJD.Invertion);
            Assert.AreEqual(expectedJD.Title, actualJD.Title);
            Assert.AreEqual(expectedJD.TitleInfo, actualJD.TitleInfo);
            Assert.AreEqual(expectedJD.Year, actualJD.Year);
            Assert.AreEqual(expectedJD.Volume, actualJD.Volume);
            Assert.AreEqual(expectedJD.Number, actualJD.Number);
            Assert.AreEqual(expectedJD.Pages, actualJD.Pages);
            Assert.AreEqual(expectedJD.Notes, actualJD.Notes);
            Assert.AreEqual(expectedJD.FullPubYear, actualJD.FullPubYear);
            Assert.AreEqual(expectedJD.FullPubVolume, actualJD.FullPubVolume);
            Assert.AreEqual(expectedJD.FullPubNumber, actualJD.FullPubNumber);
            Assert.AreEqual(expectedJD.FullPubPageRange, actualJD.FullPubPageRange);
            Assert.AreEqual(expectedJD.LastName, actualJD.LastName);
        }
    }
}