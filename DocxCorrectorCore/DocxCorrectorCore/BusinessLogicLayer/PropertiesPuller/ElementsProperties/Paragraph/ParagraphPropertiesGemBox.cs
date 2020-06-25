using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.PropertiesPuller
{
    public class ParagraphPropertiesGemBox
    {
        public int ID { get; }
        public virtual string Content { get; }
        public int SpecialSymbolsCount { get; }
        public int WordsCount { get; }
        public int SymbolCount { get; }
        public bool Lowercase { get; }
        public bool Uppercase { get; }
        public string? LastSymbolPd { get; }
        public string? FirstKeyWord { get; }
        public string? PrevElementMark { get; }
        public string? CurElementMark { get; }
        public string? NextElementMark { get; }
        public string FullBold { get; }
        public string FullItalic { get; }
        public string Alignment { get; }
        public string KeepLinesTogether { get; }
        public string KeepWithNext { get; }
        public string LeftIndentation { get; }
        public string LineSpacing { get; }
        public string NoSpaceBetweenParagraphsOfSameStyle { get; }
        public string OutlineLevel { get; }
        public string PageBreakBefore { get; }
        public string RightIndentation { get; }
        public string SpaceAfter { get; }
        public string SpaceBefore { get; }
        public string SpecialIndentation { get; }

        // Private
        private int CountSpecialSymbols(Word.Paragraph paragraph)
        {
            List<string> specialSymbols = new List<string> { "=", ".", "–", "/", ",", ":", "?", "'", "[", "]", "(", ")", "-", "…", "!", "«", "»", ";", "\"" };
            string paragraphContent = GemBoxHelper.GetParagraphContentWithoutNewLine(paragraph);

            int result = 0;

            foreach (string symbol in specialSymbols)
            {
                if (paragraphContent.Contains(symbol)) { result++; }
            }
            return result;
        }

        private int CountWords(Word.Paragraph paragraph)
        {
            string paragraphContent = GemBoxHelper.GetParagraphContentWithoutNewLine(paragraph);
            var words = paragraphContent.Split(' ');
            return words.Count();
        }

        private int CountSymbols(Word.Paragraph paragraph)
        {
            string paragraphContent = GemBoxHelper.GetParagraphContentWithoutNewLine(paragraph);
            return paragraphContent.Count();
        }

        private bool CheckLower(Word.Paragraph paragraph)
        {
            string paragraphContent = GemBoxHelper.GetParagraphContentWithoutNewLine(paragraph);
            
            foreach (var symbol in paragraphContent)
            {
                if (char.IsLetter(symbol) && !char.IsUpper(symbol))
                {
                    return false;
                }
            }

            return true;
        }

        private bool CheckUpper(Word.Paragraph paragraph)
        {
            string paragraphContent = GemBoxHelper.GetParagraphContentWithoutNewLine(paragraph);

            foreach (var symbol in paragraphContent)
            {
                if (char.IsLetter(symbol) && !char.IsLower(symbol))
                {
                    return false;
                }
            }

            return true;
        }

        private string? CheckLastSymbol(Word.Paragraph paragraph)
        {
            List<string> keySymbols = new List<string> { ".", ",", ":", ";"};
            string lastSymbol;
            try
            {
                lastSymbol = paragraph.Content.ToString().Trim().Last().ToString();
            }
            catch
            {
                return null;
            }

            foreach (var keySymbol in keySymbols)
            {
                if (keySymbol == lastSymbol) { return keySymbol; }
            }
            return null;
        }

        private string? CheckFirstKeyWord(Word.Paragraph paragraph)
        {
            string paragraphContent = GemBoxHelper.GetParagraphContentWithoutNewLine(paragraph);
            var words = paragraphContent.Split(' ');

            string firstWord;
            try
            {
                firstWord = words[0];
            }
            catch
            {
                return null;
            }

            //Рисунок
            string[] pictureKeys = new string[] { "рисунок", "рис." };
            foreach (var key in pictureKeys)
            {
                if (firstWord.IndexOf(key, StringComparison.OrdinalIgnoreCase) != -1) { return "Рисунок"; }
            }

            //Таблица
            string[] tableKeys = new string[] { "таблица", "табл." };
            foreach (var key in tableKeys)
            {
                if (firstWord.IndexOf(key, StringComparison.OrdinalIgnoreCase) != -1) { return "Таблица"; }
            }

            //Продолжение
            string[] continuationKeys = new string[] { "продолжение" };
            foreach (var key in continuationKeys)
            {
                if (firstWord.IndexOf(key, StringComparison.OrdinalIgnoreCase) != -1) { return "Продолжение"; }
            }

            //Окончание
            string[] endingKeys = new string[] { "окончание", "окон." };
            foreach (var key in endingKeys)
            {
                if (firstWord.IndexOf(key, StringComparison.OrdinalIgnoreCase) != -1) { return "Окончание"; }
            }

            //Число без точки (1)
            Regex title1Regex = new Regex(@"^\d+$");
            if (title1Regex.IsMatch(firstWord)) { return "TitleLevel1"; }

            //Число-точка-число (1.2)
            Regex title2Regex = new Regex(@"^\d+\.\d+$");
            if (title2Regex.IsMatch(firstWord)) { return "TitleLevel2"; }

            //Число-точка-число-точка-число (1.2.3)
            Regex title3Regex = new Regex(@"^\d+\.\d+\.\d+$");
            if (title3Regex.IsMatch(firstWord)) { return "TitleLevel3"; }

            //Число-точка-число-точка-число-точка-число (1.2.3.4)
            Regex title4Regex = new Regex(@"^\d+\.\d+\.\d+$");
            if (title4Regex.IsMatch(firstWord)) { return "TitleLevel4"; }

            //Дефис или тире
            Regex hyphenRegex = new Regex(@"^[-–]$");
            if (hyphenRegex.IsMatch(firstWord)) { return "listLevel1"; }

            //Буква-закрывающая круглая скобка
            Regex letterBracketRegex = new Regex(@"^[a-яА-яa-zA-z]\)$");
            if (letterBracketRegex.IsMatch(firstWord)) { return "listLevel1"; }

            //Сочетание цифр и точек с окончанием на запятую или точку с запятой
            if (paragraphContent.Count() > 3)
            {
                for (int letterIndex = 0; letterIndex < firstWord.Count() - 1; letterIndex++)
                {
                    char letter = firstWord[letterIndex];
                    if ((!Char.IsDigit(letter)) & (letter != '.'))
                    {
                        return null;
                    }
                }
                if ((paragraphContent.Last() == ',') | (paragraphContent.Last() == ';')) { return "listLevel1"; }
            }

            return null;
        }

        // Public
        public ParagraphPropertiesGemBox(int id, Word.Paragraph paragraph)
        {
            ID = id;
            Content = GemBoxHelper.GetParagraphContentWithoutNewLine(paragraph);
            SpecialSymbolsCount = CountSpecialSymbols(paragraph);
            WordsCount = CountWords(paragraph);
            SymbolCount = CountSymbols(paragraph);
            Lowercase = CheckLower(paragraph);
            Uppercase = CheckUpper(paragraph);
            LastSymbolPd = CheckLastSymbol(paragraph);
            FirstKeyWord = CheckFirstKeyWord(paragraph);
            FullBold = paragraph.CharacterFormatForParagraphMark.Bold.ToString();
            FullItalic = paragraph.CharacterFormatForParagraphMark.Italic.ToString();
            Alignment = paragraph.ParagraphFormat.Alignment.ToString();
            KeepLinesTogether = paragraph.ParagraphFormat.KeepLinesTogether.ToString();
            KeepWithNext = paragraph.ParagraphFormat.KeepWithNext.ToString();
            LeftIndentation = paragraph.ParagraphFormat.LeftIndentation.ToString();
            LineSpacing = paragraph.ParagraphFormat.LineSpacing.ToString();
            NoSpaceBetweenParagraphsOfSameStyle = paragraph.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle.ToString();
            OutlineLevel = paragraph.ParagraphFormat.OutlineLevel.ToString();
            PageBreakBefore = paragraph.ParagraphFormat.PageBreakBefore.ToString();
            RightIndentation = paragraph.ParagraphFormat.RightIndentation.ToString();
            SpaceAfter = paragraph.ParagraphFormat.SpaceAfter.ToString();
            SpaceBefore = paragraph.ParagraphFormat.SpaceBefore.ToString();
            SpecialIndentation = paragraph.ParagraphFormat.SpecialIndentation.ToString();
        }

        public ParagraphPropertiesGemBox(int id, string content)
        {
            ID = id;
            Content = content;
        }
    }
}
