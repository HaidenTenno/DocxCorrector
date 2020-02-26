using System;
using DocxCorrector.Services;
using Word = Microsoft.Office.Interop.Word;

namespace DocxCorrector.Models
{
    // Выравнивание
    public enum NormalizedAligment: int
    {
        Left,
        Center,
        Right,
        Justify,
        Other
    }

    // TODO: - Naming
    // Для проверки на наличия слов, оформленных жирным, курсивом и т. д.
    public enum ContainsStatus: int
    {
        None, // Параграф не содержит таких слов
        Contains, // Параграф содержит слова с таким оформлением
        Full // Параграф оформлен полностью таким спосом
    }

    // Правило междустрочного интервала
    public enum LineSpacingRuleVariations: int
    {
        Single,
        OneAndHalf,
        Double,
        Miltiply,
        Other
    }


    public class NormalizedProperties
    {
        // ID параграфа
        public int Id { get; set; }
        // Отступ красной строки
        public float FirstLineIndent { get; set; }
        // Выравнивание
        public int Aligment { get; set; }
        // Количество символов
        public int SymbolsCount { get; set; }
        // Префикс - число
        public int PrefixIsNumber { get; set; }
        // Префикс - строчная буква 
        public int PrefixIsLowercase { get; set; }
        // Префикс - прописная буква
        public int PrefixIsUppercase { get; set; }
        // Префикс - ["-", "־", "᠆", "‐", "‑", "‒", "–", "—", "―", "﹘", "﹣", "－"]
        public int PrefixIsDash { get; set; }
        // Суффикс - [".","!","?"]
        public int SuffixIsEndSign { get; set; }
        // Суффикс - двоеточие
        public int SuffixIsColon { get; set; }
        // Суффикс - запятая или точка с запятой
        public int SuffixIsCommaOrSemicolon { get; set; }
        // Содержит - ["-", "־", "᠆", "‐", "‑", "‒", "–", "—", "―", "﹘", "﹣", "－"]
        public int ContainsDash { get; set; }
        // Содержит ")"
        public int ContainsBracket { get; set; }
        // Размер шрифта
        public float FontSize { get; set; }
        // Междустрочный интервал
        public float LineSpacing { get; set; }
        // Правило междустрочного интервала
        public int LineSpacingRule { get; set; }
        // Курсив
        public int Italic { get; set; }
        // Жирный
        public int Bold { get; set; }
        // Цвет черный
        public int BlackColor { get; set; }

        public NormalizedProperties(Word.Paragraph paragraph, int paragraphId)
        {
            int id = paragraphId;
            float firstLineIndent = paragraph.FirstLineIndent;
            NormalizedAligment aligment;
            switch (paragraph.Alignment)
            {
                case Word.WdParagraphAlignment.wdAlignParagraphLeft:
                    aligment = NormalizedAligment.Left;
                    break;
                case Word.WdParagraphAlignment.wdAlignParagraphCenter:
                    aligment = NormalizedAligment.Center;
                    break;
                case Word.WdParagraphAlignment.wdAlignParagraphRight:
                    aligment = NormalizedAligment.Right;
                    break;
                case Word.WdParagraphAlignment.wdAlignParagraphJustify:
                    aligment = NormalizedAligment.Justify;
                    break;
                default:
                    aligment = NormalizedAligment.Other;
                    break;
            }
            int prefixIsNumber = Char.IsDigit(paragraph.Range.Text[0]) ? 1 : 0;
            int prefixIsLowercase = Char.IsLower(paragraph.Range.Text[0]) ? 1 : 0;
            int prefixIsUppercase = Char.IsUpper(paragraph.Range.Text[0]) ? 1 : 0;
            string[] dashes = new string[] { "-", "־", "᠆", "‐", "‑", "‒", "–", "—", "―", "﹘", "﹣", "－" };
            int prefixIsDash = InteropHelper.CheckIfFirstSymbolOfParagraphIs(paragraph, dashes);
            string[] endSigns = new string[] { ".", "!", "?" };
            int suffixIsEndSign = InteropHelper.CheckIfLastSymbolOfParagraphIs(paragraph, endSigns);
            string[] colon = new string[] { ":" };
            int suffixIsColon = InteropHelper.CheckIfLastSymbolOfParagraphIs(paragraph, colon);
            string[] commaAndSemicolon = new string[] { ",", ";" };
            int suffixIsCommaOrSemicolon = InteropHelper.CheckIfLastSymbolOfParagraphIs(paragraph, commaAndSemicolon);
            int containsDash = InteropHelper.CheckIfParagraphsContainsOneOf(paragraph, dashes);
            string[] bracket = new string[] { ")" };
            int containsBracket = InteropHelper.CheckIfParagraphsContainsOneOf(paragraph, bracket);
            float fontSize = paragraph.Range.Font.Size;
            float lineSpacing = paragraph.LineSpacing;
            LineSpacingRuleVariations lineSpacingRule;
            switch (paragraph.LineSpacingRule)
            {
                case Word.WdLineSpacing.wdLineSpaceSingle:
                    lineSpacingRule = LineSpacingRuleVariations.Single;
                    break;
                case Word.WdLineSpacing.wdLineSpace1pt5:
                    lineSpacingRule = LineSpacingRuleVariations.OneAndHalf;
                    break;
                case Word.WdLineSpacing.wdLineSpaceDouble:
                    lineSpacingRule = LineSpacingRuleVariations.Double;
                    break;
                case Word.WdLineSpacing.wdLineSpaceMultiple:
                    lineSpacingRule = LineSpacingRuleVariations.Miltiply;
                    break;
                default:
                    lineSpacingRule = LineSpacingRuleVariations.Other;
                    break;
            }
            ContainsStatus italic;
            switch (paragraph.Range.Italic)
            {
                case -1:
                    italic = ContainsStatus.Full;
                    break;
                case 0:
                    italic = ContainsStatus.None;
                    break;
                default:
                    italic = ContainsStatus.Contains;
                    break;
            }
            ContainsStatus bold;
            switch (paragraph.Range.Bold)
            {
                case -1:
                    bold = ContainsStatus.Full;
                    break;
                case 0:
                    bold = ContainsStatus.None;
                    break;
                default:
                    bold = ContainsStatus.Contains;
                    break;
            }
            int blackColor = (paragraph.Range.Font.Color == Word.WdColor.wdColorBlack) || (paragraph.Range.Font.Color == Word.WdColor.wdColorAutomatic) ? 1 : 0;

            Id = id;
            FirstLineIndent = firstLineIndent;
            Aligment = (int)aligment;
            SymbolsCount = paragraph.Range.Text.Length;
            PrefixIsNumber = prefixIsNumber;
            PrefixIsLowercase = prefixIsLowercase;
            PrefixIsUppercase = prefixIsUppercase;
            PrefixIsDash = prefixIsDash;
            SuffixIsEndSign = suffixIsEndSign;
            SuffixIsColon = suffixIsColon;
            SuffixIsCommaOrSemicolon = suffixIsCommaOrSemicolon;
            ContainsDash = containsDash;
            ContainsBracket = containsBracket;
            FontSize = fontSize;
            LineSpacing = lineSpacing;
            LineSpacingRule = (int)lineSpacingRule;
            Italic = (int)italic;
            Bold = (int)bold;
            BlackColor = blackColor;
        }
    }
}
