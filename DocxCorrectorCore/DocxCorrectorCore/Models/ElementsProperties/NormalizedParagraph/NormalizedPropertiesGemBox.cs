using System;
using System.Linq;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.Models
{
    public sealed class NormalizedPropertiesGemBox : NormalizedProperties
    {
        public NormalizedPropertiesGemBox(Word.Paragraph paragraph, int paragraphId)
        {
            // Номер параграфа
            int id = paragraphId;
            // Отступ первой строки
            float firstLineIndent = -(float)paragraph.ParagraphFormat.SpecialIndentation;
            // Выравнивание
            NormalizedAligment aligment;
            switch (paragraph.ParagraphFormat.Alignment)
            {
                case Word.HorizontalAlignment.Left:
                    aligment = NormalizedAligment.Left;
                    break;
                case Word.HorizontalAlignment.Center:
                    aligment = NormalizedAligment.Center;
                    break;
                case Word.HorizontalAlignment.Right:
                    aligment = NormalizedAligment.Right;
                    break;
                case Word.HorizontalAlignment.Justify:
                    aligment = NormalizedAligment.Justify;
                    break;
                default:
                    aligment = NormalizedAligment.Other;
                    break;
            }
            // Первый символ - число
            int prefixIsNumber = Char.IsDigit(paragraph.Content.ToString()[0]) ? 1 : 0;
            // Первый символ - маленькая буква
            int prefixIsLowercase = Char.IsLower(paragraph.Content.ToString()[0]) ? 1 : 0;
            // Первый символ - большая буква
            int prefixIsUppercase = Char.IsUpper(paragraph.Content.ToString()[0]) ? 1 : 0;
            // Первый симол - тире
            string[] dashes = new string[] { "-", "־", "᠆", "‐", "‑", "‒", "–", "—", "―", "﹘", "﹣", "－" };
            int prefixIsDash = GemBoxHelper.CheckIfFirstSymbolOfParagraphIs(paragraph, dashes);
            // Последний символ - знак окончания
            string[] endSigns = new string[] { ".", "!", "?" };
            int suffixIsEndSign = GemBoxHelper.CheckIfLastSymbolOfParagraphIs(paragraph, endSigns);
            // Последний символ - двоеточие
            string[] colon = new string[] { ":" };
            // Последний символ - запятая или точка с запятой
            int suffixIsColon = GemBoxHelper.CheckIfLastSymbolOfParagraphIs(paragraph, colon);
            string[] commaAndSemicolon = new string[] { ",", ";" };
            int suffixIsCommaOrSemicolon = GemBoxHelper.CheckIfLastSymbolOfParagraphIs(paragraph, commaAndSemicolon);
            // В параграфе содержится тире
            int containsDash = GemBoxHelper.CheckIfParagraphsContainsOneOf(paragraph, dashes);
            // В параграфе содержится скобка
            string[] bracket = new string[] { ")" };
            int containsBracket = GemBoxHelper.CheckIfParagraphsContainsOneOf(paragraph, bracket);
            // Размер шрифта
            float fontSize = 0;
            var elements = paragraph.GetChildElements(true, Word.ElementType.Run);
            if (elements.Count() != 0)
            {
                fontSize = (float)((Word.Run)elements.First()).CharacterFormat.Size;
                foreach (Word.Run run in elements)
                {
                    if ((float)run.CharacterFormat.Size != fontSize) { fontSize = 9999999; break; }
                }
            }
            // Междустрочный интервал
            float lineSpacing = (float)paragraph.ParagraphFormat.LineSpacing;
            // Правило междустрочного интервала
            LineSpacingRuleVariations lineSpacingRule;
            if (paragraph.ParagraphFormat.LineSpacingRule == Word.LineSpacingRule.Multiple)
            {
                switch (lineSpacing)
                {
                    case 1.0f:
                        lineSpacingRule = LineSpacingRuleVariations.Single;
                        break;
                    case 1.5f:
                        lineSpacingRule = LineSpacingRuleVariations.OneAndHalf;
                        break;
                    case 2.0f:
                        lineSpacingRule = LineSpacingRuleVariations.Double;
                        break;
                    default:
                        lineSpacingRule = LineSpacingRuleVariations.Miltiply;
                        break;
                }
            } else
            {
                lineSpacingRule = LineSpacingRuleVariations.Other;
            }
            // Курсив   
            ContainsStatus italic = ContainsStatus.None;
            if (elements.Count() != 0)
            {
                italic = ((Word.Run)elements.First()).CharacterFormat.Italic ? ContainsStatus.Full : ContainsStatus.None;
                var italicGemBox = ((Word.Run)elements.First()).CharacterFormat.Italic;
                foreach (Word.Run run in elements)
                {
                    if (run.CharacterFormat.Italic != italicGemBox) { italic = ContainsStatus.Contains; break; }
                }
            }
            // Жирный
            ContainsStatus bold = ContainsStatus.None;
            if (elements.Count() != 0)
            {
                bold = ((Word.Run)elements.First()).CharacterFormat.Bold ? ContainsStatus.Full : ContainsStatus.None;
                var boldGemBox = ((Word.Run)elements.First()).CharacterFormat.Bold;
                foreach (Word.Run run in elements)
                {
                    if (run.CharacterFormat.Bold != boldGemBox) { bold = ContainsStatus.Contains; break; }
                }
            }
            // Черный цвет
            int blackColor = 1;
            if (elements.Count() != 0)
            {
                var colorGemBox = ((Word.Run)elements.First()).CharacterFormat.FontColor;
                var blackGemBox = new Word.Color(0, 0, 0);
                blackColor = (colorGemBox.Equals(blackGemBox)) ? 1 : 0;
                foreach (Word.Run run in elements)
                {
                    if (!run.CharacterFormat.FontColor.Equals(blackGemBox)) { blackColor = 0; break; }
                }
            }

            Id = id;
            FirstLineIndent = firstLineIndent;
            Aligment = (int)aligment;
            SymbolsCount = paragraph.Content.ToString().Length - 1;
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
