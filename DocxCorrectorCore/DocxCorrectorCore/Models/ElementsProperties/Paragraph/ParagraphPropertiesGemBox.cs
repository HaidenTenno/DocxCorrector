using System;
using System.Collections.Generic;
using Word = GemBox.Document;
using DocxCorrectorCore.Services.Helpers;

namespace DocxCorrectorCore.Models
{
    public sealed class ParagraphPropertiesGemBox : ParagraphProperties
    {
        public string Content { get; }
        // First word is (Таблица, Табл, Рисунок, Рис., Рис, Табл.)
        public string? FirstKeyWord { get; }
        // Element marks
        public string? PrevElementMark { get; }
        public string? CurElementMark { get; }
        public string? NextElementMark { get; }
        // CharacterFormatForParagraphMark
        public string? FullBold { get; }
        public string? FullItalic { get; }
        // ParagraphFormat
        public string? Alignment { get; }
        public string? BackgroundColor { get; }
        public string? KeepLinesTogether { get; }
        public string? KeepWithNext { get; }
        public string? LeftIndentation { get; }
        public string? LineSpacing { get; }
        public string? LineSpacingRule { get; }
        public string? MirrorIndents { get; }
        public string? NoSpaceBetweenParagraphsOfSameStyle { get; }
        public string? OutlineLevel { get; }
        public string? PageBreakBefore { get; }
        public string? RightIndentation { get; }
        public string? RightToLeft { get; }
        public string? SpaceAfter { get; }
        public string? SpaceBefore { get; }
        public string? SpecialIndentation { get; }
        public string? Style { get; }
        public string? WidowControl { get; }
        // ListFormat
        public string? ListFormatIsList { get; }
        public string? ListStyleHash { get; }
        public string? ListItem { get; }
        public string? ListFormatLevel { get; }
        public string? ListFormat { get; }
        public string? CurrentListAligment { get; }
        public string? CurrentListIsLegal { get; }
        public string? CurrentListLevel { get; }
        public string? CurrentListNumberFormat { get; }
        public string? CurrentListNumberPosition { get; }
        public string? CurrentListNumberStyle { get; }
        public string? CurrentListRestartAfterLevel { get; }
        public string? CurrentListStartAt { get; }
        public string? CurrentListTextPosition { get; }
        public string? CurrentListTrailingCharacter { get; }

        // Private
        private string GetProperContent(Word.Paragraph paragraph)
        {
            string result = "";

            foreach (var element in paragraph.GetChildElements(false, new Word.ElementType[] { Word.ElementType.Run, Word.ElementType.Picture, Word.ElementType.Chart, Word.ElementType.Shape }))
            {
                switch (element)
                {
                    case Word.Run run:
                        result += run.Content;
                        break;
                    case Word.Picture picture:
                        result += "PICTURE ";
                        break;
                    case Word.Chart _:
                        result += "CHART ";
                        break;
                    case Word.Drawing.Shape _:
                        result += "SHAPE ";
                        break;
                    default:
                        Console.WriteLine("Unsupported element");
                        break;
                }
            }

            result = result.Trim();

            if (result == "") { result = "SPACE"; }
            return result;
        }

        // Public

        public ParagraphPropertiesGemBox(Word.Paragraph paragraph)
        {
            Content = GetProperContent(paragraph);
            FirstKeyWord = GemBoxHelper.CheckIfFirtWordOfParagraphIsOneOf(paragraph, new string[] { "Таблица", "Табл", "Рисунок", "Рис.", "Рис", "Табл." });
            // Свойства символов всего параграфа
            FullBold = paragraph.CharacterFormatForParagraphMark.Bold.ToString();
            FullItalic = paragraph.CharacterFormatForParagraphMark.Italic.ToString();
            // Свойства параграфа
            Alignment = paragraph.ParagraphFormat.Alignment.ToString();
            BackgroundColor = paragraph.ParagraphFormat.BackgroundColor.ToString();
            KeepLinesTogether = paragraph.ParagraphFormat.KeepLinesTogether.ToString();
            KeepWithNext = paragraph.ParagraphFormat.KeepWithNext.ToString();
            LeftIndentation = paragraph.ParagraphFormat.LeftIndentation.ToString();
            LineSpacing = paragraph.ParagraphFormat.LineSpacing.ToString();
            LineSpacingRule = paragraph.ParagraphFormat.LineSpacingRule.ToString();
            MirrorIndents = paragraph.ParagraphFormat.MirrorIndents.ToString();
            NoSpaceBetweenParagraphsOfSameStyle = paragraph.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle.ToString();
            OutlineLevel = paragraph.ParagraphFormat.OutlineLevel.ToString();
            PageBreakBefore = paragraph.ParagraphFormat.PageBreakBefore.ToString();
            RightIndentation = paragraph.ParagraphFormat.RightIndentation.ToString();
            RightToLeft = paragraph.ParagraphFormat.RightToLeft.ToString();
            SpaceAfter = paragraph.ParagraphFormat.SpaceAfter.ToString();
            SpaceBefore = paragraph.ParagraphFormat.SpaceBefore.ToString();
            SpecialIndentation = paragraph.ParagraphFormat.SpecialIndentation.ToString();
            Style = paragraph.ParagraphFormat.Style.ToString();
            WidowControl = paragraph.ParagraphFormat.WidowControl.ToString();
            // Свойства списка
            ListFormatIsList = paragraph.ListFormat.IsList.ToString();
            if (paragraph.ListFormat.IsList)
            {
                ListStyleHash = paragraph.ListFormat.Style.GetHashCode().ToString();
                ListItem = paragraph.ListItem.ToString();
                ListFormatLevel = paragraph.ListFormat.ListLevelNumber.ToString();
                ListFormat = paragraph.ListFormat.ListLevelFormat.ToString();
                CurrentListAligment = paragraph.ListFormat.ListLevelFormat.Alignment.ToString();
                CurrentListIsLegal = paragraph.ListFormat.ListLevelFormat.IsLegal.ToString();
                CurrentListLevel = paragraph.ListFormat.ListLevelFormat.Level.ToString();
                CurrentListNumberFormat = paragraph.ListFormat.ListLevelFormat.NumberFormat.ToString();
                CurrentListNumberPosition = paragraph.ListFormat.ListLevelFormat.NumberPosition.ToString();
                CurrentListNumberStyle = paragraph.ListFormat.ListLevelFormat.NumberStyle.ToString();
                CurrentListRestartAfterLevel = paragraph.ListFormat.ListLevelFormat.RestartAfterLevel.ToString();
                CurrentListStartAt = paragraph.ListFormat.ListLevelFormat.StartAt.ToString();
                CurrentListTextPosition = paragraph.ListFormat.ListLevelFormat.TextPosition.ToString();
                CurrentListTrailingCharacter = paragraph.ListFormat.ListLevelFormat.TrailingCharacter.ToString();
            }
        }

        // PlaceHolder constructor
        public ParagraphPropertiesGemBox(string placeHolder)
        {
            Content = placeHolder;
        }       
    }
}
