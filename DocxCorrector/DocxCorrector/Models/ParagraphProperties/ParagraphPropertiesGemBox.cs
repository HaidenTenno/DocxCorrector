#nullable enable
using System;
using System.Collections.Generic;
using Word = GemBox.Document;

namespace DocxCorrector.Models
{
    public sealed class ParagraphPropertiesGemBox : ParagraphProperties
    {
        public string Content { get; }
        // CharacterFormatForParagraphMark
        public string FullBold { get; }
        public string FullItalic { get; }
        // ParagraphFormat
        public string Alignment { get; }
        public string BackgroundColor { get; }
        public string KeepLinesTogether { get; }
        public string KeepWithNext { get; }
        public string LeftIndentation { get; }
        public string LineSpacing { get; }
        public string LineSpacingRule { get; }
        public string MirrorIndents { get; }
        public string NoSpaceBetweenParagraphsOfSameStyle { get; }
        public string OutlineLevel { get; }
        public string PageBreakBefore { get; }
        public string RightIndentation { get; }
        public string RightToLeft { get; }
        public string SpaceAfter { get; }
        public string SpaceBefore { get; }
        public string SpecialIndentation { get; }
        public string Style { get; }
        public string WidowControl { get; }
        // ListFormat
        public string ListFormatIsList { get; }
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
        // RunnersFormat
        public List<Dictionary<string, string>> RunnersFormat { get; }

        public ParagraphPropertiesGemBox(Word.Paragraph paragraph)
        {
            Content = paragraph.Content.ToString().Remove(paragraph.Content.ToString().Length - 1);
            // CharacterFormatForParagraphMark
            FullBold = paragraph.CharacterFormatForParagraphMark.Bold.ToString();
            FullItalic = paragraph.CharacterFormatForParagraphMark.Italic.ToString();
            // ParagraphFormat
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
            // ListFormat
            ListFormatIsList = paragraph.ListFormat.IsList.ToString();
            if (paragraph.ListFormat.IsList)
            {
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
            // RunnersFormat
            RunnersFormat = new List<Dictionary<string, string>>();
            foreach (Word.Run runner in paragraph.GetChildElements(true, Word.ElementType.Run))
            {
                Dictionary<string, string> runnerFormat = new Dictionary<string, string>()
                {
                    { "Bold", runner.CharacterFormat.Bold.ToString() },
                    { "\nItalic", runner.CharacterFormat.Italic.ToString() },
                    { "\nAppCaps", runner.CharacterFormat.AllCaps.ToString() },
                    { "\nBackgroundColor", runner.CharacterFormat.BackgroundColor.ToString() },
                    { "\nDoubleStrikethrough", runner.CharacterFormat.DoubleStrikethrough.ToString() },
                    { "\nFontColor", runner.CharacterFormat.FontColor.ToString() },
                    { "\nFontName", runner.CharacterFormat.FontName.ToString() },
                    { "\nHidden", runner.CharacterFormat.Hidden.ToString() },
                    { "\nHighlightColor", runner.CharacterFormat.HighlightColor.ToString() },
                    { "\nKerning", runner.CharacterFormat.Kerning.ToString() },
                    { "\nLanguage", runner.CharacterFormat.Language.ToString() },
                    { "\nPosition", runner.CharacterFormat.Position.ToString() },
                    { "\nRightToLeft", runner.CharacterFormat.RightToLeft.ToString() },
                    { "\nScaling", runner.CharacterFormat.Scaling.ToString() },
                    { "\nSize", runner.CharacterFormat.Size.ToString() },
                    { "\nSmallCaps", runner.CharacterFormat.SmallCaps.ToString() },
                    { "\nSpacing", runner.CharacterFormat.Spacing.ToString() },
                    { "\nStrikethrough", runner.CharacterFormat.Strikethrough.ToString() },
                    { "\nStyle", runner.CharacterFormat.Style.ToString() },
                    { "\nSubscript", runner.CharacterFormat.Subscript.ToString() },
                    { "\nSuperscript", runner.CharacterFormat.Superscript.ToString() },
                    { "\nUnderlineColor", runner.CharacterFormat.UnderlineColor.ToString() },
                    { "\nUnderlineStyle", runner.CharacterFormat.UnderlineStyle.ToString() }
                };

                RunnersFormat.Add(runnerFormat);
            }
        }
    }
}
