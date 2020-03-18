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
        public string FontName { get; }
        public string FontSize { get; }
        // ParagraphFormat
        public string Alignment { get; }
        public string BackgroundColor { get; }
        public string Borders { get; }
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
        public string Tabs { get; }
        public string WidowControl { get; }
        // ListFormat
        public string ListFormatIsList { get; }
        public string ListFormatLevel { get; }
        // RunnersFormat
        public List<Dictionary<string, string>> RunnersFormat { get; }

        public ParagraphPropertiesGemBox(Word.Paragraph paragraph)
        {
            Content = paragraph.Content.ToString().Remove(paragraph.Content.ToString().Length - 1);
            // CharacterFormatForParagraphMark
            FullBold = paragraph.CharacterFormatForParagraphMark.Bold.ToString();
            FullItalic = paragraph.CharacterFormatForParagraphMark.Italic.ToString();
            FontName = paragraph.CharacterFormatForParagraphMark.FontName.ToString();
            FontSize = paragraph.CharacterFormatForParagraphMark.Size.ToString();
            // ParagraphFormat
            Alignment = paragraph.ParagraphFormat.Alignment.ToString();
            BackgroundColor = paragraph.ParagraphFormat.BackgroundColor.ToString();
            Borders = paragraph.ParagraphFormat.Borders.ToString();
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
            Tabs = paragraph.ParagraphFormat.Tabs.ToString();
            WidowControl = paragraph.ParagraphFormat.WidowControl.ToString();
            // ListFormat
            ListFormatIsList = paragraph.ListFormat.IsList.ToString();
            ListFormatLevel = paragraph.ListFormat.ListLevelNumber.ToString();
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
                    { "\nBorder", runner.CharacterFormat.Border.ToString() },
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
