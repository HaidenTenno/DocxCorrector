using System;
using System.Collections.Generic;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel;
using Word = GemBox.Document;
using Newtonsoft.Json;

namespace DocxCorrectorCore.BusinessLogicLayer.PropertiesPuller
{
    public sealed class PresetValue
    {
        public readonly ParagraphClass ParagraphClass;

        // Свойства ParagraphFormat
        public readonly List<Word.HorizontalAlignment> Alignment;
        public readonly List<Word.Color> BackgroundColor;
        public readonly List<Word.BorderStyle> BorderStyle;
        public readonly List<bool> KeepLinesTogether;
        public readonly List<bool> KeepWithNext;
        public readonly List<double> LeftIndentation;
        public readonly List<double> LineSpacing;
        public readonly List<Word.LineSpacingRule> LineSpacingRule;
        public readonly List<bool> MirrorIndents;
        public readonly List<bool> NoSpaceBetweenParagraphsOfSameStyle;
        public readonly List<Word.OutlineLevel> OutlineLevel;
        public readonly List<bool> PageBreakBefore;
        public readonly List<double> RightIndentation;
        public readonly List<bool> RightToLeft;
        public readonly List<double> SpaceAfter;
        public readonly List<double> SpaceBefore;
        public readonly double SpecialIndentationLeftBorder;
        public readonly double SpecialIndentationRightBorder;
        public readonly List<bool> WidowControl;

        // Свойства CharacterFormat для всего абзаца
        public readonly List<bool> WholeParagraphAllCaps;
        public readonly List<Word.Color> WholeParagraphBackgroundColor;
        public readonly List<bool> WholeParagraphBold;
        public readonly List<Word.SingleBorder> WholeParagraphBorder;
        public readonly List<bool> WholeParagraphDoubleStrikethrough;
        public readonly List<Word.Color> WholeParagraphFontColor;
        public readonly List<string> WholeParagraphFontName;
        public readonly List<bool> WholeParagraphHidden;
        public readonly List<Word.Color> WholeParagraphHighlightColor;
        public readonly List<bool> WholeParagraphItalic;
        public readonly List<double> WholeParagraphKerning;
        public readonly List<double> WholeParagraphPosition;
        public readonly List<bool> WholeParagraphRightToLeft;
        public readonly List<int> WholeParagraphScaling;
        public readonly double WholeParagraphSizeLeftBorder;
        public readonly double WholeParagraphSizeRightBorder;
        public readonly List<bool> WholeParagraphSmallCaps;
        public readonly List<double> WholeParagraphSpacing;
        public readonly List<bool> WholeParagraphStrikethrough;
        public readonly List<bool> WholeParagraphSubscript;
        public readonly List<bool> WholeParagraphSuperscript;
        public readonly List<Word.UnderlineType> WholeParagraphUnderlineStyle;

        [JsonConstructor]
        public PresetValue(
            ParagraphClass paragraphClass,
            List<Word.HorizontalAlignment> alignment,
            List<Word.Color> backgroundColor,
            List<Word.BorderStyle> borderStyle,
            List<bool> keepLinesTogether,
            List<bool> keepWithNext,
            List<double> leftIndentation,
            List<double> lineSpacing,
            List<Word.LineSpacingRule> lineSpacingRule,
            List<bool> mirrorIndents,
            List<bool> noSpaceBetweenParagraphsOfSameStyle,
            List<Word.OutlineLevel> outlineLevel,
            List<bool> pageBreakBefore,
            List<double> rightIndentation,
            List<bool> rightToLeft,
            List<double> spaceAfter,
            List<double> spaceBefore,
            double specialIndentationLeftBorder,
            double specialIndentationRightBorder,
            List<bool> widowControl,
            List<bool> wholeParagraphAllCaps,
            List<Word.Color> wholeParagraphBackgroundColor,
            List<bool> wholeParagraphBold,
            List<Word.SingleBorder> wholeParagraphBorder,
            List<bool> wholeParagraphDoubleStrikethrough,
            List<Word.Color> wholeParagraphFontColor,
            List<string> wholeParagraphFontName,
            List<bool> wholeParagraphHidden,
            List<Word.Color> wholeParagraphHighlightColor,
            List<bool> wholeParagraphItalic,
            List<double> wholeParagraphKerning,
            List<double> wholeParagraphPosition,
            List<bool> wholeParagraphRightToLeft,
            List<int> wholeParagraphScaling,
            double wholeParagraphSizeLeftBorder,
            double wholeParagraphSizeRightBorder,
            List<bool> wholeParagraphSmallCaps,
            List<double> wholeParagraphSpacing,
            List<bool> wholeParagraphStrikethrough,
            List<bool> wholeParagraphSubscript,
            List<bool> wholeParagraphSuperscript,
            List<Word.UnderlineType> wholeParagraphUnderlineStyle
            )
        {
            ParagraphClass = paragraphClass;
            Alignment = alignment;
            BackgroundColor = backgroundColor;
            BorderStyle = borderStyle;
            KeepLinesTogether = keepLinesTogether;
            KeepWithNext = keepWithNext;
            LeftIndentation = leftIndentation;
            LineSpacing = lineSpacing;
            LineSpacingRule = lineSpacingRule;
            MirrorIndents = mirrorIndents;
            NoSpaceBetweenParagraphsOfSameStyle = noSpaceBetweenParagraphsOfSameStyle;
            OutlineLevel = outlineLevel;
            PageBreakBefore = pageBreakBefore;
            RightIndentation = rightIndentation;
            RightToLeft = rightToLeft;
            SpaceAfter = spaceAfter;
            SpaceBefore = spaceBefore;
            SpecialIndentationLeftBorder = specialIndentationLeftBorder;
            SpecialIndentationRightBorder = specialIndentationRightBorder;
            WidowControl = widowControl;
            WholeParagraphAllCaps = wholeParagraphAllCaps;
            WholeParagraphBackgroundColor = wholeParagraphBackgroundColor;
            WholeParagraphBold = wholeParagraphBold;
            WholeParagraphBorder = wholeParagraphBorder;
            WholeParagraphDoubleStrikethrough = wholeParagraphDoubleStrikethrough;
            WholeParagraphFontColor = wholeParagraphFontColor;
            WholeParagraphFontName = wholeParagraphFontName;
            WholeParagraphHidden = wholeParagraphHidden;
            WholeParagraphHighlightColor = wholeParagraphHighlightColor;
            WholeParagraphItalic = wholeParagraphItalic;
            WholeParagraphKerning = wholeParagraphKerning;
            WholeParagraphPosition = wholeParagraphPosition;
            WholeParagraphRightToLeft = wholeParagraphRightToLeft;
            WholeParagraphScaling = wholeParagraphScaling;
            WholeParagraphSizeLeftBorder = wholeParagraphSizeLeftBorder;
            WholeParagraphSizeRightBorder = wholeParagraphSizeRightBorder;
            WholeParagraphSmallCaps = wholeParagraphSmallCaps;
            WholeParagraphSpacing = wholeParagraphSpacing;
            WholeParagraphStrikethrough = wholeParagraphStrikethrough;
            WholeParagraphSubscript = wholeParagraphSubscript;
            WholeParagraphSuperscript = wholeParagraphSuperscript;
            WholeParagraphUnderlineStyle = wholeParagraphUnderlineStyle;
        }

        public PresetValue(DocumentElement documentElement)
        {
            ParagraphClass = documentElement.ParagraphClass;
            Alignment = documentElement.Alignment;
            BackgroundColor = documentElement.BackgroundColor;
            BorderStyle = documentElement.BorderStyle;
            KeepLinesTogether = documentElement.KeepLinesTogether;
            KeepWithNext = documentElement.KeepWithNext;
            LeftIndentation = documentElement.LeftIndentation;
            LineSpacing = documentElement.LineSpacing;
            LineSpacingRule = documentElement.LineSpacingRule;
            MirrorIndents = documentElement.MirrorIndents;
            NoSpaceBetweenParagraphsOfSameStyle = documentElement.NoSpaceBetweenParagraphsOfSameStyle;
            OutlineLevel = documentElement.OutlineLevel;
            PageBreakBefore = documentElement.PageBreakBefore;
            RightIndentation = documentElement.RightIndentation;
            RightToLeft = documentElement.RightToLeft;
            SpaceAfter = documentElement.SpaceAfter;
            SpaceBefore = documentElement.SpaceBefore;
            SpecialIndentationLeftBorder = documentElement.SpecialIndentationLeftBorder;
            SpecialIndentationRightBorder = documentElement.SpecialIndentationRightBorder;
            WidowControl = documentElement.WidowControl;
            WholeParagraphAllCaps = documentElement.WholeParagraphAllCaps;
            WholeParagraphBackgroundColor = documentElement.WholeParagraphBackgroundColor;
            WholeParagraphBold = documentElement.WholeParagraphBold;
            WholeParagraphBorder = documentElement.WholeParagraphBorder;
            WholeParagraphDoubleStrikethrough = documentElement.WholeParagraphDoubleStrikethrough;
            WholeParagraphFontColor = documentElement.WholeParagraphFontColor;
            WholeParagraphFontName = documentElement.WholeParagraphFontName;
            WholeParagraphHidden = documentElement.WholeParagraphHidden;
            WholeParagraphHighlightColor = documentElement.WholeParagraphHighlightColor;
            WholeParagraphItalic = documentElement.WholeParagraphItalic;
            WholeParagraphKerning = documentElement.WholeParagraphKerning;
            WholeParagraphPosition = documentElement.WholeParagraphPosition;
            WholeParagraphRightToLeft = documentElement.WholeParagraphRightToLeft;
            WholeParagraphScaling = documentElement.WholeParagraphScaling;
            WholeParagraphSizeLeftBorder = documentElement.WholeParagraphSizeLeftBorder;
            WholeParagraphSizeRightBorder = documentElement.WholeParagraphSizeRightBorder;
            WholeParagraphSmallCaps = documentElement.WholeParagraphSmallCaps;
            WholeParagraphSpacing = documentElement.WholeParagraphSpacing;
            WholeParagraphStrikethrough = documentElement.WholeParagraphStrikethrough;
            WholeParagraphSubscript = documentElement.WholeParagraphSubscript;
            WholeParagraphSuperscript = documentElement.WholeParagraphSuperscript;
            WholeParagraphUnderlineStyle = documentElement.WholeParagraphUnderlineStyle;
        }

        public PresetValue(Word.Paragraph paragraph)
        {
            ParagraphClass = ParagraphClass.NoClass;
            Alignment = new List<Word.HorizontalAlignment>() { paragraph.ParagraphFormat.Alignment };
            BackgroundColor = new List<Word.Color>() { paragraph.ParagraphFormat.BackgroundColor };
            BorderStyle = GetParagraphBorders(paragraph);
            KeepLinesTogether = new List<bool>() { paragraph.ParagraphFormat.KeepLinesTogether };
            KeepWithNext = new List<bool>() { paragraph.ParagraphFormat.KeepWithNext };
            LeftIndentation = new List<double>() { paragraph.ParagraphFormat.LeftIndentation };
            LineSpacing = new List<double>() { paragraph.ParagraphFormat.LineSpacing };
            LineSpacingRule = new List<Word.LineSpacingRule>() { paragraph.ParagraphFormat.LineSpacingRule };
            MirrorIndents = new List<bool>() { paragraph.ParagraphFormat.MirrorIndents };
            NoSpaceBetweenParagraphsOfSameStyle = new List<bool>() { paragraph.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle };
            OutlineLevel = new List<Word.OutlineLevel>() { paragraph.ParagraphFormat.OutlineLevel };
            PageBreakBefore = new List<bool>() { paragraph.ParagraphFormat.PageBreakBefore };
            RightIndentation = new List<double>() { paragraph.ParagraphFormat.RightIndentation };
            RightToLeft = new List<bool>() { paragraph.ParagraphFormat.RightToLeft };
            SpaceAfter = new List<double>() { paragraph.ParagraphFormat.SpaceAfter };
            SpaceBefore = new List<double>() { paragraph.ParagraphFormat.SpaceBefore };
            SpecialIndentationLeftBorder = paragraph.ParagraphFormat.SpecialIndentation;
            SpecialIndentationRightBorder = paragraph.ParagraphFormat.SpecialIndentation;
            WidowControl = new List<bool>() { paragraph.ParagraphFormat.WidowControl };
            WholeParagraphAllCaps = new List<bool>() { paragraph.CharacterFormatForParagraphMark.AllCaps };
            WholeParagraphBackgroundColor = new List<Word.Color>() { paragraph.CharacterFormatForParagraphMark.BackgroundColor };
            WholeParagraphBold = new List<bool>() { paragraph.CharacterFormatForParagraphMark.Bold };
            WholeParagraphBorder = new List<Word.SingleBorder>() { paragraph.CharacterFormatForParagraphMark.Border };
            WholeParagraphDoubleStrikethrough = new List<bool>() { paragraph.CharacterFormatForParagraphMark.DoubleStrikethrough };
            WholeParagraphFontColor = new List<Word.Color>() { paragraph.CharacterFormatForParagraphMark.FontColor };
            WholeParagraphFontName = new List<string>() { paragraph.CharacterFormatForParagraphMark.FontName };
            WholeParagraphHidden = new List<bool>() { paragraph.CharacterFormatForParagraphMark.Hidden };
            WholeParagraphHighlightColor = new List<Word.Color>() { paragraph.CharacterFormatForParagraphMark.HighlightColor };
            WholeParagraphItalic = new List<bool>() { paragraph.CharacterFormatForParagraphMark.Italic };
            WholeParagraphKerning = new List<double>() { paragraph.CharacterFormatForParagraphMark.Kerning };
            WholeParagraphPosition = new List<double>() { paragraph.CharacterFormatForParagraphMark.Position };
            WholeParagraphRightToLeft = new List<bool>() { paragraph.CharacterFormatForParagraphMark.RightToLeft };
            WholeParagraphScaling = new List<int>() { paragraph.CharacterFormatForParagraphMark.Scaling };
            WholeParagraphSizeLeftBorder = paragraph.CharacterFormatForParagraphMark.Size;
            WholeParagraphSizeRightBorder = paragraph.CharacterFormatForParagraphMark.Size;
            WholeParagraphSmallCaps = new List<bool>() { paragraph.CharacterFormatForParagraphMark.SmallCaps };
            WholeParagraphSpacing = new List<double>() { paragraph.CharacterFormatForParagraphMark.Spacing };
            WholeParagraphStrikethrough = new List<bool>() { paragraph.CharacterFormatForParagraphMark.Strikethrough };
            WholeParagraphSubscript = new List<bool>() { paragraph.CharacterFormatForParagraphMark.Subscript };
            WholeParagraphSuperscript = new List<bool>() { paragraph.CharacterFormatForParagraphMark.Superscript };
            WholeParagraphUnderlineStyle = new List<Word.UnderlineType>() { paragraph.CharacterFormatForParagraphMark.UnderlineStyle };
        }

        private List<Word.BorderStyle> GetParagraphBorders(Word.Paragraph paragraph)
        {
            List<Word.BorderStyle> borders = new List<Word.BorderStyle>();

            foreach (Word.SingleBorderType borderType in Enum.GetValues(typeof(Word.SingleBorderType)))
            {
                borders.Add(paragraph.ParagraphFormat.Borders[borderType].Style);
            }
            return borders;
        }

        private bool CheckParagraphFormatBorder(Word.Paragraph paragraph)
        {
            if (BorderStyle.Count == 0) { return true; }

            foreach (Word.SingleBorderType borderType in Enum.GetValues(typeof(Word.SingleBorderType)))
            {
                if (!BorderStyle.Contains(paragraph.ParagraphFormat.Borders[borderType].Style)) { return false; }
            }

            return true;
        }

        private bool CheckWholeParagraphBorder(Word.Paragraph paragraph)
        {
            if ((WholeParagraphBorder.Count != 0) & !WholeParagraphBorder.Contains(paragraph.CharacterFormatForParagraphMark.Border)) { return false; }

            return true;
        }

        private bool CheckParagraphFormat(Word.Paragraph paragraph)
        {
            if ((Alignment.Count != 0) & !Alignment.Contains(paragraph.ParagraphFormat.Alignment)) { return false; }
            if ((BackgroundColor.Count != 0) & !BackgroundColor.Contains(paragraph.ParagraphFormat.BackgroundColor)) { return false; }
            if (!CheckParagraphFormatBorder(paragraph)) { return false; }
            if ((KeepLinesTogether.Count != 0) & !KeepLinesTogether.Contains(paragraph.ParagraphFormat.KeepLinesTogether)) { return false; }
            if ((KeepWithNext.Count != 0) & !KeepWithNext.Contains(paragraph.ParagraphFormat.KeepWithNext)) { return false; }
            if ((LeftIndentation.Count != 0) & !LeftIndentation.Contains(paragraph.ParagraphFormat.LeftIndentation)) { return false; }
            if ((LineSpacing.Count != 0) & !LineSpacing.Contains(paragraph.ParagraphFormat.LineSpacing)) { return false; }
            if ((MirrorIndents.Count != 0) & !MirrorIndents.Contains(paragraph.ParagraphFormat.MirrorIndents)) { return false; }
            if ((NoSpaceBetweenParagraphsOfSameStyle.Count != 0) & !NoSpaceBetweenParagraphsOfSameStyle.Contains(paragraph.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle)) { return false; }
            if ((OutlineLevel.Count != 0) & !OutlineLevel.Contains(paragraph.ParagraphFormat.OutlineLevel)) { return false; }
            if ((PageBreakBefore.Count != 0) & !PageBreakBefore.Contains(paragraph.ParagraphFormat.PageBreakBefore)) { return false; }
            if ((RightIndentation.Count != 0) & !RightIndentation.Contains(paragraph.ParagraphFormat.RightIndentation)) { return false; }
            if ((RightToLeft.Count != 0) & !RightToLeft.Contains(paragraph.ParagraphFormat.RightToLeft)) { return false; }
            if ((SpaceAfter.Count != 0) & !SpaceAfter.Contains(paragraph.ParagraphFormat.SpaceAfter)) { return false; }
            if ((SpaceBefore.Count != 0) & !SpaceBefore.Contains(paragraph.ParagraphFormat.SpaceBefore)) { return false; }
            if ((paragraph.ParagraphFormat.SpecialIndentation < SpecialIndentationLeftBorder) | ((paragraph.ParagraphFormat.SpecialIndentation > SpecialIndentationRightBorder))) { return false; }
            if ((WidowControl.Count != 0) & !WidowControl.Contains(paragraph.ParagraphFormat.WidowControl)) { return false; }

            return true;
        }

        private bool CheckWholeParagraphCharacterFormat(Word.Paragraph paragraph)
        {
            if ((WholeParagraphAllCaps.Count != 0) & !WholeParagraphAllCaps.Contains(paragraph.CharacterFormatForParagraphMark.AllCaps)) { return false; }
            if ((WholeParagraphBackgroundColor.Count != 0) & !WholeParagraphBackgroundColor.Contains(paragraph.CharacterFormatForParagraphMark.BackgroundColor)) { return false; }
            if ((WholeParagraphBold.Count != 0) & !WholeParagraphBold.Contains(paragraph.CharacterFormatForParagraphMark.Bold)) { return false; }
            if (!CheckWholeParagraphBorder(paragraph)) { return false; }
            if ((WholeParagraphDoubleStrikethrough.Count != 0) & !WholeParagraphDoubleStrikethrough.Contains(paragraph.CharacterFormatForParagraphMark.DoubleStrikethrough)) { return false; }
            if ((WholeParagraphFontColor.Count != 0) & !WholeParagraphFontColor.Contains(paragraph.CharacterFormatForParagraphMark.FontColor)) { return false; }
            if ((WholeParagraphFontName.Count != 0) & !WholeParagraphFontName.Contains(paragraph.CharacterFormatForParagraphMark.FontName)) { return false; }
            if ((WholeParagraphHidden.Count != 0) & !WholeParagraphHidden.Contains(paragraph.CharacterFormatForParagraphMark.Hidden)) { return false; }
            if ((WholeParagraphHighlightColor.Count != 0) & !WholeParagraphHighlightColor.Contains(paragraph.CharacterFormatForParagraphMark.HighlightColor)) { return false; }
            if ((WholeParagraphItalic.Count != 0) & !WholeParagraphItalic.Contains(paragraph.CharacterFormatForParagraphMark.Italic)) { return false; }
            if ((WholeParagraphKerning.Count != 0) & !WholeParagraphKerning.Contains(paragraph.CharacterFormatForParagraphMark.Kerning)) { return false; }
            if ((WholeParagraphPosition.Count != 0) & !WholeParagraphPosition.Contains(paragraph.CharacterFormatForParagraphMark.Position)) { return false; }
            if ((WholeParagraphRightToLeft.Count != 0) & !WholeParagraphRightToLeft.Contains(paragraph.CharacterFormatForParagraphMark.RightToLeft)) { return false; }
            if ((WholeParagraphScaling.Count != 0) & !WholeParagraphScaling.Contains(paragraph.CharacterFormatForParagraphMark.Scaling)) { return false; }
            if ((paragraph.CharacterFormatForParagraphMark.Size < WholeParagraphSizeLeftBorder) | (paragraph.CharacterFormatForParagraphMark.Size > WholeParagraphSizeRightBorder)) { return false; }
            if ((WholeParagraphSmallCaps.Count != 0) & !WholeParagraphSmallCaps.Contains(paragraph.CharacterFormatForParagraphMark.SmallCaps)) { return false; }
            if ((WholeParagraphSpacing.Count != 0) & !WholeParagraphSpacing.Contains(paragraph.CharacterFormatForParagraphMark.Spacing)) { return false; }
            if ((WholeParagraphStrikethrough.Count != 0) & !WholeParagraphStrikethrough.Contains(paragraph.CharacterFormatForParagraphMark.Strikethrough)) { return false; }
            if ((WholeParagraphSubscript.Count != 0) & !WholeParagraphSubscript.Contains(paragraph.CharacterFormatForParagraphMark.Subscript)) { return false; }
            if ((WholeParagraphSuperscript.Count != 0) & !WholeParagraphSuperscript.Contains(paragraph.CharacterFormatForParagraphMark.Superscript)) { return false; }
            if ((WholeParagraphUnderlineStyle.Count != 0) & !WholeParagraphUnderlineStyle.Contains(paragraph.CharacterFormatForParagraphMark.UnderlineStyle)) { return false; }

            return true;
        }

        // Вернуть true, если параграф похож на занчение из пресета
        public bool ParagraphLooksLikePreset(Word.Paragraph paragraph)
        {
            if (!CheckParagraphFormat(paragraph)) { return false; }
            if (!CheckWholeParagraphCharacterFormat(paragraph)) { return false; }

            return true;
        }
    }
}
