using System;
using System.Collections.Generic;
using System.Linq;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace DocxCorrector.Models
{
    public sealed class ParagraphPropertiesSpire : ParagraphProperties
    {
        public string Text { get; }
        public int WordCount { get; }
        public string StyleName { get; }

        // Properties common for paragraph
        public bool NoBorders { get; }
        public float AfterSpacing { get; }
        public bool IsEmptyBackgroundColor { get; }
        public float BeforeSpacing { get; }
        public string HorizontalAlignment { get; }
        public bool IsBidi { get; }
        public bool IsFrame { get; }
        public bool KeepFollow { get; }
        public bool KeepLines { get; }
        public float LeftIndent { get; }
        public float LineSpacing { get; }
        public bool MirrorIndents { get; }
        public string OutlineLevel { get; }
        public bool OverflowPunctuation { get; }
        public float RightIndent { get; }
        public string TextAlignment { get; }
        public bool WordWrap { get; }
        public bool AfterAutoSpacing { get; }
        public bool BeforeAutoSpacing { get; }
        public float FirstLineIndent { get; }
        public bool IsKinSoku { get; }
        public bool IsWidowControl { get; }
        public string LineSpacingRule { get; }
        public bool PageBreakAfter { get; }
        public bool PageBreakBefore { get; }
        public bool AutoSpaceDE { get; }
        public bool AutoSpaceDN { get; }
        public bool IsColumnBreakAfter { get; }

        // Ranges
        public List<Dictionary<string, string>> TextRangesProperties { get; }

        // TODO: понять, что вытаскивают поля, отмеченные "??"
        public ParagraphPropertiesSpire(Paragraph paragraph)
        {
            Text = paragraph.Text;
            WordCount = paragraph.WordCount;
            StyleName = paragraph.StyleName;
            NoBorders = paragraph.Format.Borders.NoBorder;
            AfterSpacing = paragraph.Format.AfterSpacing;
            IsEmptyBackgroundColor = paragraph.Format.BackColor.IsEmpty;
            BeforeSpacing = paragraph.Format.BeforeSpacing;
            HorizontalAlignment = paragraph.Format.HorizontalAlignment.ToString();
            IsBidi = paragraph.Format.IsBidi;
            IsFrame = paragraph.Format.IsFrame; // ??
            KeepFollow = paragraph.Format.KeepFollow;
            KeepLines = paragraph.Format.KeepLines;
            LeftIndent = paragraph.Format.LeftIndent;
            LineSpacing = paragraph.Format.LineSpacing;
            MirrorIndents = paragraph.Format.MirrorIndents;
            OutlineLevel = paragraph.Format.OutlineLevel.ToString();
            OverflowPunctuation = paragraph.Format.OverflowPunc; // ??
            RightIndent = paragraph.Format.RightIndent;
            TextAlignment = paragraph.Format.TextAlignment.ToString();
            WordWrap = paragraph.Format.WordWrap;
            AfterAutoSpacing = paragraph.Format.AfterAutoSpacing; // ??
            BeforeAutoSpacing = paragraph.Format.BeforeAutoSpacing; // ??
            FirstLineIndent = paragraph.Format.FirstLineIndent;
            IsKinSoku = paragraph.Format.IsKinSoku;
            IsWidowControl = paragraph.Format.IsWidowControl; // ??
            LineSpacingRule = paragraph.Format.LineSpacingRule.ToString(); // ??
            PageBreakAfter = paragraph.Format.PageBreakAfter;
            PageBreakBefore = paragraph.Format.PageBreakBefore;
            AutoSpaceDE = paragraph.Format.AutoSpaceDE; // ??
            AutoSpaceDN = paragraph.Format.AutoSpaceDN; // ??
            IsColumnBreakAfter = paragraph.Format.IsColumnBreakAfter;
            // Runners
            TextRangesProperties = new List<Dictionary<string, string>>();
            foreach (TextRange textRange in paragraph.ChildObjects.OfType<TextRange>())
            {
                var textRangeProperty = new Dictionary<string, string>()
                {
                    { "Text", textRange.Text },
                    { "\nIsBidi", textRange.CharacterFormat.Bidi.ToString() },
                    { "\nIsBold", textRange.CharacterFormat.Bold.ToString() },
                    { "\nHasBorder", (textRange.CharacterFormat.Border.LineWidth == 0.0).ToString() },
                    { "\nIsEmbossed", textRange.CharacterFormat.Emboss.ToString() },
                    { "\nIsEngraved", textRange.CharacterFormat.Engrave.ToString() },
                    { "\nIsHidden", textRange.CharacterFormat.Hidden.ToString() },
                    { "\nIsItalic", textRange.CharacterFormat.Italic.ToString() },
                    { "\nPosition", textRange.CharacterFormat.Position.ToString() }, // ??
                    { "\nIsBigCaps", textRange.CharacterFormat.AllCaps.ToString() },
                    { "\nCharSpacing", textRange.CharacterFormat.CharacterSpacing.ToString() },
                    { "\nIsDoubleStriked", textRange.CharacterFormat.DoubleStrike.ToString() },
                    { "\nHasEmphasisMark", (textRange.CharacterFormat.EmphasisMark.ToString() != "None").ToString() },
                    { "\nFontName", textRange.CharacterFormat.FontName },
                    { "\nFontSize", textRange.CharacterFormat.FontSize.ToString() },
                    { "\nHasUnusualHiglightColor", (!textRange.CharacterFormat.HighlightColor.IsEmpty).ToString() },
                    { "\nIsShadow", textRange.CharacterFormat.IsShadow.ToString() },
                    { "\nIsStrikeout", textRange.CharacterFormat.IsStrikeout.ToString() },
                    { "\nHasLigaturesType", (textRange.CharacterFormat.LigaturesType.ToString() != "None").ToString() },
                    { "\nHasUnusualTextColor", (!textRange.CharacterFormat.TextColor.IsEmpty).ToString() },
                    { "\nTextScale", textRange.CharacterFormat.TextScale.ToString() },
                    { "\nHasUnderline", (textRange.CharacterFormat.UnderlineStyle.ToString() != "None").ToString() },
                    { "\nAllowContextualAlternates", textRange.CharacterFormat.AllowContextualAlternates.ToString() }, // ??
                    { "\nFontTypeHint", textRange.CharacterFormat.FontTypeHint.ToString() }, // ??
                    { "\nIsOutLine", textRange.CharacterFormat.IsOutLine.ToString() },
                    { "\nIsSmallCaps", textRange.CharacterFormat.IsSmallCaps.ToString() },
                    { "\nNumberFormType", textRange.CharacterFormat.NumberFormType.ToString() }, // ??
                    { "\nNumberSpaceType", textRange.CharacterFormat.NumberSpaceType.ToString() }, // ??
                    { "\nStylisticSetType", textRange.CharacterFormat.StylisticSetType.ToString() }, // ??
                    { "\nSubSuperScript", textRange.CharacterFormat.SubSuperScript.ToString() },
                    { "\nHasUnusualBackgraoundColor", (!textRange.CharacterFormat.TextBackgroundColor.IsEmpty).ToString() }
                };
                TextRangesProperties.Add(textRangeProperty);
            }
        }
    }
}
