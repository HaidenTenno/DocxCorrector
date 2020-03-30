using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocxCorrector.Services;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace DocxCorrector.Models
{
    public sealed class ParagraphPropertiesSpire : ParagraphProperties
    {
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
            var textRangesProperties = new List<string>();
            foreach (TextRange textRange in paragraph.ChildObjects.OfType<TextRange>())
            {
                var textRangeProperty = new Dictionary<string, string>
                {
                    ["Text"] = textRange.Text,
                    ["IsBidi"] = textRange.CharacterFormat.Bidi.ToString(),
                    ["IsBold"] = textRange.CharacterFormat.Bold.ToString(),
                    ["HasBorder"] = (textRange.CharacterFormat.Border.LineWidth == 0.0).ToString(),
                    ["IsEmbossed"] = textRange.CharacterFormat.Emboss.ToString(),
                    ["IsEngraved"] = textRange.CharacterFormat.Engrave.ToString(),
                    ["IsHidden"] = textRange.CharacterFormat.Hidden.ToString(),
                    ["IsItalic"] = textRange.CharacterFormat.Italic.ToString(),
                    ["Position"] = textRange.CharacterFormat.Position.ToString(), // ??
                    ["IsBigCaps"] = textRange.CharacterFormat.AllCaps.ToString(),
                    ["CharSpacing"] = textRange.CharacterFormat.CharacterSpacing.ToString(),
                    ["IsDoubleStriked"] = textRange.CharacterFormat.DoubleStrike.ToString(),
                    ["HasEmphasisMark"] = (textRange.CharacterFormat.EmphasisMark.ToString() != "None").ToString(),
                    ["FontName"] = textRange.CharacterFormat.FontName,
                    ["FontSize"] = textRange.CharacterFormat.FontSize.ToString(),
                    ["HasUnusualHiglightColor"] = (!textRange.CharacterFormat.HighlightColor.IsEmpty).ToString(),
                    ["IsShadow"] = textRange.CharacterFormat.IsShadow.ToString(),
                    ["IsStrikeout"] = textRange.CharacterFormat.IsStrikeout.ToString(),
                    ["HasLigaturesType"] = (textRange.CharacterFormat.LigaturesType.ToString() != "None").ToString(),
                    ["HasUnusualTextColor"] = (!textRange.CharacterFormat.TextColor.IsEmpty).ToString(),
                    ["TextScale"] = textRange.CharacterFormat.TextScale.ToString(),
                    ["HasUnderline"] = (textRange.CharacterFormat.UnderlineStyle.ToString() != "None").ToString(),
                    ["AllowContextualAlternates"] = textRange.CharacterFormat.AllowContextualAlternates.ToString(), // ??
                    ["FontTypeHint"] = textRange.CharacterFormat.FontTypeHint.ToString(), // ??
                    ["IsOutLine"] = textRange.CharacterFormat.IsOutLine.ToString(),
                    ["IsSmallCaps"] = textRange.CharacterFormat.IsSmallCaps.ToString(),
                    ["NumberFormType"] = textRange.CharacterFormat.NumberFormType.ToString(), // ??
                    ["NumberSpaceType"] = textRange.CharacterFormat.NumberSpaceType.ToString(), // ??
                    ["StylisticSetType"] = textRange.CharacterFormat.StylisticSetType.ToString(), // ??
                    ["SubSuperScript"] = textRange.CharacterFormat.SubSuperScript.ToString(),
                    ["HasUnusualBackgraoundColor"] = (!textRange.CharacterFormat.TextBackgroundColor.IsEmpty).ToString()
                };
                textRangesProperties.Add(JSONMaker.MakeJSON(textRangeProperty));
            }
            TextRangesProperties = String.Join(",", textRangesProperties);
        }
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

        // Property within all text ranges properties
        public string TextRangesProperties { get; }
    }
}
