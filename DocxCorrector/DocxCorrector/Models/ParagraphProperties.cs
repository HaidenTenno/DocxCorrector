using Word = Microsoft.Office.Interop.Word;

namespace DocxCorrector.Models
{
    public class ParagraphProperties
    {
        // Range
        public string Text { get; set; }
        public string FontName { get; set; }
        public string FontSize { get; set; }
        public string Bold { get; set; }
        public string Italic { get; set; }
        public string FontTextColorRGB { get; set; }
        public string FontUnderlineColor { get; set; }
        public string Underline { get; set; }
        public string FontStrikeThrough { get; set; }
        public string FontSuperscript { get; set; }
        public string FontSubscript { get; set; }
        public string FontHidden { get; set; }
        public string FontScaling { get; set; }
        public string FontPosition { get; set; }
        public string FontKerning { get; set; }
        // Paragraph
        public string OutlineLevel { get; set; }
        public string Alignment { get; set; }
        public string CharacterUnitLeftIndent { get; set; }
        public string LeftIndent { get; set; }
        public string CharacterUnitRightIndent { get; set; }
        public string RightIndent { get; set; }
        public string CharacterUnitFirstLineIndent { get; set; }
        public string MirrorIndents { get; set; }
        public string LineSpacing { get; set; }
        public string SpaceBefore { get; set; }
        public string SpaceAfter { get; set; }
        public string PageBreakBefore { get; set; }

    }
}