using System;
using Word = GemBox.Document;
using DocxCorrectorCore.Services.Helpers;

namespace DocxCorrectorCore.BusinessLogicLayer.PropertiesPuller
{
    // TODO: Наиль должен скинуть окончательный вид
    public sealed class ParagraphPropertiesGemBox : ParagraphProperties
    {
        public string Content { get; }
        public string? LastSymbolPd { get;  }
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
        public string? SpaceAfter { get; }
        public string? SpaceBefore { get; }
        public string? SpecialIndentation { get; }

        // Private
        private string GetProperContent(Word.Paragraph paragraph)
        {
            string result = "";

            foreach (var element in paragraph.GetChildElements(false, new Word.ElementType[] { Word.ElementType.Run, Word.ElementType.Picture, Word.ElementType.Chart, Word.ElementType.Shape, Word.ElementType.PreservedInline }))
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
                    case Word.PreservedInline _:
                        result += "PRESERVED INLINE ";
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
            LastSymbolPd = GemBoxHelper.CheckIfLastSymbolOfParagraphIsOneOf(paragraph, new string[] { ".", ",", ":", ";", "!" });
            FirstKeyWord = GemBoxHelper.CheckIfFirtWordOfParagraphIsOneOf(paragraph, new string[] { "Таблица", "Табл", "Рисунок", "Рис.", "Рис", "Табл." });
            // Свойства символов всего параграфа
            FullBold = paragraph.CharacterFormatForParagraphMark.Bold.ToString();
            FullItalic = paragraph.CharacterFormatForParagraphMark.Italic.ToString();
            // Свойства параграфа
            Alignment = paragraph.ParagraphFormat.Alignment.ToString();
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
            SpaceAfter = paragraph.ParagraphFormat.SpaceAfter.ToString();
            SpaceBefore = paragraph.ParagraphFormat.SpaceBefore.ToString();
            SpecialIndentation = paragraph.ParagraphFormat.SpecialIndentation.ToString();
        }

        // PlaceHolder constructor
        public ParagraphPropertiesGemBox(string placeHolder)
        {
            Content = placeHolder;
        }       
    }
}
