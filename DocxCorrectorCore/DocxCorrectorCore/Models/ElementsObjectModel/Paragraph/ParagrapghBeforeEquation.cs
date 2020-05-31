using GemBox.Document;

namespace DocxCorrectorCore.Models.ElementsObjectModel
{
    public class ParagrapghBeforeEquation : DocumentElement
    {
        //c3

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.c3;

        // Свойства ParagraphFormat
        public override HorizontalAlignment Alignment => HorizontalAlignment.Justify;
        public override bool KeepWithNext => true;
        public override OutlineLevel OutlineLevel => OutlineLevel.BodyText;
        public override bool PageBreakBefore => false;
        public override double SpecialIndentationLeftBorder => -36.85;
        public override double SpecialIndentationRightBorder => -35.45;
        
        // Свойства CharacterFormat для всего абзаца
        public override bool WholeParagraphAllCaps => false;
        public override bool WholeParagraphBold => false;
        public override bool WholeParagraphSmallCaps => false;
        
        // Свойства CharacterFormat для всего абзаца
        public override bool RunnerBold => false;
        
        // Особые свойства
        //public override StartSymbolType? StartSymbol => StartSymbolType.Upper;
        public override int EmptyLinesAfter => 0;
    }
}