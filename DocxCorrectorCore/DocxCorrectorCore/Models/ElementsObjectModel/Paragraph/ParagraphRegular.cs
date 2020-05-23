using GemBox.Document;

namespace DocxCorrectorCore.Models.ElementsObjectModel
{
    public class ParagraphRegular : DocumentElement
    {
        //c1
        
        // Свойства ParagraphFormat
        public override HorizontalAlignment Alignment => HorizontalAlignment.Justify;
        public override bool KeepWithNext => false;
        public override OutlineLevel OutlineLevel => OutlineLevel.BodyText;
        public override bool PageBreakBefore => false;
        public override double SpecialIndentationLeftBorder => -1.3; // TODO: Проверить, что это cm
        public override double SpecialIndentationRightBorder => -1.2; // TODO: Проверить, что это cm
        
        // Свойства CharacterFormat для всего абзаца
        public override bool WholeParagraphAllCaps => false;
        public override bool WholeParagraphBold => false;
        public override bool WholeParagraphSmallCaps => false;
        
        // Свойства CharacterFormat для всего абзаца
        public override bool RunnerBold => false;
        
        // Особые свойства
        public override StartSymbolType? StartSymbol => StartSymbolType.Upper;
        public override int EmptyLinesAfter => 0;
        
        // TODO: Сделать проверку окончания абзаца регуляркой (а нужна ли она вообще?..) 
        // public override string[] Suffixes => new string[] { ".", "!", "?" }; - могут пройти ".." и др.
    }
}