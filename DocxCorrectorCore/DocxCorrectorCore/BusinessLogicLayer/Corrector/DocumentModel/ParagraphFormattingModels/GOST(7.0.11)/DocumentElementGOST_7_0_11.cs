using System;
using System.Collections.Generic;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public abstract class DocumentElementGOST_7_0_11 : DocumentElement
    {
        // Свойства ParagraphFormat
        public override double SpecialIndentationLeftBorder => -36.85;
        public override double SpecialIndentationRightBorder => -34.00;

        // Свойства CharacterFormat для всего абзаца
        public override List<string> WholeParagraphFontName => new List<string> { };
        public override double WholeParagraphSizeLeftBorder => 12;
        public override double WholeParagraphSizeRightBorder => 14;

        // Свойства CharacterFormat для раннеров
        
        // Количество пустых строк (отбивок, SPACE, n0) до или после параграфа
        public override List<int> EmptyLinesBefore => new List<int> { 0 };
        public override List<int> EmptyLinesAfter => new List<int> { 0 };
    }
}
