using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public class HeadingGOST_7_32 : DocumentElementGOST_7_32
    {
        //b0

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.b0;

        // Свойства ParagraphFormat

        // Свойства CharacterFormat для всего абзаца

        // Свойства CharacterFormat для всего абзаца

        // Особые свойства
        public override List<EdgeSymbolType> StartSymbolType => new List<EdgeSymbolType> { EdgeSymbolType.CapitalLetter };
        public override List<EdgeSymbolType> LastSymbolType => new List<EdgeSymbolType> { EdgeSymbolType.SmallLetter };
    }
}
