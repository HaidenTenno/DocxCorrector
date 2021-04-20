using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public class HeadingGOST_7_0_11 : DocumentElementGOST_7_0_11
    {
        //b0

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.b0;

        // Свойства ParagraphFormat
        public override List<bool> PageBreakBefore => new List<bool> { true, false };

        // Свойства CharacterFormat для всего абзаца
        public override List<bool> WholeParagraphAllCaps => new List<bool> { };
        public override List<bool> WholeParagraphBold => new List<bool> { };

        // Свойства CharacterFormat для всего абзаца

        // Особые свойства

    }
}
