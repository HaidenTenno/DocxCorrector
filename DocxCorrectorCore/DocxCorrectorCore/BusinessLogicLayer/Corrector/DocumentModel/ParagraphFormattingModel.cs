using System;
using System.Collections.Generic;
using System.Text;
using DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel;
using DocxCorrectorCore.Models.Corrections;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector
{
    public abstract class ParagraphFormattingModel
    {
        public abstract DocumentElement? GetDocumentElementFromClass(ParagraphClass paragraphClass);
    }
}
