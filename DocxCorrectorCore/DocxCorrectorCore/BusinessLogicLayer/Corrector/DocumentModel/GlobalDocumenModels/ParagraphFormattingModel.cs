using System;
using System.Collections.Generic;
using System.Text;
using DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel;
using DocxCorrectorCore.Models.Corrections;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public abstract class ParagraphFormattingModel
    {
        public abstract DocumentElement? GetDocumentElementFromClass(ParagraphClass paragraphClass);
    }
}
