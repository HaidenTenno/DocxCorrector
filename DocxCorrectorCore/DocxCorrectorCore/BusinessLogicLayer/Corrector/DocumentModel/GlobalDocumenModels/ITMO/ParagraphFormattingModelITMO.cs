using System;
using System.Collections.Generic;
using System.Text;
using DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel;
using DocxCorrectorCore.Models.Corrections;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public class ParagraphFormattingModelITMO : ParagraphFormattingModel
    {
        public override DocumentElement? GetDocumentElementFromClass(ParagraphClass paragraphClass)
        {
            return null;
        }
    }
}
