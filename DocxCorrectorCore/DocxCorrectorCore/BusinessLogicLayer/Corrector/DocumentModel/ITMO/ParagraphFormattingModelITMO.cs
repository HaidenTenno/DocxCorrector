using System;
using System.Collections.Generic;
using System.Text;
using DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel;
using DocxCorrectorCore.Models.Corrections;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel.ITMO
{
    public class ParagraphFormattingModelITMO : ParagraphFormattingModel
    {
        public override DocumentElement? GetDocumentElementFromClass(ParagraphClass paragraphClass)
        {
            return null;
        }
    }
}
