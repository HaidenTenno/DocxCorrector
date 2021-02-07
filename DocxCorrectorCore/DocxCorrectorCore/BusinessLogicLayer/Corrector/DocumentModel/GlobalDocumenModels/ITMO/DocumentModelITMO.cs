using System;
using System.Collections.Generic;
using System.Text;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public class DocumentModelITMO : GlobalDocumentModel
    {
        public override ParagraphFormattingModel ParagraphFormattingModel => new ParagraphFormattingModelITMO();
    }
}
