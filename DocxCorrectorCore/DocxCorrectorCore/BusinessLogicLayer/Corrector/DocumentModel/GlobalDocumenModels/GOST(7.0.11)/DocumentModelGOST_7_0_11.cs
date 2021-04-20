using System;
using System.Collections.Generic;
using System.Text;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public class DocumentModelGOST_7_0_11 : GlobalDocumentModel
    {
        public override ParagraphFormattingModel ParagraphFormattingModel => new ParagraphFormattingModelGOST_7_0_11();
    }
}
