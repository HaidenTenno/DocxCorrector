using System;
using System.Collections.Generic;
using System.Text;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public abstract class GlobalDocumentModel
    {
        public abstract ParagraphFormattingModel ParagraphFormattingModel { get; }
    }
}
