using System;
using System.Collections.Generic;
using System.Text;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public class DocumentModelITMO : GlobalDocumentModel
    {
        public override ParagraphFormattingModel ParagraphFormattingModel => new ParagraphFormattingModelITMO();

        public override SourcesListFormattingModel SourcesListFormattingModel => throw new NotImplementedException();

        public override HeadlingsFormattingModel HeadlingsFormattingModel => throw new NotImplementedException();
    }
}
