using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public abstract class SourcesListFormattingModel
    {
        public abstract List<string> KeyWords { get; }

        public abstract List<Regex> Regexes { get; }

        public abstract SourcesListCorrections? CheckSourcesList(int id, List<ClassifiedParagraph> classifiedParagraphs);
    }
}
