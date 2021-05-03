using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public abstract class HeadlingsFormattingModel
    {
        public HeadlingCorrections? CheckHeadlings(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            // TODO: Продолжить тут

            return null;
        }
    }
}
