using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel
{
    public class Heading : DocumentElement
    {
        //b0

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.b0;

        // Свойства ParagraphFormat

        // Свойства CharacterFormat для всего абзаца

        // Свойства CharacterFormat для всего абзаца

        // Особые свойства

    }

    public class DocumentHeadlings
    {
        public HeadlingCorrections? CheckHeadlings(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            // TODO: Продолжить тут

            return null;
        }
    }
}
