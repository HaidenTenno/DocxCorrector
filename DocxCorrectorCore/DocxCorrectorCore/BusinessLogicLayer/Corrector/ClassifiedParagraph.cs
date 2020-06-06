using DocxCorrectorCore.Models.Corrections;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector
{
    public sealed class ClassifiedParagraph
    {
        // Тут может быть Paragraph либо Table
        public readonly Word.Element Element;
        public readonly ParagraphClass? ParagraphClass;

        public ClassifiedParagraph(Word.Element element, ParagraphClass? paragraphClass = null)
        {
            Element = element;
            ParagraphClass = paragraphClass;
        }
    }
}
