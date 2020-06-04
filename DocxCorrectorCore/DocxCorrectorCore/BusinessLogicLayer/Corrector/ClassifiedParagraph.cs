using DocxCorrectorCore.Models.Corrections;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector
{
    public sealed class ClassifiedParagraph
    {
        public readonly Word.Paragraph Paragraph;
        public readonly ParagraphClass? ParagraphClass;

        public ClassifiedParagraph(Word.Paragraph paragraph, ParagraphClass? paragraphClass = null)
        {
            Paragraph = paragraph;
            ParagraphClass = paragraphClass;
        }
    }
}
