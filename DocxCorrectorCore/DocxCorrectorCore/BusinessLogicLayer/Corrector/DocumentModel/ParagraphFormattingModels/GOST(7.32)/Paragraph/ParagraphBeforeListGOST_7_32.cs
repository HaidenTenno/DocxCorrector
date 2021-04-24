using System.Collections.Generic;
using System.Linq;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public class ParagraphBeforeListGOST_7_32 : ParagraphRegularGOST_7_32
    {
        //c2

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.c2;

        // Свойства ParagraphFormat
        public override List<bool> KeepLinesTogether => new List<bool> { true };

        // Свойства CharacterFormat для всего абзаца

        // Свойства CharacterFormat для всего абзаца

        // Особые свойства
        public override List<EdgeSymbolType> StartSymbolType => new List<EdgeSymbolType> { EdgeSymbolType.CapitalLetter };
        public override List<EdgeSymbolType> LastSymbolType => new List<EdgeSymbolType> { EdgeSymbolType.Colon };

        // Метод проверки
        public override ParagraphCorrections? CheckFormatting(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            Word.Paragraph paragraph;
            try { paragraph = (Word.Paragraph)classifiedParagraphs[id].Element; } catch { return null; }

            ParagraphCorrections? result = base.CheckFormatting(id, classifiedParagraphs);
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            // Особые свойства

            if (paragraphMistakes.Count != 0)
            {
                if (result != null)
                {
                    result.Mistakes.AddRange(paragraphMistakes);
                }
                else
                {
                    result = new ParagraphCorrections(
                        paragraphID: id,
                        paragraphClass: ParagraphClass,
                        prefix: GemBoxHelper.GetParagraphPrefix(paragraph, 20),
                        mistakes: paragraphMistakes
                    );
                }
            }

            return result;
        }
    }
}
