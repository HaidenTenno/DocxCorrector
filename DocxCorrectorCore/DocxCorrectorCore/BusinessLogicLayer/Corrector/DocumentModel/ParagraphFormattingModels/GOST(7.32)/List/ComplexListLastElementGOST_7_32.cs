using System.Collections.Generic;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;


namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public class ComplexListLastElementGOST_7_32 : ListElementGOST_7_32
    {
        //d6

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.d6;

        // Свойства ParagraphFormat

        // Свойства CharacterFormat для всего абзаца

        // Свойства CharacterFormat для всего абзаца

        // Особые свойства
        // Последний символ
        public override char LastSymbol => '.';

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