using System.Collections.Generic;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;


namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public class ParagraphRegularGOST_7_0_11 : DocumentElementGOST_7_0_11
    {
        //c1

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.c1;

        // Свойства ParagraphFormat

        // Свойства CharacterFormat для всего абзаца

        // Свойства CharacterFormat для всего абзаца

        // Особые свойства
        public override List<EdgeSymbolType> StartSymbolType => new List<EdgeSymbolType> { EdgeSymbolType.CapitalLetter };

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