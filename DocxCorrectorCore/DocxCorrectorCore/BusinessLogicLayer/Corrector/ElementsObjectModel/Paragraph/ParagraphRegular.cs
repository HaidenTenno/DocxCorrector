using System.Collections.Generic;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;


namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel
{
    public class ParagraphRegular : DocumentElement
    {
        //c1

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.c1;

        // Свойства ParagraphFormat

        // Свойства CharacterFormat для всего абзаца

        // Свойства CharacterFormat для всего абзаца

        // Особые свойства

        // Проверка первого симола
        // TODO: Переписать для Enum
        private ParagraphMistake? CheckStartSymbol(Word.Paragraph paragraph)
        {
            char firstSymbol;
            try { firstSymbol = paragraph.Content.ToString()[0]; } catch { return null; }

            if ((firstSymbol != '"') & (!char.IsUpper(firstSymbol)))
            {
                return new ParagraphMistake(
                    message: "Параграф должен начинаться с большой буквы",
                    advice: "ТУТ БУДЕТ СОВЕТ"
                );
            }

            return null;
        }

        // Метод проверки
        public override ParagraphCorrections? CheckFormatting(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            Word.Paragraph paragraph;
            try { paragraph = (Word.Paragraph)classifiedParagraphs[id].Element; } catch { return null; }

            ParagraphCorrections? result = base.CheckFormatting(id, classifiedParagraphs);
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            // Особые свойства
            // Проверка первого символа
            ParagraphMistake? startSymbolMistake = CheckStartSymbol(paragraph);
            if (startSymbolMistake != null) { paragraphMistakes.Add(startSymbolMistake); }

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