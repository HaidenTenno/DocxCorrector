using System.Collections.Generic;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel
{
    public class Heading1stLevel : Heading
    {
        //b1

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.b1;

        // Свойства ParagraphFormat
        public override Word.HorizontalAlignment Alignment => Word.HorizontalAlignment.Center;
        public override bool KeepWithNext => false;
        public override Word.OutlineLevel OutlineLevel => Word.OutlineLevel.Level1;
        public override bool PageBreakBefore => true;
        public override double SpecialIndentationLeftBorder => 0;
        public override double SpecialIndentationRightBorder => 0;

        // Свойства CharacterFormat для всего абзаца
        public override bool? WholeParagraphAllCaps => true;
        public override bool WholeParagraphBold => true;

        // Свойства CharacterFormat для всего абзаца

        // Особые свойства
        public override int EmptyLinesAfter => 1;

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