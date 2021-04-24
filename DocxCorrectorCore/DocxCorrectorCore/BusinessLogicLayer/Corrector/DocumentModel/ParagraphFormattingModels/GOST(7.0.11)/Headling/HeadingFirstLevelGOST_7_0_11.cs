using System.Collections.Generic;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public class HeadingFirstLevelGOST_7_0_11 : HeadingGOST_7_0_11
    {
        //b1

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.b1;

        // Свойства ParagraphFormat
        public override List<Word.HorizontalAlignment> Alignment => new List<Word.HorizontalAlignment> { Word.HorizontalAlignment.Center };
        public override List<bool> KeepWithNext => new List<bool> { false };
        public override List<Word.OutlineLevel> OutlineLevel => new List<Word.OutlineLevel> { Word.OutlineLevel.Level1 };
        public override List<bool> PageBreakBefore => new List<bool> { true };
        public override double SpecialIndentationLeftBorder => 0;
        public override double SpecialIndentationRightBorder => 0;

        // Свойства CharacterFormat для всего абзаца
        public override List<bool> WholeParagraphAllCaps => new List<bool> { true };
        public override List<bool> WholeParagraphBold => new List<bool> { true };

        // Свойства CharacterFormat для всего абзаца

        // Особые свойства
        public override List<int> EmptyLinesBefore => new List<int> { 3 };
        public override List<int> EmptyLinesAfter => new List<int> { 3 };

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