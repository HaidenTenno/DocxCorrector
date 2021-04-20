using System;
using System.Collections.Generic;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public class HeadingOtherLevelsGOST_7_0_11 : HeadingGOST_7_0_11
    {
        //b2
        //b3
        //b4

        // Класс элемента
        public override ParagraphClass ParagraphClass => paragraphClass;
        private readonly ParagraphClass paragraphClass;

        // Свойства ParagraphFormat
        public override List<Word.HorizontalAlignment> Alignment => new List<Word.HorizontalAlignment> { Word.HorizontalAlignment.Center };
        public override List<bool> KeepWithNext => new List<bool> { false };
        public override List<bool> NoSpaceBetweenParagraphsOfSameStyle => new List<bool> { true, false };
        public override List<Word.OutlineLevel> OutlineLevel => ParagraphClass switch
        {
            ParagraphClass.b2 => new List<Word.OutlineLevel> { Word.OutlineLevel.Level2 },
            ParagraphClass.b3 => new List<Word.OutlineLevel> { Word.OutlineLevel.Level3 },
            ParagraphClass.b4 => new List<Word.OutlineLevel> { Word.OutlineLevel.Level4 },
            _ => throw new ArgumentException(message: "invalid paragraph class", paramName: nameof(ParagraphClass))
        };

        // Свойства CharacterFormat для всего абзаца
        public override List<bool> WholeParagraphBold => new List<bool> { true };

        // Свойства CharacterFormat для всего абзаца

        // Особые свойства
        public override List<int> EmptyLinesBefore => new List<int> { 3 };
        public override List<int> EmptyLinesAfter => new List<int> { 3 };

        public HeadingOtherLevelsGOST_7_0_11(ParagraphClass paragraphClass)
        {
            this.paragraphClass = paragraphClass;
        }

        // Проверка первого симола
        // TODO: Переписать для Enum
        private ParagraphMistake? CheckStartSymbol(Word.Paragraph paragraph)
        {
            char firstSymbol;
            try { firstSymbol = paragraph.Content.ToString()[0]; } catch { return null; }

            if ((firstSymbol != '"') & (!char.IsUpper(firstSymbol)))
            {
                return new ParagraphMistake(
                    message: "Параграф должен начинаться с большой буквы"
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