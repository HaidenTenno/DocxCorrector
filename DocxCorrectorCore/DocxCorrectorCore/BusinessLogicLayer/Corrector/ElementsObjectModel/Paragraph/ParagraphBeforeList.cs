using System.Collections.Generic;
using System.Linq;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel
{
    public class ParagraphBeforeList : ParagraphRegular
    {
        //c2

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.c2;

        // Свойства ParagraphFormat
        public override List<bool> KeepLinesTogether => new List<bool> { true };

        // Свойства CharacterFormat для всего абзаца

        // Свойства CharacterFormat для всего абзаца

        // Особые свойства

        // Проверка первого симола
        // TODO: Переписать для Enum
        private ParagraphMistake? CheckLastSymbol(Word.Paragraph paragraph)
        {
            char lastSymbol;
            string paragraphContent = GemBoxHelper.GetParagraphContentWithoutNewLine(paragraph);
            try { lastSymbol = paragraphContent.Last(); } catch { return null; }

            if ((lastSymbol != ':'))
            {
                return new ParagraphMistake(
                    message: "Параграф перед списком должен заканчиваться на двоеточие"
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
            // Проверка последнего символа
            ParagraphMistake? lastSymbolMistake = CheckLastSymbol(paragraph);
            if (lastSymbolMistake != null) { paragraphMistakes.Add(lastSymbolMistake); }

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
