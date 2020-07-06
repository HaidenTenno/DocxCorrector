using System.Collections.Generic;
using System.Text.RegularExpressions;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel
{
    public class TableSign : DocumentElement, IRegexSupportable
    {
        // TODO: Обратить внимание (пока класс f0)
        //f0

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.f0;
        public override bool KeepLinesTogether => true;
        public override bool KeepWithNext => true;

        // Свойства ParagraphFormat

        // Свойства CharacterFormat для всего абзаца

        // Свойства CharacterFormat для всего абзаца

        // Особые свойства

        // IRegexSupportable
        public List<Regex> Regexes => new List<Regex>
        {
            new Regex (@"^Таблица (?>[А-ЕЖИК-НП-ЦШЩЭЮЯ]\.[\d]+|[\d]+(?>\.[\d]+)?)(?> - .*)?")
        };

        private ParagraphMistake? CheckRegexMatch(Word.Paragraph paragraph)
        {
            string paragraphContent = GemBoxHelper.GetParagraphContentWithoutNewLine(paragraph);
            foreach (Regex regex in Regexes)
            {
                if (regex.IsMatch(paragraphContent))
                {
                    return null;
                }
            }

            return new ParagraphMistake(
                message: "Запись подписи к таблице не соответствует ни одному из шаблонов",
                advice: "ТУТ БУДЕТ СОВЕТ"
            );
        }

        // Метод проверки
        public override ParagraphCorrections? CheckFormatting(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            Word.Paragraph paragraph;
            try { paragraph = (Word.Paragraph)classifiedParagraphs[id].Element; } catch { return null; }

            ParagraphCorrections? result = base.CheckFormatting(id, classifiedParagraphs);
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            // Особые свойства
            ParagraphMistake? regexMistake = CheckRegexMatch(paragraph);
            if (regexMistake != null) { paragraphMistakes.Add(regexMistake); }

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
