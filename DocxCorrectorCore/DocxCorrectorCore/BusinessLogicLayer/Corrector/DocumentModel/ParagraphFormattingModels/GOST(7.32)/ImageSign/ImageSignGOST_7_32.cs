using System.Collections.Generic;
using System.Text.RegularExpressions;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public class ImageSignGOST_7_32 : DocumentElementGOST_7_32, IRegexSupportable
    {
        //h1

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.h1;

        // Свойства ParagraphFormat
        public override List<Word.HorizontalAlignment> Alignment => new List<Word.HorizontalAlignment> { Word.HorizontalAlignment.Center };
        public override List<bool> KeepLinesTogether => new List<bool> { true };
        public override double SpecialIndentationLeftBorder => 0;
        public override double SpecialIndentationRightBorder => 0;


        // Свойства CharacterFormat для всего абзаца

        // Свойства CharacterFormat для всего абзаца

        // Особые свойства

        // IRegexSupportable
        public List<Regex> Regexes => new List<Regex> 
        { 
            new Regex (@"^Рисунок (?>[А-ЕЖИК-НП-ЦШЩЭЮЯ]\.[\d]+|[\d]+(?>\.[\d]+)?)(?> - .*)?$") 
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
                message: "Запись подписи к рисунку не соответствует ни одному из шаблонов"
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
            // Проверка соответствия шаблону
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
