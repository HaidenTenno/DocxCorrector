using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public class TableSignGOST_7_32 : DocumentElementGOST_7_32, IRegexSupportable
    {
        //f0

        // Класс элемента
        public override ParagraphClass ParagraphClass => paragraphClass;
        private readonly ParagraphClass paragraphClass;
        public override List<bool> KeepLinesTogether => new List<bool> { true };
        public override List<bool> KeepWithNext => new List<bool> { true };
        public override List<double> LineSpacing => new List<double> { 1.0, 1.5 };
        public override double SpecialIndentationLeftBorder => 0;
        public override double SpecialIndentationRightBorder => 0;

        // Свойства ParagraphFormat

        // Свойства CharacterFormat для всего абзаца

        // Свойства CharacterFormat для всего абзаца

        // Особые свойства

        // IRegexSupportable
        public List<Regex> Regexes => ParagraphClass switch
        {
            ParagraphClass.f1 => new List<Regex> { new Regex(@"^Таблица (?>[А-ЕЖИК-НП-ЦШЩЭЮЯ]\.[\d]+|[\d]+(?>\.[\d]+)?)(?> - .*)?") },
            ParagraphClass.f3 => new List<Regex> { new Regex(@"^Таблица (?>[А-ЕЖИК-НП-ЦШЩЭЮЯ]\.[\d]+|[\d]+(?>\.[\d]+)?)(?> - .*)?") },
            ParagraphClass.f5 => new List<Regex> { new Regex(@"^Таблица (?>[А-ЕЖИК-НП-ЦШЩЭЮЯ]\.[\d]+|[\d]+(?>\.[\d]+)?)(?> - .*)?") },
            _ => throw new ArgumentException(message: "invalid paragraph class", paramName: nameof(ParagraphClass))
        };

        public TableSignGOST_7_32(ParagraphClass paragraphClass)
        {
            this.paragraphClass = paragraphClass;
        }

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
                message: "Запись подписи к таблице не соответствует ни одному из шаблонов"
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

        public override ParagraphCorrections? CheckSingleParagraphFormatting(int id, Word.Paragraph paragraph)
        {
            ParagraphCorrections? result = base.CheckSingleParagraphFormatting(id, paragraph);

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
