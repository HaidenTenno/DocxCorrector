using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public class SourcesListElementGOST_7_0_11: ListElementGOST_7_0_11
    {
        // r0

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.r0;

        // Свойства ParagraphFormat

        // Свойства CharacterFormat для всего абзаца

        // Свойства CharacterFormat для всего абзаца

        // Особые свойства

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


        public SourcesListMistake? CheckSourcesListElement(int id, List<Regex> regexes, Word.Element sourcesListElement)
        {
            Word.Paragraph sourcesListElementParagraph;
            try { sourcesListElementParagraph = (Word.Paragraph)sourcesListElement; }
            catch
            {
                return new SourcesListMistake(
                    paragraphID: id,
                    prefix: "TABLE",
                    message: $"В списке литературы не может стоять таблица"
                );
            }

            ParsedListElement parsedListElement = new ParsedListElement(sourcesListElementParagraph);

            foreach (Regex regex in regexes)
            {
                if (regex.IsMatch(parsedListElement.Content))
                {
                    return null;
                }
            }

            return new SourcesListMistake(
                paragraphID: id,
                prefix: GemBoxHelper.GetParagraphPrefix(sourcesListElementParagraph, 20),
                message: "Элемент списка литературы не соответствует ни одному из шаблонов"
            );
        }
    }
}