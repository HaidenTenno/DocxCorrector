using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocxCorrectorCore.Services.Helpers;
using DocxCorrectorCore.Models.Corrections;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel
{
    public sealed class ParsedListElement
    {
        public string Marker { get; }
        public string Body { get; }

        public string Content
        {
            get
            {
                return string.Join(" ", new string[] { Marker, Body });
            }
        }

        public ParsedListElement(Word.Paragraph paragraph)
        {
            if (paragraph.ListFormat.IsList)
            {
                Marker = paragraph.ListItem.ToString();
                Body = GemBoxHelper.GetParagraphContentWithoutNewLine(paragraph);
            }
            else
            {
                string content = GemBoxHelper.GetParagraphContentWithoutNewLine(paragraph);
                List<string> words = content.Split(' ').ToList();
                try { Marker = words[0]; } catch { Marker = ""; }
                words.RemoveAt(0);
                Body = string.Join(" ", words.ToArray());
            }
        }
    }

    public class ListElement : DocumentElement, IRegexSupportable
    {
        //d0

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.d0;

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


        public virtual List<Regex> Regexes => throw new NotImplementedException();
    }
}
