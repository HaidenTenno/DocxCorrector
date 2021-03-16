using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocxCorrectorCore.Services.Helpers;
using DocxCorrectorCore.Models.Corrections;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
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

    public class ListElementGOST_7_32 : DocumentElementGOST_7_32, IRegexSupportable
    {
        //d0

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.d0;

        // Свойства ParagraphFormat

        // Свойства CharacterFormat для всего абзаца

        // Свойства CharacterFormat для всего абзаца

        // Особые свойства

        // Свойства списка
        // Маркер
        // TODO: !!!
        public virtual List<string> MarkerFormats => new List<string> { "—", "–", "−", "-", "%1)", "%1" };

        // Последний символ
        public virtual char LastSymbol => ',';

        // IRegexSupportable
        public virtual List<Regex> Regexes => throw new NotImplementedException();

        // Проверка последнего символа
        private ParagraphMistake? CheckLastSymbol(Word.Paragraph paragraph)
        {
            char lastSymbol;
            string paragraphContent = GemBoxHelper.GetParagraphContentWithoutNewLine(paragraph);
            try { lastSymbol = paragraphContent.Last(); } catch { return null; }

            if (lastSymbol != LastSymbol)
            {
                return new ParagraphMistake(
                    message: "Неверный последний символ"
                );
            }

            return null;
        }

        // Проверить свойства списка
        protected List<ParagraphMistake> CheckListFormat(Word.Paragraph paragraph)
        {
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            // Если элемент создан средствами Word
            if (paragraph.ListFormat.IsList)
            {
                if (!MarkerFormats.Contains(paragraph.ListFormat.ListLevelFormat.NumberFormat))
                {
                    ParagraphMistake mistake = new ParagraphMistake(
                        message: $"Неверный формат маркера"
                    );
                    paragraphMistakes.Add(mistake);
                }
            }
            // Если параграф не элемент списка
            else
            {
                ParsedListElement parsedListElement = new ParsedListElement(paragraph);

                // TODO: !!!
                ParagraphMistake markerParagraphMistake = new ParagraphMistake(
                    message: "Невозможно определить правильность формата маркера",
                    advice: "Попробуйте создать список средствами Word"
                );
                paragraphMistakes.Add(markerParagraphMistake);
            }

            // Проверка последнего символа
            ParagraphMistake? lastSymbolMistake = CheckLastSymbol(paragraph);
            if (lastSymbolMistake != null) { paragraphMistakes.Add(lastSymbolMistake); }

            return paragraphMistakes;
        }

        // Метод проверки
        public override ParagraphCorrections? CheckFormatting(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            Word.Paragraph paragraph;
            try { paragraph = (Word.Paragraph)classifiedParagraphs[id].Element; } catch { return null; }

            ParagraphCorrections? result = base.CheckFormatting(id, classifiedParagraphs);
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            // Особые свойства
            // Свойства списка
            List<ParagraphMistake> listMistakes = CheckListFormat(paragraph);
            paragraphMistakes.AddRange(listMistakes);


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

        // Выполнить сравнение по свойствам (не включая особые) с параграфом paragraph
        public override ParagraphCorrections? CheckSingleParagraphFormatting(int id, Word.Paragraph paragraph)
        {
            ParagraphCorrections? result = base.CheckSingleParagraphFormatting(id, paragraph);

            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            // Свойства списка
            List<ParagraphMistake> listMistakes = CheckListFormat(paragraph);
            paragraphMistakes.AddRange(listMistakes);

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
