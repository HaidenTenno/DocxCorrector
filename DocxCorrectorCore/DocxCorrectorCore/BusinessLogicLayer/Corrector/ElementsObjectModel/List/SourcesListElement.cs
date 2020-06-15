using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel
{
    public class SourcesListElement //: DocumentElement
    {

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

    public class SourcesList
    {
        public List<string> KeyWords = new List<string> { "Список литературы", "Список источников", "Список использованных", "Список использованной" };

        public List<Regex> Regexes => new List<Regex>
        {
            new Regex(@"\d( [A-ZА-Я][a-zа-я]*\,? ([A-ZА-Я]\.){1,2}\,?){1,3} [A-ZА-Я].* (- ([A-ZА-Я]\. ?){1,3})?\: [A-ZА-Я][a-zа-я].*\, [1-9]\d{2,3}\. - [1-9]\d*( \- [1-9]\d*)? с\."),
            new Regex(@"\d ([А-ЯA-Z][a-zа-я]{1,} ([A-ZА-Я]\.\,?){1,2}){1,3} [А-ЯA-Z].* \/\/ ([А-ЯA-Z].*\: [А-ЯA-Z].*\/ ([а-яa-z].*\; [А-ЯA-Z].*(\. - [А-ЯA-Z].*\: )?\,) [1-9]\d{2,3}\. - С.[1-9]\d{1,}-\d*\.|[А-ЯA-Z].*\: [А-ЯA-Z].*\/ г.[А-ЯA-Z].*\, [(][а-я]* [1-9]\d{2,3} г\.[)]\.( - [A-ZА-Яа-яa-z]\.([1-9]\d*\.)?){0,4}\, [1-9]\d*\. - С.[1-9]\d*(\.|\-[1-9]\d*)\.|[A-ZА-Я].* \- [1-9]\d{2,3}\. - [A-Z] \d*\. - С\.[1-9]\d*(\-|\,)[1-9]\d*\.)"),
            new Regex(@"\d [A-ZА-Я].*\.(\: .*\. - [1,2]\d{3}\. | - URL: http[s]?:\/\/([a-z]*\.){1,})([a-zA-Z].[^а-яА-Я]*\/)(\([а-яА-Я ]*(0[1-9]|1[0-9]|2[0-9]|3[0-1])\.(0[1-9]|1[0-2])\.([1,2]\d{3}\)\.))?|[a-zA-Z].[^а-яА-Я]*[a-zа-яА-Я ]*(0[1-9]|1[0-9]|2[0-9]|3[0,1])\.(0[1-9]|1[0-2])\.([1,2]\d{3})\)\."),
            new Regex(@"\n\d{1,3}\.\ (([А-Я]([а-я]{1,})\ ([А-Я]([а-я]{0,1})\.\ ){1,2}))")
        };

        public SourcesListCorrections? CheckSourcesList(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            List<SourcesListMistake> sourcesListMistakes = new List<SourcesListMistake>();

            // Идем до конца документа ИЛИ пока не встретим следующий заголовок
            for (int sourcesListElementIndex = id + 1; sourcesListElementIndex < classifiedParagraphs.Count(); sourcesListElementIndex++)
            {
                if (classifiedParagraphs[sourcesListElementIndex].ParagraphClass == null) { continue; }

                if (classifiedParagraphs[sourcesListElementIndex].ParagraphClass == ParagraphClass.b1) { break; }

                var standartSourcesListElement = new SourcesListElement();

                // ПРОВЕРКА НАЧИНАЕТСЯ
                SourcesListMistake? currentSourcesListMistakes = standartSourcesListElement.CheckSourcesListElement(sourcesListElementIndex, Regexes, classifiedParagraphs[sourcesListElementIndex].Element);
                if (currentSourcesListMistakes != null) { sourcesListMistakes.Add(currentSourcesListMistakes); }
            }

            if (sourcesListMistakes.Count != 0)
            {
                return new SourcesListCorrections(
                    paragraphID: id,
                    prefix: GemBoxHelper.GetParagraphPrefix((Word.Paragraph)classifiedParagraphs[id].Element, 20),
                    mistakes: sourcesListMistakes
                );
            }
            else
            {
                return null;
            }
        }
    }
}