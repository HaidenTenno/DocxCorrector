using System;
using System.Linq;
using System.Collections.Generic;
using DocxCorrectorCore.BusinessLogicLayer.Corrector;
using Word = GemBox.Document;
using DocxCorrectorCore.Models.Corrections;

namespace DocxCorrectorCore.Services.Helpers
{
    internal static class GemBoxHelper
    {
        internal static readonly Dictionary<Word.ElementType, string> SkippableElements = new Dictionary<Word.ElementType, string>
        {
            { Word.ElementType.Picture, "!PICTURE!" },
            { Word.ElementType.Chart, "!CHART!" },
            { Word.ElementType.Shape, "!SHAPE!" },
            { Word.ElementType.PreservedInline, "!PRESERVEDINLINE!" },
            { Word.ElementType.Table, "!TABLE!" }
        };

        // Ввод лицензионного ключа
        internal static void SetLicense()
        {
            Word.ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }
        
        // Открыть документ filePath
        internal static Word.DocumentModel? OpenDocument(string filePath)
        {
            try
            {
                Word.DocumentModel document = Word.DocumentModel.Load(filePath);
                document.CalculateListItems();
                // TODO: NOT SUPPORTED IN OUR DLL
                //document.GetPaginator(new Word.PaginatorOptions() { UpdateFields = true });
                return document;
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
#endif
                Console.WriteLine("Can't open document");
                return null;
            }
        }

        // Получить первые prefixLength символов параграфа paragraph (если длина меньше, то вернуть весь параграф)
        internal static string GetParagraphPrefix(Word.Paragraph paragraph, int prefixLength)
        {
            string result = paragraph.Content.ToString().Length > prefixLength ? paragraph.Content.ToString().Substring(0, prefixLength) : paragraph.Content.ToString();
            return result.Trim();
        }

        // Получить текстовое содержимое параграфа (без символа следующей строки)
        internal static string GetParagraphContentWithoutNewLine(Word.Paragraph paragraph)
        {
            string result = "";
            foreach (Word.Run runner in paragraph.GetChildElements(false,  Word.ElementType.Run))
            {
                result += runner.Content;
            }

            result = result.Trim();

            return result;
        }

        // Получить содержимое параграфа с учетом пропускаемых элементов
        internal static string GetParagraphContentWithSkippables(Word.Paragraph paragraph)
        {
            string result = "";
            foreach (var element in paragraph.GetChildElements(false, new Word.ElementType[] { Word.ElementType.Run, Word.ElementType.Picture, Word.ElementType.Chart, Word.ElementType.Shape, Word.ElementType.PreservedInline }))
            {
                switch (element)
                {
                    case Word.Run run:
                        result += run.Content;
                        break;
                    case Word.Picture picture:
                        result += SkippableElements[Word.ElementType.Picture];
                        break;
                    case Word.Chart _:
                        result += SkippableElements[Word.ElementType.Chart];
                        break;
                    case Word.Drawing.Shape _:
                        result += SkippableElements[Word.ElementType.Shape];
                        break;
                    case Word.PreservedInline _:
                        result += SkippableElements[Word.ElementType.PreservedInline];
                        break;
                    default:
                        Console.WriteLine("Unsupported element");
                        break;
                }
            }

            result = result.Trim();
            if (result == "") { result = "!SPACE!"; }
            return result;
        }

        // Получить список классифицированных параграфов для документа с помощью результатов классификации
        internal static List<ClassifiedParagraph> CombineParagraphsWithClassificationResult(Word.DocumentModel document, List<ClassificationResult> classificationResultList)
        {
            List<ClassifiedParagraph> classifiedParagraphs = new List<ClassifiedParagraph>();

            List<Word.Element> elements = new List<Word.Element>();
            foreach (Word.Section section in document.GetChildElements(recursively: false, filterElements: Word.ElementType.Section))
            {
                foreach (var element in section.GetChildElements(recursively: false, filterElements: new Word.ElementType[] { Word.ElementType.Paragraph, Word.ElementType.Table }))
                {
                    elements.Add(element);
                }
            }

            int classificationResultIndex = 0;
            int paragraphIndex = 0;
            foreach (Word.Element element in elements)
            {
                int classifiedParagraphIndex;
                try { classifiedParagraphIndex = classificationResultList[classificationResultIndex].Id; } catch { return classifiedParagraphs; }
                if (paragraphIndex < classifiedParagraphIndex)
                {
                    //ParagraphClass? paragraphClass = null;

                    //// Пропускаемые элементы
                    //// Таблицы
                    //if (element is Word.Tables.Table) { paragraphClass = ParagraphClass.e0; }
                   
                    //if (element is Word.Paragraph paragraph)
                    //{
                    //    // Списки
                    //    // TODO: Какой конкретный класс перечисления
                    //    if (paragraph.ListFormat.IsList) { paragraphClass = ParagraphClass.d0; }

                    //    string paragraphContentWithSkippables = GemBoxHelper.GetParagraphContentWithSkippables(paragraph);
                    //    // Картинки
                    //    // TODO: Какой конкретный класс картинки
                    //    if (paragraphContentWithSkippables == SkippableElements[Word.ElementType.Picture]) { paragraphClass = ParagraphClass.g0; }
                    //}

                    //classifiedParagraphs.Add(new ClassifiedParagraph(elements[paragraphIndex], paragraphClass));
                    paragraphIndex++;
                    continue;
                }

                classifiedParagraphs.Add(new ClassifiedParagraph(elements[paragraphIndex], classificationResultList[classificationResultIndex].ParagraphClass));

                classificationResultIndex++;
                paragraphIndex++;
            }

            return classifiedParagraphs;
        }
    }
}
