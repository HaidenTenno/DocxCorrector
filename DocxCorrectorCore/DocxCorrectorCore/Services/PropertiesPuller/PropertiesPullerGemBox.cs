using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DocxCorrectorCore.Models;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.Services.PropertiesPuller
{
    public sealed class PropertiesPullerGemBox : PropertiesPuller, IPropertiesPullerAsync
    {
        // Public
        public PropertiesPullerGemBox()
        {
            GemBoxHelper.SetLicense();
        }

        // PropertiesPuller
        // Напечатать содержимое документа filePath
        public override void PrintContent(string filePath)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return; }

            Console.WriteLine(document.Content.ToString());
        }

        // Вывести в консоль информацию о структуре документа
        public override void PrintDocumentStructureInfo(string filePath)
        {
            static void PrintChildsInfo(Word.Element element, int interation)
            {
                var childElements = element.GetChildElements(false);

                if (childElements.Count() == 0) { return; }

                var prefixStr = "";
                for (int i = 0; i < interation; i++)
                {
                    prefixStr += "\t";
                }

                foreach (var childElement in childElements)
                {
                    Console.WriteLine($"{prefixStr}{childElement} -> {childElement.ElementType}");
                    PrintChildsInfo(childElement, interation + 1);
                }
            }

            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return; }

            PrintChildsInfo(document, 0);
        }


        // Вывести в консоль информацию о содержании документа filePath
        public override void PrintTableOfContenstsInfo(string filePath)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return; }

            foreach (Word.TableOfEntries toe in document.GetChildElements(recursively: true, filterElements: Word.ElementType.TableOfEntries))
            {
                Console.WriteLine($"CONTENT: {toe.Content}");
                Console.WriteLine($"INSTRUCTION TEXT: {toe.InstructionText}");
                foreach (var entry in toe.Entries)
                {
                    Console.WriteLine($"ENTRY: {entry.Content}");
                }
                Console.WriteLine($"FIELD TYPE: {toe.FieldType}");
                Console.WriteLine($"IS DIRTY: {toe.IsDirty}");                
            }
        }

        // Получить свойства всех параграфов документа filePath
        public override List<ParagraphProperties> GetAllParagraphsProperties(string filePath)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return new List<ParagraphProperties>(); }

            List<ParagraphProperties> allParagraphProperties = new List<ParagraphProperties>();

            foreach (Word.Section section in document.GetChildElements(recursively: false, filterElements: Word.ElementType.Section))
            {
                foreach (var element in section.GetChildElements(recursively: false, filterElements: new Word.ElementType[] { Word.ElementType.Paragraph, Word.ElementType.Table }))
                {
                    ParagraphProperties paragraphProperties;
                    switch (element)
                    {
                        case Word.Paragraph paragraph:
                            paragraphProperties = new ParagraphPropertiesGemBox(paragraph: paragraph);
                            allParagraphProperties.Add(paragraphProperties);
                            break;
                        case Word.Tables.Table _:
                            paragraphProperties = new ParagraphPropertiesGemBox(placeHolder: "TABLE");
                            allParagraphProperties.Add(paragraphProperties);
                            break;
                        default:
                            Console.WriteLine("Unsupported element");
                            break;
                    }
                }
            }

            return allParagraphProperties;
        }

        // Получить свойства страниц документа filePath
        public override List<PageProperties> GetAllPagesProperties(string filePath)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return new List<PageProperties>(); }

            List<PageProperties> pageProperties = new List<PageProperties>();

            var pages = document.GetPaginator().Pages;

            int pageNumber = 1;
            foreach (var page in pages)
            {
                PageProperties currentPageProperties = new PagePropertiesGemBox(page: page, pageNumber: pageNumber);
                pageProperties.Add(currentPageProperties);
                pageNumber++;
            }

            return pageProperties;
        }

        // Получить свойства секций документа filePath
        public override List<SectionProperties> GetAllSectionsProperties(string filePath)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return new List<SectionProperties>(); }

            List<SectionProperties> allSectionsProperties = new List<SectionProperties>();

            int sectionNumber = 1;
            foreach (Word.Section section in document.GetChildElements(recursively: true, filterElements: Word.ElementType.Section))
            {
                SectionProperties currentSectionProperties = new SectionPropertiesGemBox(section: section, sectionNumber: sectionNumber);
                allSectionsProperties.Add(currentSectionProperties);
                sectionNumber++;
            }

            return allSectionsProperties;
        }

        // Получить свойства верхних/нижних (type) колонтитулов документа filePath
        public override List<HeaderFooterInfo> GetHeadersFootersInfo(HeaderFooterType type, string filePath)
        {
            static List<Word.HeaderFooterType> GetChosenHeaderFooterType(HeaderFooterType type)
            {
                return type switch
                {
                    HeaderFooterType.Header => new List<Word.HeaderFooterType>()
                    {
                        Word.HeaderFooterType.HeaderFirst,
                        Word.HeaderFooterType.HeaderEven,
                        Word.HeaderFooterType.HeaderDefault
                    },
                    HeaderFooterType.Footer => new List<Word.HeaderFooterType>()
                    {
                        Word.HeaderFooterType.FooterFirst,
                        Word.HeaderFooterType.FooterEven,
                        Word.HeaderFooterType.FooterDefault
                    },
                    _ => throw new NotSupportedException()
                };
            }

            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return new List<HeaderFooterInfo>(); }

            List<HeaderFooterInfo> headersFootersInfo = new List<HeaderFooterInfo>();

            List<Word.HeaderFooterType> chosenTypes = GetChosenHeaderFooterType(type);

            foreach (Word.Section section in document.GetChildElements(true, Word.ElementType.Section))
            {
                foreach (Word.HeaderFooter headerFooter in section.HeadersFooters)
                {
                    if (chosenTypes.Contains(headerFooter.HeaderFooterType))
                    {
                        HeaderFooterInfo headerFooterInfo = new HeaderFooterInfoGemBox(headerFooter);
                        headersFootersInfo.Add(headerFooterInfo);
                    }
                }
            }
            return headersFootersInfo;
        }

        // IPropertiesPullerAsync
        // Private
        private Task<ParagraphProperties> GetParagraphPropertiesAsync(Word.Paragraph paragraph)
        {
            return Task.Run(() => (ParagraphProperties)new ParagraphPropertiesGemBox(paragraph));
        }

        // Public
        // Для уверенности, что интерфейс реализуют только наследники Correcor
        public PropertiesPuller PropertiesPuller => this;

        // Асинхронно получить свойства всех параграфов
        public async Task<List<ParagraphProperties>> GetAllParagraphsPropertiesAsync(string filePath)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return new List<ParagraphProperties>(); }

            List<Task<ParagraphProperties>> listOfTasks = new List<Task<ParagraphProperties>>();

            foreach (Word.Paragraph paragraph in document.GetChildElements(recursively: true, filterElements: Word.ElementType.Paragraph))
            {
                listOfTasks.Add(GetParagraphPropertiesAsync(paragraph));
            }

            var result = await Task.WhenAll(listOfTasks);
            return result.ToList();
        }
    }
}
