using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.PropertiesPuller
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
        public override List<ParagraphPropertiesGemBox> GetAllParagraphsProperties(string filePath)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return new List<ParagraphPropertiesGemBox>(); }

            List<ParagraphPropertiesGemBox> allParagraphProperties = new List<ParagraphPropertiesGemBox>();

            int paragraphID = 0;

            foreach (Word.Section section in document.GetChildElements(recursively: false, filterElements: Word.ElementType.Section))
            {
                foreach (var element in section.GetChildElements(recursively: false, filterElements: new Word.ElementType[] { Word.ElementType.Paragraph, Word.ElementType.Table }))
                {
                    ParagraphPropertiesGemBox paragraphProperties;

                    // Пропуск НЕ параграфов
                    if (!(element is Word.Paragraph paragraph)) { paragraphID++; continue; }
                    // Пропуск списков
                    if (paragraph.ListFormat.IsList) { paragraphID++; continue; }

                    string paragraphContentWithSkippables = GemBoxHelper.GetParagraphContentWithSkippables(paragraph);
                    // Пропуск картинок
                    if (paragraphContentWithSkippables == GemBoxHelper.SkippableElements[Word.ElementType.Picture]) { paragraphID++; continue; }
                    // Пропуск графиков
                    if (paragraphContentWithSkippables == GemBoxHelper.SkippableElements[Word.ElementType.Chart]) { paragraphID++; continue; }
                    // Пропуск фигур
                    if (paragraphContentWithSkippables == GemBoxHelper.SkippableElements[Word.ElementType.Shape]) { paragraphID++; continue; }
                    // Пропуск старых элементов doc (preserved inline)
                    if (paragraphContentWithSkippables == GemBoxHelper.SkippableElements[Word.ElementType.PreservedInline]) { paragraphID++; continue; }
                    // Пропуск SPACEов
                    if (paragraphContentWithSkippables == "!SPACE!") { paragraphID++; continue; }

                    paragraphProperties = new ParagraphPropertiesGemBox(paragraphID, paragraph);
                    allParagraphProperties.Add(paragraphProperties);

                    paragraphID++;
                }
            }

            return allParagraphProperties;
        }

        // Получить свойства всех параграфов документа filePath (для таблицы 0)
        public override List<ParagraphPropertiesTableZero> GetAllParagraphsPropertiesForTableZero(string filePath)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return new List<ParagraphPropertiesTableZero>(); }

            List<ParagraphPropertiesTableZero> allParagraphProperties = new List<ParagraphPropertiesTableZero>();

            int paragraphID = 0;

            foreach (Word.Section section in document.GetChildElements(recursively: false, filterElements: Word.ElementType.Section))
            {
                foreach (var element in section.GetChildElements(recursively: false, filterElements: new Word.ElementType[] { Word.ElementType.Paragraph, Word.ElementType.Table }))
                {
                    ParagraphPropertiesTableZero paragraphProperties;

                    if (element is Word.Tables.Table) { paragraphProperties = new ParagraphPropertiesTableZero(paragraphID, GemBoxHelper.SkippableElements[Word.ElementType.Table]); }
                    else
                    {
                        if (!(element is Word.Paragraph paragraph)) { paragraphID++; continue; }
                        paragraphProperties = new ParagraphPropertiesTableZero(paragraphID, paragraph);
                    }
                    allParagraphProperties.Add(paragraphProperties);

                    paragraphID++;
                }
            }

            return allParagraphProperties;
        }

        // Получить свойства страниц документа filePath
        public override List<PagePropertiesGemBox> GetAllPagesProperties(string filePath)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return new List<PagePropertiesGemBox>(); }

            List<PagePropertiesGemBox> pageProperties = new List<PagePropertiesGemBox>();

            var pages = document.GetPaginator().Pages;

            int pageNumber = 1;
            foreach (var page in pages)
            {
                PagePropertiesGemBox currentPageProperties = new PagePropertiesGemBox(page: page, pageNumber: pageNumber);
                pageProperties.Add(currentPageProperties);
                pageNumber++;
            }

            return pageProperties;
        }

        // Получить свойства секций документа filePath
        public override List<SectionPropertiesGemBox> GetAllSectionsProperties(string filePath)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return new List<SectionPropertiesGemBox>(); }

            List<SectionPropertiesGemBox> allSectionsProperties = new List<SectionPropertiesGemBox>();

            int sectionNumber = 1;
            foreach (Word.Section section in document.GetChildElements(recursively: true, filterElements: Word.ElementType.Section))
            {
                SectionPropertiesGemBox currentSectionProperties = new SectionPropertiesGemBox(section: section, sectionNumber: sectionNumber);
                allSectionsProperties.Add(currentSectionProperties);
                sectionNumber++;
            }

            return allSectionsProperties;
        }

        // Получить свойства верхних/нижних (type) колонтитулов документа filePath
        public override List<HeaderFooterInfoGemBox> GetHeadersFootersInfo(HeaderFooterType type, string filePath)
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
            if (document == null) { return new List<HeaderFooterInfoGemBox>(); }

            List<HeaderFooterInfoGemBox> headersFootersInfo = new List<HeaderFooterInfoGemBox>();

            List<Word.HeaderFooterType> chosenTypes = GetChosenHeaderFooterType(type);

            foreach (Word.Section section in document.GetChildElements(true, Word.ElementType.Section))
            {
                foreach (Word.HeaderFooter headerFooter in section.HeadersFooters)
                {
                    if (chosenTypes.Contains(headerFooter.HeaderFooterType))
                    {
                        HeaderFooterInfoGemBox headerFooterInfo = new HeaderFooterInfoGemBox(headerFooter);
                        headersFootersInfo.Add(headerFooterInfo);
                    }
                }
            }
            return headersFootersInfo;
        }

        // IPropertiesPullerAsync
        // Private
        private Task<ParagraphPropertiesGemBox> GetParagraphPropertiesAsync(int id, Word.Paragraph paragraph)
        {
            return Task.Run(() => new ParagraphPropertiesGemBox(id, paragraph));
        }

        // Public
        // Для уверенности, что интерфейс реализуют только наследники Correcor
        public PropertiesPuller PropertiesPuller => this;

        // Асинхронно получить свойства всех параграфов
        public async Task<List<ParagraphPropertiesGemBox>> GetAllParagraphsPropertiesAsync(string filePath)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return new List<ParagraphPropertiesGemBox>(); }

            List<Task<ParagraphPropertiesGemBox>> listOfTasks = new List<Task<ParagraphPropertiesGemBox>>();

            int paragraphID = 0;

            foreach (Word.Section section in document.GetChildElements(recursively: false, filterElements: Word.ElementType.Section))
            {
                foreach (var element in section.GetChildElements(recursively: false, filterElements: new Word.ElementType[] { Word.ElementType.Paragraph, Word.ElementType.Table }))
                {
                    Task<ParagraphPropertiesGemBox> paragraphPropertiesTask;

                    // Пропуск НЕ параграфов
                    if (!(element is Word.Paragraph paragraph)) { paragraphID++; continue; }
                    // Пропуск списков
                    if (paragraph.ListFormat.IsList) { paragraphID++; continue; }

                    string paragraphContentWithSkippables = GemBoxHelper.GetParagraphContentWithSkippables(paragraph);
                    // Пропуск картинок
                    if (paragraphContentWithSkippables == GemBoxHelper.SkippableElements[Word.ElementType.Picture]) { paragraphID++; continue; }
                    // Пропуск графиков
                    if (paragraphContentWithSkippables == GemBoxHelper.SkippableElements[Word.ElementType.Chart]) { paragraphID++; continue; }
                    // Пропуск фигур
                    if (paragraphContentWithSkippables == GemBoxHelper.SkippableElements[Word.ElementType.Shape]) { paragraphID++; continue; }
                    // Пропуск старых элементов doc (preserved inline)
                    if (paragraphContentWithSkippables == GemBoxHelper.SkippableElements[Word.ElementType.PreservedInline]) { paragraphID++; continue; }
                    // Пропуск SPACEов
                    if (paragraphContentWithSkippables == "!SPACE!") { paragraphID++; continue; }

                    paragraphPropertiesTask = GetParagraphPropertiesAsync(paragraphID, paragraph);
                    listOfTasks.Add(paragraphPropertiesTask);

                    paragraphID++;
                }
            }

            var result = await Task.WhenAll(listOfTasks);
            return result.ToList();
        }
    }
}
