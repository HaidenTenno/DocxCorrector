#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DocxCorrectorCore.Models;
using Word = GemBox.Document;

namespace DocxCorrectorCore.Services.Corrector
{
    public sealed class CorrectorGemBox : Corrector, ICorrecorAsync
    {
        // Private
        // Ввод лицензионного ключа
        private void SetLicense()
        {
            Word.ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }

        // Открыть документ filePath
        private Word.DocumentModel? OpenDocument(string filePath)
        {
            try
            {
                Word.DocumentModel document = Word.DocumentModel.Load(filePath);
                document.CalculateListItems();
                document.GetPaginator(new Word.PaginatorOptions() { UpdateFields = true });
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

        // Public
        // IDisposable
        public override void Dispose() { }

        // Corrector
        public CorrectorGemBox()
        {
            SetLicense();
        }

        // Получить свойства всех параграфов документа filePath
        public override List<ParagraphProperties> GetAllParagraphsProperties(string filePath)
        {
            Word.DocumentModel? document = OpenDocument(filePath: filePath);
            if (document == null) { return new List<ParagraphProperties>(); }

            List<ParagraphProperties> allParagraphProperties = new List<ParagraphProperties>();

            foreach (Word.Paragraph paragraph in document.GetChildElements(recursively: true, filterElements: Word.ElementType.Paragraph))
            {
                ParagraphProperties paragraphProperties = new ParagraphPropertiesGemBox(paragraph: paragraph);
                allParagraphProperties.Add(paragraphProperties);
            }

            return allParagraphProperties;
        }

        // Получить свойства секций документа filePath
        public override List<SectionProperties> GetAllSectionsProperties(string filePath)
        {
            Word.DocumentModel? document = OpenDocument(filePath: filePath);
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

        //Получить свойства всех страниц документа filePath
        public override List<PageProperties> GetAllPagesProperties(string filePath)
        {
            Word.DocumentModel? document = OpenDocument(filePath: filePath);
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

        // Получить нормализованные свойства параграфов документа filePath (Для классификатора Ромы)
        public override List<NormalizedProperties> GetNormalizedProperties(string filePath)
        {
            Word.DocumentModel? document = OpenDocument(filePath: filePath);
            if (document == null) { return new List<NormalizedProperties>(); }

            List<NormalizedProperties> allNormalizedProperties = new List<NormalizedProperties>();

            int iteration = 0;
            foreach (Word.Paragraph paragraph in document.GetChildElements(recursively: true, filterElements: Word.ElementType.Paragraph))
            {
                NormalizedProperties normalizedParagraphProperties = new NormalizedPropertiesGemBox(paragraph: paragraph, paragraphId: iteration);
                allNormalizedProperties.Add(normalizedParagraphProperties);
                iteration++;
            }

            return allNormalizedProperties;
        }

        // Получить свойства верхних/нижних колонтитулов документа filePath
        public override List<HeaderFooterInfo> GetHeadersFootersInfo(HeaderFooterType type, string filePath)
        {
            List<Word.HeaderFooterType> GetChosenHeaderFooterType(HeaderFooterType type)
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

            Word.DocumentModel? document = OpenDocument(filePath: filePath);
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

        // Печать всех абзацев документа filePath
        public override void PrintAllParagraphs(string filePath)
        {
            Word.DocumentModel? document = OpenDocument(filePath: filePath);
            if (document == null) { return; }

            foreach (Word.Paragraph paragraph in document.GetChildElements(recursively: true, filterElements: Word.ElementType.Paragraph))
            {
                int elementWitDifferentStyleCount = paragraph.GetChildElements(true, Word.ElementType.Run).Count();
                Console.WriteLine($"В этом параграфе {elementWitDifferentStyleCount} элемент(ов) с разным оформлением");
                foreach (Word.Run run in paragraph.GetChildElements(recursively: true, filterElements: Word.ElementType.Run)) 
                {
                    string text = run.Text;
                    Console.WriteLine(text);
                }
                Console.WriteLine();
            }
        }

        // Получить списк ошибок для документа filePath, с учетом того, что все параграфы в нем типа elementType
        public override List<ParagraphResult> GetMistakesForElementType(string filePath, ElementType elementType)
        {
            throw new NotImplementedException();
        }

        // ICorrectorAsync
        // Private
        private Task<ParagraphProperties> GetParagraphPropertiesAsync(Word.Paragraph paragraph)
        {
            return Task.Run(() => (ParagraphProperties)new ParagraphPropertiesGemBox(paragraph));
        }

        private Task<NormalizedProperties> GetNormalizedPropertiesAsync(Word.Paragraph paragraph, int paragraphId)
        {
            return Task.Run(() => (NormalizedProperties)new NormalizedPropertiesGemBox(paragraph, paragraphId));
        }

        // Public
        public Corrector Corrector => throw new NotImplementedException();

        public async Task<List<ParagraphProperties>> GetAllParagraphsPropertiesAsync(string filePath)
        {
            Word.DocumentModel? document = OpenDocument(filePath: filePath);
            if (document == null) { return new List<ParagraphProperties>(); }

            List<Task<ParagraphProperties>> listOfTasks = new List<Task<ParagraphProperties>>();

            foreach (Word.Paragraph paragraph in document.GetChildElements(recursively: true, filterElements: Word.ElementType.Paragraph))
            {
                listOfTasks.Add(GetParagraphPropertiesAsync(paragraph));
            }

            var result = await Task.WhenAll(listOfTasks);
            return result.ToList();
        }

        public async Task<List<NormalizedProperties>> GetNormalizedPropertiesAsync(string filePath)
        {
            Word.DocumentModel? document = OpenDocument(filePath: filePath);
            if (document == null) { return new List<NormalizedProperties>(); }

            List<Task<NormalizedProperties>> listOfTasks = new List<Task<NormalizedProperties>>();

            int iteration = 0;
            foreach (Word.Paragraph paragraph in document.GetChildElements(recursively: true, filterElements: Word.ElementType.Paragraph))
            {
                listOfTasks.Add(GetNormalizedPropertiesAsync(paragraph, iteration));
                iteration++;
            }

            var result = await Task.WhenAll(listOfTasks);
            return result.ToList();
        }
    }
}
