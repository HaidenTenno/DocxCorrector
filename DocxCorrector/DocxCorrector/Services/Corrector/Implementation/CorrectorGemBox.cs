#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DocxCorrector.Models;
using Word = GemBox.Document;

namespace DocxCorrector.Services.Corrector
{
    public sealed class CorrectorGemBox : Corrector, ICorrecorAsync
    {
        // Private
        // Ввод лицензионного ключа
        private void SetLicense()
        {
            Word.ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }

        // Открыть документ
        private Word.DocumentModel? OpenDocument(string filePath)
        {
            try
            {
                Word.DocumentModel document = Word.DocumentModel.Load(filePath);
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

        // Получить свойства всех параграфов
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

        //Получить свойства всех страниц
        public override List<PageProperties> GetAllPagesProperties(string filePath)
        {
            throw new NotImplementedException();
        }

        // Получить нормализованные свойства параграфов (Для классификатора Ромы)
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

        // Печать всех абзацев
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

        // Получить списк ошибок для выбранного документа, с учетом того, что все параграфы в нем типа elementType
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
