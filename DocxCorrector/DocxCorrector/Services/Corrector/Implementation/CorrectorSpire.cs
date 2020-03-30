#nullable enable
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using DocxCorrector.Models;
using Word = Spire.Doc;

namespace DocxCorrector.Services.Corrector
{
    public sealed class CorrectorSpire : Corrector, ICorrecorAsync
    {
        // Private
        // Открыть документ filePath
        private Word.Document? OpenDocument(string filePath)
        {
            try
            {
                Word.Document document = new Word.Document(filePath);
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
        // Получить свойства всех параграфов документа filePath
        public override List<ParagraphProperties> GetAllParagraphsProperties(string filePath)
        {
            Word.Document? document = OpenDocument(filePath);
            if (document == null) { return new List<ParagraphProperties>(); }

            List<ParagraphProperties> allParagraphProperties = new List<ParagraphProperties>();

            foreach (Word.Section section in document.Sections)
            {
                foreach (Word.Documents.Paragraph paragraph in section.Paragraphs)
                {
                    var paragraphProperties = new ParagraphPropertiesSpire(paragraph: paragraph);
                    allParagraphProperties.Add(paragraphProperties);
                }
            }

            document.Close();
            return allParagraphProperties;
        }

        // Получить свойства страниц документа filePath
        public override List<PageProperties> GetAllPagesProperties(string filePath)
        {
            throw new NotImplementedException();
        }

        // Получить нормализованные свойства параграфов документа filePath (Для классификатора Ромы)
        public override List<NormalizedProperties> GetNormalizedProperties(string filePath)
        {
            throw new NotImplementedException();
        }

        // Печать всех абзацев документа filePath
        public override void PrintAllParagraphs(string filePath)
        {
            throw new NotImplementedException();
        }

        // Получить спискок ошибок для документа filePath, с учетом того, что все параграфы в нем типа elementType
        public override List<ParagraphResult> GetMistakesForElementType(string filePath, ElementType elementType)
        {
            throw new NotImplementedException();
        }

        // ICorrectorAsync
        // Private
        private Task<ParagraphProperties> GetParagraphPropertiesAsync(Word.Documents.Paragraph paragraph)
        {
            return Task.Run(() => (ParagraphProperties)new ParagraphPropertiesSpire(paragraph));
        }

        //private Task<NormalizedProperties> GetNormalizedPropertiesAsync(Word.Documents.Paragraph paragraph, int paragraphId)
        //{
        //    return Task.Run(() => (NormalizedProperties)new NormalizedPropertiesSpire(paragraph, paragraphId));
        //}

        // Public
        public Corrector Corrector => throw new NotImplementedException();

        public async Task<List<ParagraphProperties>> GetAllParagraphsPropertiesAsync(string filePath)
        {
            Word.Document? document = OpenDocument(filePath);
            if (document == null) { return new List<ParagraphProperties>(); }

            List<Task<ParagraphProperties>> listOfTasks = new List<Task<ParagraphProperties>>();

            foreach (Word.Section section in document.Sections)
            {
                foreach (Word.Documents.Paragraph paragraph in section.Paragraphs)
                {
                    listOfTasks.Add(GetParagraphPropertiesAsync(paragraph));
                }
            }

            var result = await Task.WhenAll(listOfTasks);

            document.Close();
            return result.ToList();
        }

        public async Task<List<NormalizedProperties>> GetNormalizedPropertiesAsync(string filePath)
        {
            throw new NotImplementedException();
        }
    }
}
