using System;
using System.Collections.Generic;
using DocxCorrector.Models;
using DocxCorrector.Services;
using Word = Microsoft.Office.Interop.Word;

namespace DocxCorrector.Services.Corrector
{
    class CorrectorInteropExeption : Exception
    {
        public CorrectorInteropExeption(string message) : base(message) { }
    }

    public sealed class CorrectorInterop : Corrector
    {
        // Private
        private Word.Application App;
        private Word.Document Document;

        // Приготовится к началу работы
        private void OpenApp()
        {
            try
            {
                if (App != null) { CloseApp(); }
                App = new Word.Application { Visible = false };
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
                CloseApp();
#endif
            }
        }

        // Приготовится к окончанию работы
        private void CloseApp()
        {
            try
            {
                if (App != null) { App.Quit(); }
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
#endif
            }
        }

        // Открыть документ
        private void OpenDocument()
        {
            try
            {
                Document = App.Documents.Open(FileName: FilePath, Visible: true);
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
#endif
                throw new CorrectorInteropExeption(message: "Can't open document");
            }
        }

        // Закрыть документ
        private void CloseDocument()
        {
            try
            {
                if (Document != null) { App.Documents.Close(); }
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
#endif
            }
        }
        
        // Закрыть документ без сохранения
        private void CloseDocumentWithoutSavingChanges()
        {
            try
            {
                if (Document != null) { Document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); }
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
#endif
            }
        }

        // Corrector
        public CorrectorInterop(string filePath = null) : base(filePath) { }

        // Получение JSON-а со списком ошибок
        public override List<ParagraphResult> GetMistakes()
        {
            try
            {
                OpenApp();
                OpenDocument();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                CloseApp();
                return new List<ParagraphResult>();
            }

            List<ParagraphResult> paragraphResults = new List<ParagraphResult>();

            // TODO: - Remove
            ParagraphResult testResult = new ParagraphResult
            {
                ParagraphID = 0,
                Type = ElementType.Paragraph,
                Prefix = "TestParagraph",
                Mistakes = new List<Mistake> { new Mistake(message: "Русские буквы") }
            };
            paragraphResults.Add(testResult);

            // TODO: - Implement method

            CloseApp();

            return paragraphResults;
        }

        // Получить свойства всех параграфов
        public override List<ParagraphProperties> GetAllParagraphsProperties()
        {
            try
            {
                OpenApp();
                OpenDocument();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                CloseApp();
                return new List<ParagraphProperties>();
            }

            List<ParagraphProperties> allParagraphsProperties = new List<ParagraphProperties>();

            foreach (Word.Paragraph paragraph in Document.Paragraphs)
            {
                ParagraphProperties paragraphProperties = new ParagraphProperties(paragraph);
                allParagraphsProperties.Add(paragraphProperties);
            }

            CloseApp();

            return allParagraphsProperties;
        }
        
        //Получить свойства всех страниц
        public override List<PageProperties> GetAllPagesProperties()
        {
            List<PageProperties> result = new List<PageProperties>();
            try
            {
                OpenApp();
                OpenDocument();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                CloseApp();
            }
            
            int totalPageNumber = Document.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
            for (int pageNumber = 1; pageNumber <= totalPageNumber; pageNumber++)
            {
                Word.Range pageRange = Document.Range().GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToAbsolute, pageNumber);
                PageProperties currentPageProperties = new PageProperties(pageSetup: pageRange.PageSetup, pageNumber: pageNumber);
                result.Add(currentPageProperties);
            }

            CloseDocumentWithoutSavingChanges();
            CloseApp();
            return result;
        }

        // Получить нормализованные свойства параграфов (Для классификатора Ромы)
        public override List<NormalizedProperties> GetNormalizedProperties()
        {
            try
            {
                OpenApp();
                OpenDocument();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                CloseApp();
                return new List<NormalizedProperties>();
            }

            List<NormalizedProperties> allNormalizedProperties = new List<NormalizedProperties>();

            int iteration = 0;
            foreach (Word.Paragraph paragraph in Document.Paragraphs)
            {
                NormalizedProperties normalizedParagraphProperties = new NormalizedProperties(paragraph: paragraph, paragraphId: iteration);
                allNormalizedProperties.Add(normalizedParagraphProperties);
                iteration++;
            }

            CloseApp();

            return allNormalizedProperties;
        }

        // Печать всех абзацев
        public override void PrintAllParagraphs()
        {
            try
            {
                OpenApp();
                OpenDocument();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                CloseApp();
                return;
            }

            foreach (Word.Paragraph paragraph in Document.Paragraphs)
            {
                Console.WriteLine(paragraph.Range.Text);
            }

            CloseApp();
        }
    }
}