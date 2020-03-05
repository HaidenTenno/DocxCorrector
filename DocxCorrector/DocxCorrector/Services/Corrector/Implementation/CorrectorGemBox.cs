#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using DocxCorrector.Models;
using Word = GemBox.Document;

namespace DocxCorrector.Services.Corrector
{
    class CorrectorGemBoxExeption : Exception
    {
        public CorrectorGemBoxExeption(string message) : base(message) { }
    }

    public sealed class CorrectorGemBox : Corrector
    {
        // Private
        private Word.DocumentModel? Document;
        
        // Ввод лицензионного ключа
        private void SetLicense()
        {
            Word.ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }

        // Открыть документ
        private void OpenDocument()
        {
            try
            {
                Document = Word.DocumentModel.Load(FilePath);
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
#endif
                throw new CorrectorGemBoxExeption(message: "Can't open document");
            }
        }

        // Corrector
        public CorrectorGemBox(string? filePath = null) : base(filePath)
        {
            SetLicense();
        }

        // Получение JSON-а со списком ошибок
        public override List<ParagraphResult> GetMistakes()
        {
            throw new NotImplementedException();
        }

        // Получить свойства всех параграфов
        public override List<ParagraphProperties> GetAllParagraphsProperties()
        {
            throw new NotImplementedException();
        }

        //Получить свойства всех страниц
        public override List<PageProperties> GetAllPagesProperties()
        {
            throw new NotImplementedException();
        }

        // Получить нормализованные свойства параграфов (Для классификатора Ромы)
        public override List<NormalizedProperties> GetNormalizedProperties()
        {
            try
            {
                OpenDocument();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return new List<NormalizedProperties>();
            }

            List<NormalizedProperties> allNormalizedProperties = new List<NormalizedProperties>();

            int iteration = 0;
            foreach (Word.Paragraph paragraph in Document!.GetChildElements(recursively: true, filterElements: Word.ElementType.Paragraph))
            {
                NormalizedProperties normalizedParagraphProperties = new NormalizedPropertiesGemBox(paragraph: paragraph, paragraphId: iteration);
                allNormalizedProperties.Add(normalizedParagraphProperties);
                iteration++;
            }

            return allNormalizedProperties;
        }

        // Печать всех абзацев
        public override void PrintAllParagraphs()
        {
            try
            {
                OpenDocument();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return;
            }

            foreach (Word.Paragraph paragraph in Document!.GetChildElements(recursively: true, filterElements: Word.ElementType.Paragraph))
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
        public override List<ParagraphResult> GetMistakesForElementType(ElementType elementType)
        {
            throw new NotImplementedException();
        }
    }
}
