#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using DocxCorrector.Models;
using Word = GemBox.Document;

namespace DocxCorrector.Services.Corrector
{
    public sealed class CorrectorGemBox : Corrector
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
        // Corrector
        public CorrectorGemBox()
        {
            SetLicense();
        }

        // Получить свойства всех параграфов
        public override List<ParagraphProperties> GetAllParagraphsProperties(string filePath)
        {
            throw new NotImplementedException();
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
    }
}
