#nullable enable
using System;
using System.Collections.Generic;
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
            throw new NotImplementedException();
        }

        // Печать всех абзацев
        public override void PrintAllParagraphs()
        {
            if (Document == null)
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
            }

            foreach (Word.Paragraph paragraph in Document!.GetChildElements(recursively: true, filterElements: Word.ElementType.Paragraph))
            {
                string text = paragraph.Content.ToString();
                Console.WriteLine(text);
            }
        }

        // Получить списк ошибок для выбранного документа, с учетом того, что все параграфы в нем типа elementType
        public override List<ParagraphResult> GetMistakesForElementType(ElementType elementType)
        {
            throw new NotImplementedException();
        }
    }
}
