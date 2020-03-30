#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocxCorrector.Models;
using Spire.Doc;
using Spire.Doc.Documents;

namespace DocxCorrector.Services.Corrector
{
    public sealed class CorrectorSpire : Corrector
    {
        private static Document? Document { get; set; }

        private static void OpenDocument(string filePath)
        {
            try
            {
                Document = new Document();
                Document.LoadFromFile(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Can't open document");
            }
        }

        // Получить свойства всех параграфов документа filePath
        public override List<ParagraphProperties> GetAllParagraphsProperties(string filePath)
        {
            var allParagraphProperties = new List<ParagraphProperties>();
            OpenDocument(filePath);
            if (Document != null)
            {
                foreach (Section section in Document.Sections)
                {
                    foreach (Paragraph paragraph in section.Paragraphs)
                    {
                        var paragraphProperties = new ParagraphPropertiesSpire(paragraph: paragraph);
                        allParagraphProperties.Add(paragraphProperties);
                    }
                }
                Document.Close();
            }
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

        // Вспомогательные на момент разработки методы, которые, возможно, подлежат удалению
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

        // IDisposable
        public override void Dispose()
        {
            throw new NotImplementedException();
        }
    }
}
