using System;
using System.Collections.Generic;
using DocxCorrector.Models;

namespace DocxCorrector.Services.Corrector
{
    public abstract class Corrector : IDisposable
    {
        // Получить свойства всех параграфов документа filePath
        public abstract List<ParagraphProperties> GetAllParagraphsProperties(string filePath);

        // Получить свойства страниц документа filePath
        public abstract List<PageProperties> GetAllPagesProperties(string filePath);

        // Получить нормализованные свойства параграфов документа filePath (Для классификатора Ромы)
        public abstract List<NormalizedProperties> GetNormalizedProperties(string filePath);

        // Вспомогательные на момент разработки методы, которые, возможно, подлежат удалению
        // Печать всех абзацев документа filePath
        public abstract void PrintAllParagraphs(string filePath);

        // Получить спискок ошибок для документа filePath, с учетом того, что все параграфы в нем типа elementType
        public abstract List<ParagraphResult> GetMistakesForElementType(string filePath, ElementType elementType);

        // IDisposable
        public abstract void Dispose();
    }
}
