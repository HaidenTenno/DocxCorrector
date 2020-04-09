using System;
using System.Collections.Generic;
using DocxCorrectorCore.Models;

namespace DocxCorrectorCore.Services.Corrector
{
    public abstract class Corrector : IDisposable
    {
        // Получить свойства всех параграфов документа filePath
        public abstract List<ParagraphProperties> GetAllParagraphsProperties(string filePath);

        // Получить свойства страниц документа filePath
        public abstract List<PageProperties> GetAllPagesProperties(string filePath);

        // Получить свойства секций документа filePath
        public abstract List<SectionProperties> GetAllSectionsProperties(string filePath);

        // Получить нормализованные свойства параграфов документа filePath (Для классификатора Ромы)
        public abstract List<NormalizedProperties> GetNormalizedProperties(string filePath);

        // Получить свойства верхних/нижних колонтитулов документа filePath
        public abstract List<HeaderFooterInfo> GetHeadersFootersInfo(HeaderFooterType type, string filePath);

        // Вспомогательные на момент разработки методы, которые, возможно, подлежат удалению
        // Печать всех абзацев документа filePath
        public abstract void PrintAllParagraphs(string filePath);

        // Получить спискок ошибок для документа filePath, с учетом того, что все параграфы в нем типа elementType
        public abstract List<ParagraphResult> GetMistakesForElementType(string filePath, ElementType elementType);

        // Сохранить документ filePath как pdf resultPath
        public abstract void SaveDocumentAsPdf(string filePath, string resultFilePath);

        // Сохранить страницы документа filePath как отдельные pdf в директории resultPath
        public abstract void SavePagesAsPdf(string filePath, string resultFilePath);

        // IDisposable
        public abstract void Dispose();
    }
}
