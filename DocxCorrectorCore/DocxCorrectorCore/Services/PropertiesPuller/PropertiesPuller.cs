using System.Collections.Generic;
using DocxCorrectorCore.Models;

namespace DocxCorrectorCore.Services.PropertiesPuller
{
    public abstract class PropertiesPuller
    {
        // Получить свойства всех параграфов документа filePath
        public abstract List<ParagraphProperties> GetAllParagraphsProperties(string filePath);

        // Получить свойства страниц документа filePath
        public abstract List<PageProperties> GetAllPagesProperties(string filePath);

        // Получить свойства секций документа filePath
        public abstract List<SectionProperties> GetAllSectionsProperties(string filePath);

        // Получить нормализованные свойства параграфов документа filePath
        public abstract List<NormalizedProperties> GetNormalizedProperties(string filePath);

        // Получить свойства верхних/нижних (type) колонтитулов документа filePath
        public abstract List<HeaderFooterInfo> GetHeadersFootersInfo(HeaderFooterType type, string filePath);
    }
}
