using System.Collections.Generic;
using DocxCorrectorCore.Models;

namespace DocxCorrectorCore.Services.PropertiesPuller
{
    public abstract class PropertiesPuller
    {
        // Напечатать содержимое документа filePath
        public abstract void PrintContent(string filePath);

        // Вывести в консоль информацию о структуре документа
        public abstract void PrintDocumentStructureInfo(string filePath);

        // Получить свойства содержания документа filePath
        public abstract void PrintTableOfContenstsInfo(string filePath);

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
