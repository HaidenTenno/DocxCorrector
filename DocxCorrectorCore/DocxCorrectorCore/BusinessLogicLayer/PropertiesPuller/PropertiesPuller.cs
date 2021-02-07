using System.Collections.Generic;

namespace DocxCorrectorCore.BusinessLogicLayer.PropertiesPuller
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
        public abstract List<ParagraphPropertiesGemBox> GetAllParagraphsProperties(string filePath);

        // Получить свойства всех параграфов документа filePath (для таблицы 0)
        public abstract List<ParagraphPropertiesTableZero> GetAllParagraphsPropertiesForTableZero(string filePath);

        // Получить свойства страниц документа filePath
        public abstract List<PagePropertiesGemBox> GetAllPagesProperties(string filePath);

        // Получить свойства секций документа filePath
        public abstract List<SectionPropertiesGemBox> GetAllSectionsProperties(string filePath);

        // Получить свойства верхних/нижних (type) колонтитулов документа filePath
        public abstract List<HeaderFooterInfoGemBox> GetHeadersFootersInfo(HeaderFooterType type, string filePath);

        // MARK: НИРМА 2020-2021
        // Получить свойства всех параграфов документа filePath + проставить там возможные классы из файла с пресетами presetsPath
        public abstract List<ParagraphPropertiesWithPresets> GetAllParagraphPropertiesWithPresets(string filePath, CombinedPresetValues combinedPresetValues);
    }
}
