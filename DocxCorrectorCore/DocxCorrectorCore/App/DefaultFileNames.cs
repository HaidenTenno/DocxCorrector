namespace DocxCorrectorCore.App
{
    public static class DefaultFileNames
    {
        // Файл для записи ошибок
        public const string MistakesFileName = "mistakes.json";
        // Файл для записи свойств страниц
        public const string PagesPropertiesFileName = "pagesProperties.json";
        // Файл для записи свойств секций
        public const string SectionsPropertiesFileName = "sectionsProperties.json";
        // Файл для записи свойств колонтитулов
        public const string HeadersFootersInfoFileName = "headersFootersInfo.json";
        // Файл со свойствами параграфов
        public const string ParagraphsPropertiesFileName = "properties.csv";
        // Файл со свойствами параграфов (ДЛЯ ТАБЛИЦЫ 0)
        public const string ParagraphsPropertiesForTableZeroFileName = "propertiesTable0.csv";
        // Называния csv файлов для тестирования синхронных/асинхронных методов
        public const string SyncParagraphsSyncIterationFileName = "syncParagraphsSyncIteration.csv";
        public const string SyncParagraphsAsyncIterationFileName = "syncParagraphAsyncIteration.csv";
        public const string AsyncParagraphsSyncIterationFileName = "asyncParagraphsSyncIteration.csv";
        public const string AsyncParagraphsAsyncIterationFileName = "asyncParagraphsAsyncIteration.csv";
    }
}
