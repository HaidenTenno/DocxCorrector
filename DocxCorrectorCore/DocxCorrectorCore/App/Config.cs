
namespace DocxCorrector.App
{
    public static class Config
    {
        // Docx файл для проверки
        public const string DocFilePath = @"C:\Users\haide\Desktop\testDoc.docx";
        // Файл для записи ошибок
        public const string MistakesFilePath = @"C:\Users\haide\Desktop\mistakes.json";
        // Файл для записи свойств страниц
        public const string PagesPropertiesFilePath = @"C:\Users\haide\Desktop\pagesProperties.json";
        // Корневая директория с файлами, из которых нужно вытянуть свойства
        public const string FilesToInpectDirectoryPath = @"C:\Users\haide\Desktop\FilesToInspect";
        // Называние для файла со свойствами параграфа
        public const string ParagraphPropertiesFileName = @"\properties.csv";
        // Название для файла с нормализованными свойствами параграфа
        public const string NormalizedPropertiesFileName = @"\normalizedProperties.csv";
        // Называния csv файлов для тестирования синхронных/асинхронных методов
        public const string SyncParagraphsSyncIteration = @"\syncParagraphsSyncIteration.csv";
        public const string SyncParagraphsAsyncIteration = @"\syncParagraphAsyncIteration.csv";
        public const string AsyncParagraphsSyncIteration = @"\asyncParagraphsSyncIteration.csv";
        public const string AsyncParagraphsAsyncIteration = @"\asyncParagraphsAsyncIteration.csv";
    }
}
