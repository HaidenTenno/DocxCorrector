﻿
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
        // Файл для записи свойств параграфов
        public const string PropertiesFilePath = @"C:\Users\haide\Desktop\properties.csv";
        // Корневая директория с файлами, из которых нужно вытянуть свойства
        public const string FilesToInpectDirectoryPath = @"C:\Users\haide\Desktop\FilesToInspect";
    }
}
