
namespace DocxCorrector.App
{
    public static class Config
    {
        // Docx файл для проверки
        //public const string DocFilePath = @"C:\Users\Дмитрий\Desktop\paragraph_check\paragraph_check.docx";
        public const string DocFilePath = @"C:\Users\haide\Desktop\testDoc.docx";
        // Файл для записи ошибок
        public const string MistakesFilePath = @"C:\Users\haide\Desktop\mistakes.json";
        // Файл для записи свойств параграфов
        public const string PropertiesFilePath = @"C:\Users\haide\Desktop\properties.csv";
        // Корневая директория с файлами, из которых нужно вытянуть свойства
        public const string FilesToInpectDirectoryPath = @"C:\Users\Дмитрий\Desktop\paragraph_check\FilesToInspect";
    }
}
