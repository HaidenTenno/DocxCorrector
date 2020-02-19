using System;

namespace DocxCorrector.Services.Corrector
{
    public abstract class Corrector
    {
        // Путь к docx файлу
        public string FilePath { get; }

        public Corrector(string filePath)
        {
            FilePath = filePath;
        }
                
        // Получение JSON-а со списком ошибок
        public abstract string GetMistakesJSON();

        // MARK: - Вспомогательные на момент разработки методы, котоые возможно подлежат удалению
        // Печать всех абзацев
        public abstract void PrintAllParagraphs();

        // Напечатать свойства первого параграфа
        public abstract void PrintFirstParagraphProperties();
    }
}
