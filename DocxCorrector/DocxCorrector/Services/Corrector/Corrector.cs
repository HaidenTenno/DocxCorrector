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

        // Печать всех абзацев
        public abstract void PrintAllParagraphs();

        // Получение JSON-а со списком ошибок
        public abstract string GetMistakesJSON();
    }
}
