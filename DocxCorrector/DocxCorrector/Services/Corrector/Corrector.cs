using System;
using System.Collections.Generic;

namespace DocxCorrector.Services.Corrector
{
    public abstract class Corrector
    {
        // Путь к docx файлу
        public string FilePath { get; set; }

        public Corrector(string filePath)
        {
            FilePath = filePath;
        }
                
        // Получение JSON-а со списком ошибок
        public abstract string GetMistakesJSON();

        // Получить свойства всех параграфов
        public abstract List<Models.ParagraphProperties> GetAllParagraphsProperties();

        // MARK: - Вспомогательные на момент разработки методы, котоые возможно подлежат удалению
        // Печать всех абзацев
        public abstract void PrintAllParagraphs();

        // Напечатать свойства первого параграфа
        public abstract void PrintFirstParagraphProperties();
        // Напечатать свойства первых двух слов первого параграфа
        public abstract void PrintFirstTwoWordsProperties();
    }
}
