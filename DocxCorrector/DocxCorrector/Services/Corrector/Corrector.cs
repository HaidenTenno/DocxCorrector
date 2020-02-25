using System;
using System.Collections.Generic;
using DocxCorrector.Models;

namespace DocxCorrector.Services.Corrector
{
    public abstract class Corrector
    {
        // Путь к docx файлу
        public virtual string FilePath { get; set; }

        public Corrector(string filePath = null)
        {
            FilePath = filePath;
        }
        
        // Получение JSON-а со свойствами абзацев
        public abstract string GetAllPagesPropertiesJSON(List<PageProperties> allPageProperties);

        // Получение JSON-а со списком ошибок
        public abstract string GetMistakesJSON();

        // Получить свойства всех параграфов
        public abstract List<Models.ParagraphProperties> GetAllParagraphsProperties();

        public abstract List<Models.PageProperties> GetAllPagesProperties();

        // Получить нормализованные свойства параграфов (Для классификатора Ромы)
        public abstract List<Models.NormalizedProperties> GetNormalizedProperties();

        // Вспомогательные на момент разработки методы, котоые возможно подлежат удалению
        // Печать всех абзацев
        public abstract void PrintAllParagraphs();
    }
}
