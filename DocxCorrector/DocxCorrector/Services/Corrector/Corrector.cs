using System;
using System.Collections.Generic;

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

        // Получение JSON-а со списком ошибок
        public abstract string GetMistakesJSON();

        // Получить свойства всех параграфов
        public abstract List<Models.ParagraphProperties> GetAllParagraphsProperties();

        // Получить нормализованные свойства параграфов (Для классификатора Ромы)
        public abstract List<Models.NormalizedProperties> GetNormalizedProperties();

        // Вспомогательные на момент разработки методы, котоые возможно подлежат удалению
        // Печать всех абзацев
        public abstract void PrintAllParagraphs();

        // Получить JSON со списком ошибок для выбранного документа, с учетом того, что все параграфы в нем типа elementType
        public abstract string GetMistakesJSONForElementType(Models.ElementType elementType);
    }
}
