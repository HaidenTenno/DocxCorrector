#nullable enable
using System;
using System.Collections.Generic;
using DocxCorrector.Models;

namespace DocxCorrector.Services.Corrector
{
    public abstract class Corrector
    {
        // Путь к docx файлу
        public virtual string? FilePath { get; set; }

        public Corrector(string? filePath = null)
        {
            FilePath = filePath;
        }

        // Получение списка ошибок
        public abstract List<ParagraphResult> GetMistakes();

        // Получить свойства всех параграфов
        public abstract List<ParagraphProperties> GetAllParagraphsProperties();

        // Получить свойства страниц документа
        public abstract List<PageProperties> GetAllPagesProperties();

        // Получить нормализованные свойства параграфов (Для классификатора Ромы)
        public abstract List<NormalizedProperties> GetNormalizedProperties();

        // Вспомогательные на момент разработки методы, которые, возможно, подлежат удалению
        // Печать всех абзацев
        public abstract void PrintAllParagraphs();

        // Получить спискок ошибок для выбранного документа, с учетом того, что все параграфы в нем типа elementType
        public abstract List<ParagraphResult> GetMistakesForElementType(ElementType elementType);
    }
}
