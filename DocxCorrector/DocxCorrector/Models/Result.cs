using System.Collections.Generic;

namespace DocxCorrector.Models
{
    // Тип параграфа
    public enum ElementType
    {
        Paragraph, // Абзац
        Headline, // Заголовок
        List, // Список
        SourcesList, // Список источников
        Image, // Рисунок
        ImageSign // Подпись к рисунку
    }

    // Результат проверки для параграфа
    public sealed class ParagraphResult
    {
        // ID параграфа
        public int ParagraphID { get; set; }
        // Тип параграфа
        public ElementType Type { get; set; }
        // Начало параграфа (20 символов)
        public string Suffix { get; set; }
        // Ошибки в параграфе
        public List<Mistake> Mistakes { get; set; }
    }
    public sealed class Mistake
    {
        // Сообщение об ошибке
        public string Message { get; set; }
    }
}