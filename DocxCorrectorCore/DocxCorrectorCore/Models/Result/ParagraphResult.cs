using System.Collections.Generic;

namespace DocxCorrector.Models
{
    // Тип параграфа
    public enum ElementType
    {
        Paragraph, // Абзац
        List, // Список
        ImageSign, // Подпись к рисунку
        Headline, // Заголовок
        SourcesList, // Список источников
        Image // Рисунок
    }

    // Результат проверки для параграфа
    public sealed class ParagraphResult
    {
        // ID параграфа
        public int ParagraphID { get; set; }
        // Тип параграфа
        public ElementType Type { get; set; }
        // Начало параграфа (20 символов)
        public string Prefix { get; set; }
        // Ошибки в параграфе
        public List<Mistake> Mistakes { get; set; }
    }
    public sealed class Mistake
    {
        // Сообщение об ошибке
        public string Message { get; set; }

        public Mistake(string message)
        {
            Message = message;
        }
    }
}