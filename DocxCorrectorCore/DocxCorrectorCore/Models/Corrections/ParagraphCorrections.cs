using System.Collections.Generic;

namespace DocxCorrectorCore.Models
{
    // Тип параграфа
    public enum ParagraphClass
    {
        Paragraph, // Абзац
        List, // Элемент списка
        ImageSign, // Подпись к рисунку
        Headline, // Заголовок
        SourcesList, // Элемент списка источников
        Image // Рисунок
    }

    // Результат проверки для параграфа
    public sealed class ParagraphCorrections
    {
        // ID параграфа
        public int ParagraphID { get; set; }
        // Тип параграфа
        public ParagraphClass ParagraphClass { get; set; }
        // Начало параграфа (20 символов)
        public string Prefix { get; set; }
        // Ошибки в параграфе
        public List<ParagraphMistake> Mistakes { get; set; }
    }
    public sealed class ParagraphMistake
    {
        // Сообщение об ошибке
        public string Message { get; set; }

        public ParagraphMistake(string message)
        {
            Message = message;
        }
    }
}