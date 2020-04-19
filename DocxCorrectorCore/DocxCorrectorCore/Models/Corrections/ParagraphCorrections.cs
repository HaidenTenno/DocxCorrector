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
        public readonly int ParagraphID;
        // Тип параграфа
        public readonly ParagraphClass ParagraphClass;
        // Начало параграфа (20 символов)
        public readonly string Prefix;
        // Ошибки в параграфе
        public readonly List<ParagraphMistake> Mistakes;

        public ParagraphCorrections(int paragraphID, ParagraphClass paragraphClass, string prefix, List<ParagraphMistake> mistakes)
        {
            ParagraphID = paragraphID;
            ParagraphClass = paragraphClass;
            Prefix = prefix;
            Mistakes = mistakes;
        }
    }
    public sealed class ParagraphMistake
    {
        // Сообщение об ошибке
        public readonly string Message;
        // Совет по исправлению
        public readonly string Advice;

        public ParagraphMistake(string message, string advice)
        {
            Message = message;
            Advice = advice;
        }
    }
}