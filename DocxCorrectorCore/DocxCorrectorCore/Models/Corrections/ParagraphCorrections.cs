using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace DocxCorrectorCore.Models.Corrections
{
    // Тип параграфа
    [JsonConverter(typeof(StringEnumConverter))]
    public enum ParagraphClass
    {
        // НЕТ A (содержание)
        b0, // ЗАГОЛОВОК
        b1, // Заголовок 1го уровня
        b2, // Заголовок 2го уровня
        b3, // Заголовок 3го уровня
        b4, // Заголовок 4го уровня
        c0, // АБЗАЦ
        c1, // Обычный абзац
        c2, // Абзац перед списком
        c3, // Абзац перед формулой
        d0, // ЭЛЕМЕНТ ПЕРЕЧИСЛЕНИЯ
        d1, // Первый элемент простого перечисления
        d2, // Элемент простого перечисления
        d3, // Последний элемент простого перечисления
        d4, // Первый элемент сложного перечисления
        d5, // Элемент сложного перчисления, находящийся между первым и послебним элементами списка
        d6, // Последний элемент сложного перечисления
        // НЕТ E (таблица)
        f0, // ПОДПИСЬ ТАБЛИЦЫ
        f1, // Подпись и нумерация таблиц в тексте
        f2, // Подпись и нумерация таблиц в приложении
        f3, // Подпись и нумерация продолжения таблицы в тексте
        f4, // Подпись и нумерация продолжения таблицы в приложении
        f5, // Подпись и нумерация окончания таблицы в тексте
        f6, // Подпись и нумерация окончания таблицы в приложении
        g0, // РИСУНОК
        g1, // Рисунок как отдельный параграф
        g2, // Рисунок плавающий
        g3, // Рисунок внутри текста
        h0, // ПОДПИСЬ К РИСУНКУ
        h1, // Подпись к рисунку, который как отельный параграф
        h2, // Подпись к рисунку в приложении
        h3, // Подпись к рисунку с пояснениями в основном тексте
        h4, // Подпись к рисунку с поясниениями в приложении
        i0, // ФОРМУЛА
        i1, // Формула в основном тексте
        i2, // Формула внутри параграфа
        j0, // ПОДПИСЬ ФОРМУЛЫ
        // НЕТ K (приложения)
        // НЕТ M (сноски)
        // НЕТ R (список литературы)
        // НЕТ N (SPACE)
        // НЕТ P (элемент листинга)
    }

    // Модель для результатов классификации
    public sealed class ClassificationResult
    {
        public int Id { get; set; }
        public ParagraphClass ParagraphClass { get; set; }
    } 


    // Результат проверки для параграфа
    public sealed class ParagraphCorrections
    {
        // ID параграфа (Его порядковый номер)
        public readonly int ParagraphID;
        // Тип параграфа
        public readonly ParagraphClass ParagraphClass;
        // Начало параграфа (20 символов)
        public readonly string Prefix;
        // Ошибки в параграфе
        public List<ParagraphMistake> Mistakes;

        public ParagraphCorrections(int paragraphID, ParagraphClass paragraphClass, string prefix, List<ParagraphMistake> mistakes)
        {
            ParagraphID = paragraphID;
            ParagraphClass = paragraphClass;
            Prefix = prefix;
            Mistakes = mistakes;
        }

        public static ParagraphCorrections TestParagraphCorrection
        {
            get
            {
                ParagraphCorrections testCorrection = new ParagraphCorrections(
                    paragraphID: 0,
                    paragraphClass: ParagraphClass.c0,
                    prefix: "Test prefix",
                    mistakes: new List<ParagraphMistake>
                    {
                        new ParagraphMistake(
                        message: "No mistake (PARAGRAPH)",
                        advice: "Nothing to advice"
                        )
                    }
                );
                return testCorrection;
            }
        }
    }
    public sealed class ParagraphMistake
    {
        // Сообщение об ошибке
        public readonly string Message;
        // Совет по исправлению
        public readonly string Advice;
        // Важность ошибки
        public readonly MistakeImportance Importance;

        public ParagraphMistake(string message, string advice = "Advice expected", MistakeImportance importance = MistakeImportance.Regular)
        {
            Message = message;
            Advice = advice;
            Importance = importance;
        }
    }
}