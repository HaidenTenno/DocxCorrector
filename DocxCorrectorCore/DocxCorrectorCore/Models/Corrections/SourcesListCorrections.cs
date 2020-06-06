using System.Collections.Generic;

namespace DocxCorrectorCore.Models.Corrections
{
    public sealed class SourcesListCorrections
    {
        // ID параграфа ЗАГОЛОВКА СПИСКА ЛИТЕРАТУРЫ (Его порядковый номер)
        public readonly int ParagraphID;
        // Начало параграфа (20 символов)
        public readonly string Prefix;
        // Ошибки в списке литературы
        public List<SourcesListMistake> Mistakes;

        public SourcesListCorrections(int paragraphID, string prefix, List<SourcesListMistake> mistakes)
        {
            ParagraphID = paragraphID;
            Prefix = prefix;
            Mistakes = mistakes;
        }

        public static SourcesListCorrections TestSourcesListCorrection
        {
            get
            {
                SourcesListCorrections testCorrection = new SourcesListCorrections(
                    paragraphID: 0,
                    prefix: "Test prefix",
                    mistakes: new List<SourcesListMistake>
                    {
                        new SourcesListMistake(
                        paragraphID: 0,
                        prefix: "Text element prefix",
                        message: "No mistake (SOURCES LIST)",
                        advice: "Nothing to advice"
                        )
                    }
                );
                return testCorrection;
            }
        }
    }

    public sealed class SourcesListMistake
    {
        // ID параграфа ЭЛЕМЕНТА СПИСКА ЛИТЕРАТУРЫ (Его порядковый номер)
        public readonly int ParagraphID;
        // Начало параграфа (20 символов)
        public readonly string Prefix;
        // Сообщение об ошибке
        public readonly string Message;
        // Совет по исправлению
        public readonly string Advice;
        // Важность ошибки
        public readonly MistakeImportance Importance;

        public SourcesListMistake(int paragraphID, string prefix, string message, string advice = "Advice expected", MistakeImportance importance = MistakeImportance.Regular)
        {
            ParagraphID = paragraphID;
            Prefix = prefix;
            Message = message;
            Advice = advice;
            Importance = importance;
        }
    }
}
