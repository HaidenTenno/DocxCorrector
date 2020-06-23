using System.Collections.Generic;

namespace DocxCorrectorCore.Models.Corrections
{
    public sealed class HeadlingCorrections
    {
        // Ошибки в заголовке
        public List<HeadlingMistake> Mistakes;

        public HeadlingCorrections(List<HeadlingMistake> mistakes)
        {
            Mistakes = mistakes;
        }

        public static HeadlingCorrections TestHeadlingCorrection
        {
            get
            {
                HeadlingCorrections testCorrection = new HeadlingCorrections(
                    mistakes: new List<HeadlingMistake>
                    {
                        new HeadlingMistake(
                        paragraphID: 0,
                        prefix: "Test prefix",
                        message: "No mistake (HEADLING)",
                        advice: "Nothing to advice"
                        )
                    }
                );
                return testCorrection;
            }
        }
    }

    public sealed class HeadlingMistake
    {
        // ID параграфа (Его порядковый номер)
        public readonly int ParagraphID;
        // Начало параграфа (20 символов)
        public readonly string Prefix;
        // Сообщение об ошибке
        public readonly string Message;
        // Совет по исправлению
        public readonly string Advice;
        // Важность ошибки
        public readonly MistakeImportance Importance;

        public HeadlingMistake(int paragraphID, string prefix, string message, string advice = "Advice expected", MistakeImportance importance = MistakeImportance.Regular)
        {
            ParagraphID = paragraphID;
            Prefix = prefix;
            Message = message;
            Advice = advice;
            Importance = importance;
        }
    }
}
