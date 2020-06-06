using System.Collections.Generic;

namespace DocxCorrectorCore.Models.Corrections
{
    public sealed class TableCorrections
    {
        // ID параграфа ТАБЛИЦЫ (Его порядковый номер)
        public readonly int ParagraphID;
        // Ошибки в таблице
        public List<TableMistake> Mistakes;

        public TableCorrections(int paragraphID, List<TableMistake> mistakes)
        {
            ParagraphID = paragraphID;
            Mistakes = mistakes;
        }

        public static TableCorrections TestTableCorrection
        {
            get
            {
                TableCorrections testCorrection = new TableCorrections(
                    paragraphID: 0,
                    mistakes: new List<TableMistake>
                    {
                        new TableMistake(
                            message: "No mistake (TABLE)",
                            advice: "Nothing to advice"
                        )
                    }
                );
                return testCorrection;
            }
        }
    }

    public sealed class TableMistake
    {
        // Сообщение об ошибке
        public readonly string Message;
        // Совет по исправлению
        public readonly string Advice;
        // Важность ошибки
        public readonly MistakeImportance Importance;

        public TableMistake(string message, string advice = "Advice expected", MistakeImportance importance = MistakeImportance.Regular)
        {
            Message = message;
            Advice = advice;
            Importance = importance;
        }
    }
}