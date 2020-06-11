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
                            row: 0,
                            column: 0,
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
        // Строка
        public int Row;
        // Столбец
        public int Column;
        // Сообщение об ошибке
        public readonly string Message;
        // Совет по исправлению
        public readonly string Advice;
        // Важность ошибки
        public readonly MistakeImportance Importance;

        public TableMistake(int row, int column, string message, string advice = "Advice expected", MistakeImportance importance = MistakeImportance.Regular)
        {
            Row = row;
            Column = column;
            Message = message;
            Advice = advice;
            Importance = importance;
        }
    }
}