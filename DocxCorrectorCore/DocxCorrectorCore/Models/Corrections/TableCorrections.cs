namespace DocxCorrectorCore.Models.Corrections
{
    public sealed class TableCorrections
    {
        // ID параграфа ЗАГОЛОВКА СПИСКА ЛИТЕРАТУРЫ (Его порядковый номер)
        public readonly int ParagraphID;
        // Начало параграфа (20 символов)
        public readonly string Prefix;
        // Сообщение об ошибке
        public string Message;

        public TableCorrections(int paragraphID, string prefix, string message)
        {
            ParagraphID = paragraphID;
            Prefix = prefix;
            Message = message;
        }

        public static TableCorrections TestTableCorrection
        {
            get
            {
                TableCorrections testCorrection = new TableCorrections(
                    paragraphID: 0,
                    prefix: "Test prefix",
                    message: "NO MISTAKE"
                );
                return testCorrection;
            }
        }
    }
}