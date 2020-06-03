
namespace DocxCorrectorCore.Models.Corrections
{
    public sealed class SourcesListCorrections
    {
        // ID параграфа ЗАГОЛОВКА СПИСКА ЛИТЕРАТУРЫ (Его порядковый номер)
        public readonly int ParagraphID;
        // Начало параграфа (20 символов)
        public readonly string Prefix;
        // Сообщение об ошибке
        public string Message;

        public SourcesListCorrections(int paragraphID, string prefix, string message)
        {
            ParagraphID = paragraphID;
            Prefix = prefix;
            Message = message;
        }

        public static SourcesListCorrections TestSourcesListCorrection
        {
            get
            {
                SourcesListCorrections testCorrection = new SourcesListCorrections(
                    paragraphID: 0,
                    prefix: "Test prefix",
                    message: "NO MISTAKE"
                );
                return testCorrection;
            }
        }
    }
}
