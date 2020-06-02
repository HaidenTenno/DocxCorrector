using System;
using System.Collections.Generic;
using System.Text;

namespace DocxCorrectorCore.Models
{
    public sealed class SourcesListCorrections
    {
        // ID параграфа (Его порядковый номер)
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
