using System;
using Word = GemBox.Document;

// TODO: Remove
namespace DocxCorrectorCore.BusinessLogicLayer.FixDocument
{
    public sealed class FixedDocument
    {
        public readonly Word.DocumentModel? Document;
        public readonly string Info;

        public FixedDocument(Word.DocumentModel? document, string info)
        {
            Document = document;
            Info = info;
        }
    }
}
