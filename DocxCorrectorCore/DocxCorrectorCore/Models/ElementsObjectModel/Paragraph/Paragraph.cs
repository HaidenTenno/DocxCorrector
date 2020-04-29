using System;

namespace DocxCorrectorCore.Models.ElementsObjectModel
{
    public class Paragraph : DocumentElement
    {
        public override string[] Suffixes => new string[] { ".", "!", "?" };
    }

}
