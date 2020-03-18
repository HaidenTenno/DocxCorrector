using System;

namespace DocxCorrector.Models.ElementsObjectModel
{
    public class Paragraph : DocumentElement
    {
        public override string[] Suffixes => new string[] { ".", "!", "?" };
    }

}
