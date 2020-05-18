using System;

namespace DocxCorrectorCore.Models.ElementsObjectModel
{
    public class ParagraphBeforeList : Paragraph
    {
        //c2
        public override bool KeepWithNext => true;
        public override string[] Suffixes => new string[] { ":" };
    }
}
