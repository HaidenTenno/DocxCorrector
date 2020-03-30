using System;

namespace DocxCorrector.Models.ElementsObjectModel
{
    public class ParagraphBeforeList : Paragraph
    {
        public override bool KeepWithNext => true;
        public override string[] Suffixes => new string[] { ":" };
    }

}
