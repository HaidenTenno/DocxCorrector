using System;

namespace DocxCorrectorCore.Models.ElementsObjectModel
{
    public class ParagrapghBeforeEquation : Paragraph
    {
        //c3
        public override bool KeepWithNext => true;
    }
}