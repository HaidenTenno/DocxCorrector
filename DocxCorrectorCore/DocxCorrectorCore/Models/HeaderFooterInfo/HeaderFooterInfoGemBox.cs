using System;
using System.Collections.Generic;
using Word = GemBox.Document;

namespace DocxCorrectorCore.Models
{
    public sealed class HeaderFooterInfoGemBox : HeaderFooterInfo
    {
        public List<ParagraphProperties> HeaderFooterParagraphProperties { get; }
        public HeaderFooterInfoGemBox(Word.HeaderFooter headerFooter)
        {
            HeaderFooterParagraphProperties = new List<ParagraphProperties>();
            foreach(Word.Paragraph paragraph in headerFooter.GetChildElements(true, Word.ElementType.Paragraph))
            {
                HeaderFooterParagraphProperties.Add(new ParagraphPropertiesGemBox(paragraph));
            }
        }
        
    }
}
