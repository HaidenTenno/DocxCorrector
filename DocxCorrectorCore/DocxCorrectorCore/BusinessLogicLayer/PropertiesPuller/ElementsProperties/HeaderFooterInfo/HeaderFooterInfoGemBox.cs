using System.Collections.Generic;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.PropertiesPuller
{
    public enum HeaderFooterType
    {
        Header,
        Footer
    }

    public sealed class HeaderFooterInfoGemBox
    {
        public List<ParagraphPropertiesGemBox> HeaderFooterParagraphProperties { get; }

        public HeaderFooterInfoGemBox(Word.HeaderFooter headerFooter)
        {
            HeaderFooterParagraphProperties = new List<ParagraphPropertiesGemBox>();

            int paragraphID = 0;
            foreach(Word.Paragraph paragraph in headerFooter.GetChildElements(true, Word.ElementType.Paragraph))
            {
                HeaderFooterParagraphProperties.Add(new ParagraphPropertiesGemBox(paragraphID, paragraph));
                paragraphID++;
            }
        }
        
    }
}
