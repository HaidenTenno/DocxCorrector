using System;
using Word = GemBox.Document;

namespace DocxCorrector.Models
{
    public sealed class ParagraphPropertiesGemBox : ParagraphProperties
    {
        public string CharacterFormatForParagraphMark { get; set; }
        public string Content { get; set; }
        public string Document { get; set; }
        public string ElementType { get; set; }
        public string Inlines { get; set; }
        public string ListFormat { get; set; }
        public string ListItem { get; set; }
        public string ParagraphFormat { get; set; }
        public string Parent { get; set; }
        public string ParentCollection { get; set; }

        public ParagraphPropertiesGemBox(Word.Paragraph paragraph)
        {
            CharacterFormatForParagraphMark = paragraph.CharacterFormatForParagraphMark.ToString();
            Content = paragraph.Content.ToString();
            Document = paragraph.Document.ToString();
            ElementType = paragraph.ElementType.ToString();
            Inlines = paragraph.Inlines.ToString();
            ListFormat = paragraph.ListFormat.ToString();
            try
            {
                ListItem = paragraph.ListItem.ToString();
            }
            catch (Exception ex)
            {
                ListItem = "NONE";
            }
            ParagraphFormat = paragraph.ParagraphFormat.ToString();
            Parent = paragraph.Parent.ToString();
            ParentCollection = paragraph.ParentCollection.ToString();
        }
    }
}
