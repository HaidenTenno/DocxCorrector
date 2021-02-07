using System;
using System.Collections.Generic;
using DocxCorrectorCore.Services.Helpers;
using DocxCorrectorCore.Models.Corrections;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.PropertiesPuller
{
    public class ParagraphPropertiesTableZero: ParagraphPropertiesGemBox
    {
        // Public
        public bool ListFormatIsList { get; }
        public string? ListFormatListItem { get; }
        public string? ListFormatLevel { get; }
        public string? ListFormatAlignment { get; }
        public string? ListFormatNumberFormat { get; }
        public string? ListFormatNumberStyle { get; }

        private string? GetClassIfNeeded(Word.Paragraph paragraph)
        {
            string content = GemBoxHelper.GetParagraphContentWithSkippables(paragraph);
            if (paragraph.ListFormat.IsList) { return ParagraphClass.d2.ToString(); }

            return GetClassIfNeeded(content);
        }

        private string GetClassIfNeeded(string content)
        {
            Dictionary<string, string> keyValues = new Dictionary<string, string>
            {
                { GemBoxHelper.SkippableElements[Word.ElementType.Picture], ParagraphClass.g0.ToString() },
                { GemBoxHelper.SkippableElements[Word.ElementType.Chart], ParagraphClass.g0.ToString() },
                { GemBoxHelper.SkippableElements[Word.ElementType.Shape], ParagraphClass.g0.ToString() },
                { GemBoxHelper.SkippableElements[Word.ElementType.Table], ParagraphClass.e0.ToString() },
                { "!SPACE!", ParagraphClass.n0.ToString() }
            };

            return keyValues.TryGetValue(content, out var result) ? result : ""; 
        }
        public ParagraphPropertiesTableZero(int id, Word.Paragraph paragraph) : base(id, paragraph)
        {
            Content = GemBoxHelper.GetParagraphContentWithSkippables(paragraph);
            ListFormatIsList = paragraph.ListFormat.IsList;
            if (ListFormatIsList)
            {
                ListFormatListItem = paragraph.ListItem.ToString();
                ListFormatLevel = paragraph.ListFormat.ListLevelNumber.ToString();
                ListFormatAlignment = paragraph.ListFormat.ListLevelFormat.Alignment.ToString();
                ListFormatNumberFormat = paragraph.ListFormat.ListLevelFormat.NumberFormat.ToString();
                ListFormatNumberStyle = paragraph.ListFormat.ListLevelFormat.NumberStyle.ToString();
            }

            CurElementMark = GetClassIfNeeded(paragraph);
        }

        public ParagraphPropertiesTableZero(int id, string content) : base(id, content) 
        {
            Content = content;
            CurElementMark = GetClassIfNeeded(content);
        }
    }
}
