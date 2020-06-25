using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;
using ServiceStack;

namespace DocxCorrectorCore.BusinessLogicLayer.PropertiesPuller
{
    public sealed class ParagraphPropertiesTableZero: ParagraphPropertiesGemBox
    {
        // Public
        public bool ListFormatIsList { get; }
        public string? ListFormatListItem { get; }
        public string? ListFormatLevel { get; }
        public string? ListFormatAlignment { get; }
        public string? ListFormatNumberFormat { get; }
        public string? ListFormatNumberStyle { get; }

        public override string Content { get; }

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
        }

        public ParagraphPropertiesTableZero(int id, string content) : base(id, content) 
        {
            Content = content;
        }
    }
}
