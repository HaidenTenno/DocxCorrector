using System.Collections.Generic;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.PropertiesPuller
{
    public sealed class SectionPropertiesGemBox : SectionProperties
    {
        public int SectionNumber { get; }
        public List<Dictionary<string,string>> HeadersFooters { get; }
        public Dictionary<string,string> PageSetup { get; }

        public SectionPropertiesGemBox(Word.Section section, int sectionNumber)
        {
            SectionNumber = sectionNumber;
            HeadersFooters = new List<Dictionary<string, string>>();
            foreach (var element in section.HeadersFooters)
            {
                Dictionary<string, string> headerFooter = new Dictionary<string, string>()
                {
                    { "Content", element.Content.ToString() },
                    { "HeaderFooterType", element.HeaderFooterType.ToString() },
                    { "IsHeader", element.IsHeader.ToString() },
                };
                HeadersFooters.Add(headerFooter);
            }
            PageSetup = new Dictionary<string, string>()
            {
                { "PageSetupLineNumberCountBy", section.PageSetup.LineNumberCountBy.ToString() },
                { "PageSetupLineNumberDistanceFromText", section.PageSetup.LineNumberDistanceFromText.ToString() },
                { "PageSetupLineNumberRestartSetting", section.PageSetup.LineNumberRestartSetting.ToString() },
                { "PageSetupLineStartingNumber", section.PageSetup.LineStartingNumber.ToString() },
                { "PageSetupOrientation", section.PageSetup.Orientation.ToString() },
                { "PageSetupPageColor", section.PageSetup.PageColor.ToString() },
                { "PageSetupPageHeight", section.PageSetup.PageHeight.ToString() },
                { "PageSetupPageNumberStyle", section.PageSetup.PageNumberStyle.ToString() },
                { "PageSetupPageWidth", section.PageSetup.PageWidth.ToString() },
                { "PageSetupPaperType", section.PageSetup.PaperType.ToString() },
                { "PageSetupRightToLeft", section.PageSetup.RightToLeft.ToString() },
            };
        }
    }
}
