using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.PropertiesPuller
{
    public sealed class PagePropertiesGemBox : PageProperties
    {
        // TODO: SET ONLY??
        public int PageNumber { get; set; }
        public double Height { get; set; }
        public double Width { get; set; }

        public PagePropertiesGemBox(Word.DocumentModelPage page, int pageNumber)
        {
            PageNumber = pageNumber;
            Height = page.Height;
            Width = page.Width;
        }
    }
}
