using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.PropertiesPuller
{
    public sealed class PagePropertiesGemBox : PageProperties
    {
        public int PageNumber { get; }
        public double Height { get; }
        public double Width { get; }

        public PagePropertiesGemBox(Word.DocumentModelPage page, int pageNumber)
        {
            PageNumber = pageNumber;
            Height = page.Height;
            Width = page.Width;
        }
    }
}
