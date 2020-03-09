using System;
using Word = Microsoft.Office.Interop.Word;

namespace DocxCorrector.Models
{
    public sealed class PagePropertiesInterop : PageProperties
    {
        public PagePropertiesInterop(Word.PageSetup pageSetup, int pageNumber)
        {
            PageNumber = pageNumber;
            BottomMargin = pageSetup.BottomMargin;
            DifferentFirstPageHeaderFooter = Convert.ToBoolean(pageSetup.DifferentFirstPageHeaderFooter);
            FooterDistance = pageSetup.FooterDistance;
            Gutter = pageSetup.Gutter;
            HeaderDistance = pageSetup.HeaderDistance;
            LeftMargin = pageSetup.LeftMargin;
            LineNumbering = Convert.ToBoolean(pageSetup.LineNumbering.Active);
            MirrorMargins = Convert.ToBoolean(pageSetup.MirrorMargins);
            OddAndEvenPagesHeaderFooter = Convert.ToBoolean(pageSetup.OddAndEvenPagesHeaderFooter);
            Orientation = Convert.ToString(pageSetup.Orientation);
            PageHeight = pageSetup.PageHeight;
            PageWidth = pageSetup.PageWidth;
            PaperSize = Convert.ToString(pageSetup.PaperSize);
            RightMargin = pageSetup.RightMargin;
            SectionDirection = Convert.ToString(pageSetup.SectionDirection);
            SectionStart = Convert.ToString(pageSetup.SectionStart);
            TextColumns = pageSetup.TextColumns.Count;
            TopMargin = pageSetup.TopMargin;
            TwoPagesOnOne = pageSetup.TwoPagesOnOne;
            VerticalAlignment = Convert.ToString(pageSetup.VerticalAlignment);
        }
    }
}
