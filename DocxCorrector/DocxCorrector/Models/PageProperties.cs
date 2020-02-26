using System;
using Word = Microsoft.Office.Interop.Word;

namespace DocxCorrector.Models
{
    public class PageProperties
    {
        public int PageNumber { get; set; }
        public float BottomMargin { get; set; }
        public bool DifferentFirstPageHeaderFooter { get; set; } // int
        public float FooterDistance { get; set; }
        public float Gutter { get; set; }
        public float HeaderDistance { get; set; }
        public float LeftMargin { get; set; }
        public bool LineNumbering { get; set; }
        public bool MirrorMargins { get; set; } // int
        public bool OddAndEvenPagesHeaderFooter { get; set; }
        public string Orientation { get; set; }
        public float PageHeight { get; set; }
        public float PageWidth { get; set; }
        public string PaperSize { get; set; }
        public float RightMargin { get; set; }
        public string SectionDirection { get; set; }
        public string SectionStart { get; set; }
        public int TextColumns { get; set; }
        public float TopMargin { get; set; }
        public bool TwoPagesOnOne { get; set; }
        public string VerticalAlignment { get; set; }

        public PageProperties (Word.PageSetup pageSetup, int pageNumber)
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