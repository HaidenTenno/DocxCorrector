namespace DocxCorrector.Models
{
    public class PageProperties
    {
        public int PageNumber { get; set; }
        
//        Application - не нужно
//        BookFoldPrinting - не нужно
//        BookFoldPrintingSheets - не нужно
//        BookFoldRevPrinting - не нужно
//        Creator - не нужно
//        FirstPageTray - не нужно
//        GutterOnTop - не нужно
//        GutterPos - не нужно
//        GutterStyle - не нужно
//        OtherPagesTray - не нужно
//        Parent - не нужно
//        ShowGrid - не нужно
//        SuppressEndnotes - не нужно

//        CharsLine - ???СЕТКА ДОКУМЕНТА
//        LayoutMode - ???СЕТКА ДОКУМЕНТА
//        LinesPage - ???СЕТКА ДОКУМЕНТА

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
    }
}