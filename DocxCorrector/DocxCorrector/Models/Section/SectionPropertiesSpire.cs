using System;
using Word = Spire.Doc;

namespace DocxCorrector.Models
{
    public sealed class SectionPropertiesSpire : SectionProperties
    {
        public int SectionNumber { get; }
        public string ColumnsCount { get; }
        public string SectionBreakType { get; }
        public string EndnotePosition { get; }
        public string EndnoteNumberFormat { get; }
        public string EndnoteRestartRule { get; }
        public string EndnoteStartNumber { get; }
        public string FootnotePosition { get; }
        public string FootnoteNumberFormat { get; }
        public string FootnoteRestartRule { get; }
        public string FootnoteStartNumber { get; }
        public string TextDirection { get; }
        public string Bidi { get; }
        public string Borders { get; }
        public string MarginBottom { get; }
        public string MarginLeft { get; }
        public string MarginRight { get; }
        public string MarginTop { get; }
        public string Orientation { get; }
        public string ClientHeight { get; }
        public string ClientWidth { get; }
        public string FooterDistance { get; }
        public string HeaderDistance { get; }
        public string Height { get; }
        public string Width { get; }
        public string VerticalAlignment { get; }
        public string ColumnsLineBetween { get; }
        public string IsEqualColumnWidth { get; }
        public string LineNumberingStep { get; }
        public string PageNumberStyle { get; }
        public string PageStartingNumber { get; }
        public string RestartPageNumbering { get; }
        public string IsFrontPageBorder { get; }
        public string LineNumberingRestartMode { get; }
        public string LineNumberingStartValue { get; }
        public string PageBorderIncludeFooter { get; }
        public string PageBorderIncludeHeader { get; }
        public string PageBorderOffsetFrom { get; }
        public string PageBordersApplyType { get; }
        public string DifferentFirstPageHeaderFooter { get; }
        public string LineNumberingDistanceFromText { get; }
        public string DifferentOddAndEvenPagesHeaderFooter { get; }
        public SectionPropertiesSpire(Word.Section section, int sectionNumber)
        {
            SectionNumber = sectionNumber;
            ColumnsCount = section.Columns.Count.ToString();
            SectionBreakType = section.BreakCode.ToString();
            EndnotePosition = section.EndnoteOptions.Position.ToString();
            EndnoteNumberFormat = section.EndnoteOptions.NumberFormat.ToString();
            EndnoteRestartRule = section.EndnoteOptions.RestartRule.ToString();
            EndnoteStartNumber = section.EndnoteOptions.StartNumber.ToString();
            FootnotePosition = section.FootnoteOptions.Position.ToString();
            FootnoteNumberFormat = section.FootnoteOptions.NumberFormat.ToString();
            FootnoteRestartRule = section.FootnoteOptions.RestartRule.ToString();
            FootnoteStartNumber = section.FootnoteOptions.StartNumber.ToString();
            TextDirection = section.TextDirection.ToString();
            Bidi = section.PageSetup.Bidi.ToString();
            Borders = section.PageSetup.Borders.NoBorder.ToString();
            MarginBottom = section.PageSetup.Margins.Bottom.ToString();
            MarginLeft = section.PageSetup.Margins.Left.ToString();
            MarginRight = section.PageSetup.Margins.Right.ToString();
            MarginTop = section.PageSetup.Margins.Top.ToString();
            Orientation = section.PageSetup.Orientation.ToString();
            ClientHeight = section.PageSetup.ClientHeight.ToString(); // ??
            ClientWidth = section.PageSetup.ClientWidth.ToString(); // ??
            FooterDistance = section.PageSetup.FooterDistance.ToString();
            HeaderDistance = section.PageSetup.HeaderDistance.ToString();
            Height = section.PageSetup.PageSize.Height.ToString();
            Width = section.PageSetup.PageSize.Width.ToString();
            VerticalAlignment = section.PageSetup.VerticalAlignment.ToString();
            ColumnsLineBetween = section.PageSetup.ColumnsLineBetween.ToString();
            IsEqualColumnWidth = section.PageSetup.EqualColumnWidth.ToString();
            LineNumberingStep = section.PageSetup.LineNumberingStep.ToString();
            PageNumberStyle = section.PageSetup.PageNumberStyle.ToString();
            PageStartingNumber = section.PageSetup.PageStartingNumber.ToString();
            RestartPageNumbering = section.PageSetup.RestartPageNumbering.ToString();
            IsFrontPageBorder = section.PageSetup.IsFrontPageBorder.ToString(); // ??
            LineNumberingRestartMode = section.PageSetup.LineNumberingRestartMode.ToString();
            LineNumberingStartValue = section.PageSetup.LineNumberingStartValue.ToString();
            PageBorderIncludeFooter = section.PageSetup.PageBorderIncludeFooter.ToString();
            PageBorderIncludeHeader = section.PageSetup.PageBorderIncludeHeader.ToString();
            PageBorderOffsetFrom = section.PageSetup.PageBorderOffsetFrom.ToString();
            PageBordersApplyType = section.PageSetup.PageBordersApplyType.ToString();
            DifferentFirstPageHeaderFooter = section.PageSetup.DifferentFirstPageHeaderFooter.ToString();
            LineNumberingDistanceFromText = section.PageSetup.LineNumberingDistanceFromText.ToString();
            DifferentOddAndEvenPagesHeaderFooter = section.PageSetup.DifferentOddAndEvenPagesHeaderFooter.ToString();
        }
    }
}