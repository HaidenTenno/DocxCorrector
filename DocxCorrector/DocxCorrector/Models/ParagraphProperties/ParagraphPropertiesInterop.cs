using System;
using Word = Microsoft.Office.Interop.Word;

namespace DocxCorrector.Models
{
    // TODO: Убрать лишние поля и понять какие нужно раскрыть

    public sealed class ParagraphPropertiesInterop : ParagraphProperties
    {
        // Range
        public string Text { get; set; }
        public string Bold { get; set; }
        public string Italic { get; set; }
        public string Underline { get; set; }
        public string BoldBi { get; set; }
        public string Bookmarks { get; set; }
        public string Borders { get; set; }
        public string Case { get; set; }
        public string Characters { get; set; }
        public string CharacterWidth { get; set; }
        public string CombineCharacters { get; set; }
        public string ContentControls { get; set; }
        public string Creator { get; set; }
        public string DisableCharacterSpaceGrid { get; set; }
        public string Document { get; set; }
        public string Duplicate { get; set; }
        public string Editors { get; set; }
        public string EmpasisMark { get; set; }
        public string End { get; set; }
        public string EndnoteOptions { get; set; }
        public string Endnotes { get; set; }
        public string Fields { get; set; }
        public string Find { get; set; }
        public string FitTextWidth { get; set; }
        public string Footnotes { get; set; }
        public string FormattedText { get; set; }
        public string FormFields { get; set; }
        public string Frames { get; set; }
        public string GrammarChecked { get; set; }
        public string GrammaticalErrors { get; set; }
        public string HighlightColorIndex { get; set; }
        public string HorizontallnVertical { get; set; }
        public string HTMLDicisions { get; set; }
        public string Hyperlinks { get; set; }
        public string InlineShapes { get; set; }
        public string IsEndOfMark { get; set; }
        public string ItalicBi { get; set; }
        public string Kana { get; set; }
        public string LanguageDetected { get; set; }
        public string LanguageID { get; set; }
        public string LanguageIDFarEst { get; set; }
        public string LanguageIDOther { get; set; }
        public string ListFormat { get; set; }
        public string ListParagraphs { get; set; }
        public string NoProofing { get; set; }
        public string OMaths { get; set; }
        public string Orientation { get; set; }
        public string PageSetup { get; set; }
        public string ParagraphFormat { get; set; }
        public string Paragraphs { get; set; }
        public string PreviousBookmarkID { get; set; }
        public string ReadabilityStatistics { get; set; }
        public string Revisions { get; set; }
        public string Sections { get; set; }
        public string Sentenses { get; set; }
        public string Shading { get; set; }
        public string ShapeRange { get; set; }
        public string ShowAll { get; set; }
        public string SmartTags { get; set; }
        public string SpellingChecked { get; set; }
        public string SpellingErrors { get; set; }
        public string Subdocuments { get; set; }
        public string SynonymInfo { get; set; }
        public string Tables { get; set; }
        public string TextRetrievalMode { get; set; }
        public string TextVisibleOnScreen { get; set; }
        public string TopLevelTables { get; set; }
        public string TwoLinesInOne { get; set; }
        public string Words { get; set; }
        // Font
        public string FontName { get; set; }
        public string FontSize { get; set; }
        public string FontUnderlineColor { get; set; }
        public string FontStrikeThrough { get; set; }
        public string FontSuperscript { get; set; }
        public string FontSubscript { get; set; }
        public string FontHidden { get; set; }
        public string FontScaling { get; set; }
        public string FontPosition { get; set; }
        public string FontKerning { get; set; }
        public string FontAllCaps { get; set; }
        public string FontApplication { get; set; }
        public string FontBoldBi { get; set; }
        public string FontBorders { get; set; }
        public string FontColor { get; set; }
        public string FontColorIndex { get; set; }
        public string FontColorIndexBi { get; set; }
        public string FontContextualAlternates { get; set; }
        public string FontCreator { get; set; }
        public string FontDiacriricColor { get; set; }
        public string FontDoubleStrikeThrough { get; set; }
        public string FontDuplicate { get; set; }
        public string FontEmboss { get; set; }
        public string FontEmphasisMark { get; set; }
        public string FontEngrave { get; set; }
        public string FontItalic { get; set; }
        public string FontItalicBi { get; set; }
        public string FontLigatures { get; set; }
        public string FontNameAscii { get; set; }
        public string FontNameBi { get; set; }
        public string FontNameFarEast { get; set; }
        public string FontNameOther { get; set; }
        public string FontNumberForm { get; set; }
        public string FontNumberSpacing { get; set; }
        public string FontOutline { get; set; }
        public string FontShading { get; set; }
        public string FontShadow { get; set; }
        public string FontSizeBi { get; set; }
        public string FontSmallCaps { get; set; }
        public string FontStylisticSet { get; set; }
        public string FontUnderline { get; set; }
        // Paragraph
        public string OutlineLevel { get; set; }
        public string Alignment { get; set; }
        public string CharacterUnitLeftIndent { get; set; }
        public string LeftIndent { get; set; }
        public string CharacterUnitRightIndent { get; set; }
        public string RightIndent { get; set; }
        public string CharacterUnitFirstLineIndent { get; set; }
        public string MirrorIndents { get; set; }
        public string LineSpacing { get; set; }
        public string SpaceBefore { get; set; }
        public string SpaceAfter { get; set; }
        public string PageBreakBefore { get; set; }
        public string AddSpaceBetweenFarEastAndAlpha { get; set; }
        public string AddSpaceBetweenFarEastAndDigit { get; set; }
        public string Application { get; set; }
        public string AutoAdjustRightIndent { get; set; }
        public string BaseLineAlignment { get; set; }
        public string ParagraphBorders { get; set; }
        public string CollapsedState { get; set; }
        public string CollapseHEadingByDefault { get; set; }
        public string ParagraphCreator { get; set; }
        public string DisableLineHeightGrid { get; set; }
        public string DropCap { get; set; }
        public string FarEastLineBreakControl { get; set; }
        public string FirstLineIndent { get; set; }
        public string HalfWidthPunctuationOnTopOfLine { get; set; }
        public string HalfWidthPunctuation { get; set; }
        public string Hyphenation { get; set; }
        public string IsStyleSeparator { get; set; }
        public string KeepTogether { get; set; }
        public string KeepWithNext { get; set; }
        public string LineSpacingRule { get; set; }
        public string LineUnitAfter { get; set; }
        public string LineUnitBefore { get; set; }
        public string NoLineNumber { get; set; }
        public string ParagraphParent { get; set; }
        public string ReadingOrder { get; set; }
        public string ParagraphShading { get; set; }
        public string SpaceAfterAuto { get; set; }
        public string ParagraphStyle { get; set; }
        public string TabStops { get; set; }
        public string TextboxTightWrap { get; set; }
        public string TextID { get; set; }
        public string WindowControl { get; set; }
        public string WordWrap { get; set; }

        public ParagraphPropertiesInterop(Word.Paragraph paragraph)
        {
            if (paragraph == null) { return; }

            Text = paragraph.Range.Text.ToString();
            Underline = paragraph.Range.Underline.ToString();
            Bold = paragraph.Range.Bold.ToString();
            Italic = paragraph.Range.Italic.ToString();
            BoldBi = paragraph.Range.BoldBi.ToString();
            Bookmarks = paragraph.Range.Bookmarks.ToString();
            Borders = paragraph.Range.Borders.ToString();
            Case = paragraph.Range.Case.ToString();
            Characters = paragraph.Range.Characters.ToString();
            CharacterWidth = paragraph.Range.CharacterWidth.ToString();
            CombineCharacters = paragraph.Range.CombineCharacters.ToString();
            ContentControls = paragraph.Range.ContentControls.ToString();
            Creator = paragraph.Range.Creator.ToString();
            DisableCharacterSpaceGrid = paragraph.Range.DisableCharacterSpaceGrid.ToString();
            Document = paragraph.Range.Document.ToString();
            Duplicate = paragraph.Range.Duplicate.ToString();
            Editors = paragraph.Range.Editors.ToString();
            EmpasisMark = paragraph.Range.EmphasisMark.ToString();
            End = paragraph.Range.End.ToString();
            EndnoteOptions = paragraph.Range.EndnoteOptions.ToString();
            Endnotes = paragraph.Range.Endnotes.ToString();
            Fields = paragraph.Range.Fields.ToString();
            Find = paragraph.Range.Find.ToString();
            FitTextWidth = paragraph.Range.FitTextWidth.ToString();
            Footnotes = paragraph.Range.Footnotes.ToString();
            FormattedText = paragraph.Range.FormattedText.ToString();
            FormFields = paragraph.Range.FormFields.ToString();
            Frames = paragraph.Range.Frames.ToString();
            GrammarChecked = paragraph.Range.GrammarChecked.ToString();
            GrammaticalErrors = paragraph.Range.GrammaticalErrors.ToString();
            HighlightColorIndex = paragraph.Range.HighlightColorIndex.ToString();
            HorizontallnVertical = paragraph.Range.HorizontalInVertical.ToString();
            HTMLDicisions = paragraph.Range.HTMLDivisions.ToString();
            Hyperlinks = paragraph.Range.Hyperlinks.ToString();
            InlineShapes = paragraph.Range.InlineShapes.ToString();
            IsEndOfMark = paragraph.Range.IsEndOfRowMark.ToString();
            ItalicBi = paragraph.Range.ItalicBi.ToString();
            Kana = paragraph.Range.Kana.ToString();
            LanguageDetected = paragraph.Range.LanguageDetected.ToString();
            LanguageID = paragraph.Range.LanguageID.ToString();
            LanguageIDFarEst = paragraph.Range.LanguageIDFarEast.ToString();
            LanguageIDOther = paragraph.Range.LanguageIDOther.ToString();
            ListFormat = paragraph.Range.ListFormat.ToString();
            ListParagraphs = paragraph.Range.ListParagraphs.ToString();
            NoProofing = paragraph.Range.NoProofing.ToString();
            OMaths = paragraph.Range.OMaths.ToString();
            Orientation = paragraph.Range.Orientation.ToString();
            PageSetup = paragraph.Range.PageSetup.ToString();
            ParagraphFormat = paragraph.Range.ParagraphFormat.ToString();
            Paragraphs = paragraph.Range.Paragraphs.ToString();
            PreviousBookmarkID = paragraph.Range.PreviousBookmarkID.ToString();
            ReadabilityStatistics = paragraph.Range.ReadabilityStatistics.ToString();
            Revisions = paragraph.Range.Revisions.ToString();
            Sections = paragraph.Range.Sections.ToString();
            Sentenses = paragraph.Range.Sentences.ToString();
            Shading = paragraph.Range.Shading.ToString();
            ShapeRange = paragraph.Range.ShapeRange.ToString();
            ShowAll = paragraph.Range.ShowAll.ToString();
            SmartTags = paragraph.Range.SmartTags.ToString();
            SpellingChecked = paragraph.Range.SpellingChecked.ToString();
            SpellingErrors = paragraph.Range.SpellingErrors.ToString();
            Subdocuments = paragraph.Range.Subdocuments.ToString();
            SynonymInfo = paragraph.Range.SynonymInfo.ToString();
            Tables = paragraph.Range.Tables.ToString();
            TextRetrievalMode = paragraph.Range.TextRetrievalMode.ToString();
            TextVisibleOnScreen = paragraph.Range.TextVisibleOnScreen.ToString();
            TopLevelTables = paragraph.Range.TopLevelTables.ToString();
            TwoLinesInOne = paragraph.Range.TwoLinesInOne.ToString();
            Words = paragraph.Range.Words.ToString();
            FontName = paragraph.Range.Font.Name.ToString();
            FontSize = paragraph.Range.Font.Size.ToString();
            FontUnderlineColor = paragraph.Range.Font.UnderlineColor.ToString();
            FontStrikeThrough = paragraph.Range.Font.StrikeThrough.ToString();
            FontSuperscript = paragraph.Range.Font.Superscript.ToString();
            FontSubscript = paragraph.Range.Font.Superscript.ToString();
            FontHidden = paragraph.Range.Font.Hidden.ToString();
            FontScaling = paragraph.Range.Font.Scaling.ToString();
            FontPosition = paragraph.Range.Font.Position.ToString();
            FontKerning = paragraph.Range.Font.Kerning.ToString();
            FontApplication = paragraph.Range.Font.Application.ToString();
            FontBoldBi = paragraph.Range.Font.BoldBi.ToString();
            FontBorders = paragraph.Range.Font.Borders.ToString();
            FontColor = paragraph.Range.Font.Color.ToString();
            FontColorIndex = paragraph.Range.Font.ColorIndex.ToString();
            FontColorIndexBi = paragraph.Range.Font.ColorIndexBi.ToString();
            FontContextualAlternates = paragraph.Range.Font.ContextualAlternates.ToString();
            FontCreator = paragraph.Range.Font.Creator.ToString();
            FontDiacriricColor = paragraph.Range.Font.DiacriticColor.ToString();
            FontDoubleStrikeThrough = paragraph.Range.Font.DoubleStrikeThrough.ToString();
            FontDuplicate = paragraph.Range.Font.Duplicate.ToString();
            FontEmboss = paragraph.Range.Font.Emboss.ToString();
            FontEmphasisMark = paragraph.Range.Font.EmphasisMark.ToString();
            FontEngrave = paragraph.Range.Font.Engrave.ToString();
            FontItalic = paragraph.Range.Font.Italic.ToString();
            FontItalicBi = paragraph.Range.Font.ItalicBi.ToString();
            FontLigatures = paragraph.Range.Font.Ligatures.ToString();
            FontNameAscii = paragraph.Range.Font.NameAscii.ToString();
            FontNameBi = paragraph.Range.Font.NameBi.ToString();
            FontNameFarEast = paragraph.Range.Font.NameFarEast.ToString();
            FontNameOther = paragraph.Range.Font.NameOther.ToString();
            FontNumberForm = paragraph.Range.Font.NumberForm.ToString();
            FontNumberSpacing = paragraph.Range.Font.NumberSpacing.ToString();
            FontOutline = paragraph.Range.Font.Outline.ToString();
            FontShading = paragraph.Range.Font.Shading.ToString();
            FontShadow = paragraph.Range.Font.Shadow.ToString();
            FontSizeBi = paragraph.Range.Font.SizeBi.ToString();
            FontSmallCaps = paragraph.Range.Font.SmallCaps.ToString();
            FontStylisticSet = paragraph.Range.Font.StylisticSet.ToString();
            FontUnderline = paragraph.Range.Font.Underline.ToString();
            // Paragraph 
            OutlineLevel = paragraph.OutlineLevel.ToString();
            Alignment = paragraph.Alignment.ToString();
            CharacterUnitLeftIndent = paragraph.CharacterUnitLeftIndent.ToString();
            LeftIndent = paragraph.LeftIndent.ToString();
            CharacterUnitRightIndent = paragraph.CharacterUnitLeftIndent.ToString();
            RightIndent = paragraph.RightIndent.ToString();
            CharacterUnitFirstLineIndent = paragraph.CharacterUnitFirstLineIndent.ToString();
            MirrorIndents = paragraph.MirrorIndents.ToString();
            LineSpacing = paragraph.LineSpacing.ToString();
            SpaceBefore = paragraph.SpaceBefore.ToString();
            SpaceAfter = paragraph.SpaceAfter.ToString();
            PageBreakBefore = paragraph.PageBreakBefore.ToString();
            AddSpaceBetweenFarEastAndAlpha = paragraph.AddSpaceBetweenFarEastAndAlpha.ToString();
            AddSpaceBetweenFarEastAndDigit = paragraph.AddSpaceBetweenFarEastAndDigit.ToString();
            Application = paragraph.Application.ToString();
            AutoAdjustRightIndent = paragraph.AutoAdjustRightIndent.ToString();
            BaseLineAlignment = paragraph.BaseLineAlignment.ToString();
            ParagraphBorders = paragraph.Borders.ToString();
            CollapsedState = paragraph.CollapsedState.ToString();
            CollapseHEadingByDefault = paragraph.CollapseHeadingByDefault.ToString();
            ParagraphCreator = paragraph.Creator.ToString();
            DisableLineHeightGrid = paragraph.DisableLineHeightGrid.ToString();
            DropCap = paragraph.DropCap.ToString();
            FarEastLineBreakControl = paragraph.FarEastLineBreakControl.ToString();
            FirstLineIndent = paragraph.FirstLineIndent.ToString();
            HalfWidthPunctuationOnTopOfLine = paragraph.HalfWidthPunctuationOnTopOfLine.ToString();
            HalfWidthPunctuation = paragraph.HalfWidthPunctuationOnTopOfLine.ToString();
            Hyphenation = paragraph.Hyphenation.ToString();
            IsStyleSeparator = paragraph.IsStyleSeparator.ToString();
            KeepTogether = paragraph.KeepTogether.ToString();
            KeepWithNext = paragraph.KeepWithNext.ToString();
            LineSpacingRule = paragraph.LineSpacingRule.ToString();
            LineUnitAfter = paragraph.LineUnitAfter.ToString();
            LineUnitBefore = paragraph.LineUnitBefore.ToString();
            NoLineNumber = paragraph.NoLineNumber.ToString();
            ParagraphParent = paragraph.Parent.ToString();
            ReadingOrder = paragraph.ReadingOrder.ToString();
            ParagraphShading = paragraph.Shading.ToString();
            SpaceAfterAuto = paragraph.SpaceAfter.ToString();
            ParagraphStyle = paragraph.get_Style().ToString();
            TabStops = paragraph.TabStops.ToString();
            TextboxTightWrap = paragraph.TextboxTightWrap.ToString();
            TextID = paragraph.TextID.ToString();
            WindowControl = paragraph.WidowControl.ToString();
            WordWrap = paragraph.WordWrap.ToString();
        }
    }
}
