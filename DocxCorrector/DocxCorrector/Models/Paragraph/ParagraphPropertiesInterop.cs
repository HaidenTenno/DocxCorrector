using System;
using Word = Microsoft.Office.Interop.Word;

namespace DocxCorrector.Models
{
    public sealed class ParagraphPropertiesInterop : ParagraphProperties
    {
        // Range
        public string Text { get; }
        public string Bold { get; }
        public string Italic { get; }
        public string Underline { get; }
        public string BoldBi { get; }
        public string Bookmarks { get; }
        public string Borders { get; }
        public string Case { get; }
        public string Characters { get; }
        public string CharacterWidth { get; }
        public string CombineCharacters { get; }
        public string ContentControls { get; }
        public string Creator { get; }
        public string DisableCharacterSpaceGrid { get; }
        public string Document { get; }
        public string Duplicate { get; }
        public string Editors { get; }
        public string EmpasisMark { get; }
        public string End { get; }
        public string EndnoteOptions { get; }
        public string Endnotes { get; }
        public string Fields { get; }
        public string Find { get; }
        public string FitTextWidth { get; }
        public string Footnotes { get; }
        public string FormattedText { get; }
        public string FormFields { get; }
        public string Frames { get; }
        public string GrammarChecked { get; }
        public string GrammaticalErrors { get; }
        public string HighlightColorIndex { get; }
        public string HorizontallnVertical { get; }
        public string HTMLDicisions { get; }
        public string Hyperlinks { get; }
        public string InlineShapes { get; }
        public string IsEndOfMark { get; }
        public string ItalicBi { get; }
        public string Kana { get; }
        public string LanguageDetected { get; }
        public string LanguageID { get; }
        public string LanguageIDFarEst { get; }
        public string LanguageIDOther { get; }
        public string ListFormat { get; }
        public string ListParagraphs { get; }
        public string NoProofing { get; }
        public string OMaths { get; }
        public string Orientation { get; }
        public string PageSetup { get; }
        public string ParagraphFormat { get; }
        public string Paragraphs { get; }
        public string PreviousBookmarkID { get; }
        public string ReadabilityStatistics { get; }
        public string Revisions { get; }
        public string Sections { get; }
        public string Sentenses { get; }
        public string Shading { get; }
        public string ShapeRange { get; }
        public string ShowAll { get; }
        public string SmartTags { get; }
        public string SpellingChecked { get; }
        public string SpellingErrors { get; }
        public string Subdocuments { get; }
        public string SynonymInfo { get; }
        public string Tables { get; }
        public string TextRetrievalMode { get; }
        public string TextVisibleOnScreen { get; }
        public string TopLevelTables { get; }
        public string TwoLinesInOne { get; }
        public string Words { get; }
        // Font
        public string FontName { get; }
        public string FontSize { get; }
        public string FontUnderlineColor { get; }
        public string FontStrikeThrough { get; }
        public string FontSuperscript { get; }
        public string FontSubscript { get; }
        public string FontHidden { get; }
        public string FontScaling { get; }
        public string FontPosition { get; }
        public string FontKerning { get; }
        public string FontAllCaps { get; }
        public string FontApplication { get; }
        public string FontBoldBi { get; }
        public string FontBorders { get; }
        public string FontColor { get; }
        public string FontColorIndex { get; }
        public string FontColorIndexBi { get; }
        public string FontContextualAlternates { get; }
        public string FontCreator { get; }
        public string FontDiacriricColor { get; }
        public string FontDoubleStrikeThrough { get; }
        public string FontDuplicate { get; }
        public string FontEmboss { get; }
        public string FontEmphasisMark { get; }
        public string FontEngrave { get; }
        public string FontItalic { get; }
        public string FontItalicBi { get; }
        public string FontLigatures { get; }
        public string FontNameAscii { get; }
        public string FontNameBi { get; }
        public string FontNameFarEast { get; }
        public string FontNameOther { get; }
        public string FontNumberForm { get; }
        public string FontNumberSpacing { get; }
        public string FontOutline { get; }
        public string FontShading { get; }
        public string FontShadow { get; }
        public string FontSizeBi { get; }
        public string FontSmallCaps { get; }
        public string FontStylisticSet { get; }
        public string FontUnderline { get; }
        // Paragraph
        public string OutlineLevel { get; }
        public string Alignment { get; }
        public string CharacterUnitLeftIndent { get; }
        public string LeftIndent { get; }
        public string CharacterUnitRightIndent { get; }
        public string RightIndent { get; }
        public string CharacterUnitFirstLineIndent { get; }
        public string MirrorIndents { get; }
        public string LineSpacing { get; }
        public string SpaceBefore { get; }
        public string SpaceAfter { get; }
        public string PageBreakBefore { get; }
        public string AddSpaceBetweenFarEastAndAlpha { get; }
        public string AddSpaceBetweenFarEastAndDigit { get; }
        public string Application { get; }
        public string AutoAdjustRightIndent { get; }
        public string BaseLineAlignment { get; }
        public string ParagraphBorders { get; }
        public string CollapsedState { get; }
        public string CollapseHEadingByDefault { get; }
        public string ParagraphCreator { get; }
        public string DisableLineHeightGrid { get; }
        public string DropCap { get; }
        public string FarEastLineBreakControl { get; }
        public string FirstLineIndent { get; }
        public string HalfWidthPunctuationOnTopOfLine { get; }
        public string HalfWidthPunctuation { get; }
        public string Hyphenation { get; }
        public string IsStyleSeparator { get; }
        public string KeepTogether { get; }
        public string KeepWithNext { get; }
        public string LineSpacingRule { get; }
        public string LineUnitAfter { get; }
        public string LineUnitBefore { get; }
        public string NoLineNumber { get; }
        public string ParagraphParent { get; }
        public string ReadingOrder { get; }
        public string ParagraphShading { get; }
        public string SpaceAfterAuto { get; }
        public string ParagraphStyle { get; }
        public string TabStops { get; }
        public string TextboxTightWrap { get; }
        public string TextID { get; }
        public string WindowControl { get; }
        public string WordWrap { get; }

        public ParagraphPropertiesInterop(Word.Paragraph paragraph)
        {
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
