using System;
using System.Collections.Generic;
using DocxCorrector.Models;
using Word = Microsoft.Office.Interop.Word;

namespace DocxCorrector.Services.Corrector
{
    class InteropExeption: Exception
    {
        public InteropExeption(string message) : base(message){ }
    } 

    public sealed class CorrectorInterop : Corrector
    {
        // Private
        private Word.Application App;
        private Word.Document Document;

        // Приготовится к началу работы
        private void OpenApp()
        {
            try
            {
                if (App != null) { CloseApp(); }
                App = new Word.Application { Visible = false };
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
                CloseApp();
#endif
            }
        }

        // Приготовится к окончанию работы
        private void CloseApp()
        {
            try
            {
                if (App != null) { App.Quit(); }
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
#endif
            }
        }

        // Открыть документ
        private void OpenDocument()
        {
            try
            {
                Document = App.Documents.Open(FileName: FilePath, Visible: true);
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
#endif
                throw new InteropExeption(message: "Can't open document");
            }
        }

        // Закрыть документ
        private void CloseDocument()
        {
            try
            {
                if (Document != null) { App.Documents.Close(); }
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
#endif
            }
        }
        
        // Закрыть документ без сохранения
        private void CloseDocumentWithoutSavingChanges()
        {
            try
            {
                if (Document != null) { Document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); }
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
#endif
            }
        }

        // Получить свойства параграфа
        private ParagraphProperties GetParagraphProperties(Word.Paragraph paragraph)
        {
            ParagraphProperties paragraphProperties = new ParagraphProperties
            {
                // Range
                Text = paragraph.Range.Text.ToString(),
                Underline = paragraph.Range.Underline.ToString(),
                Bold = paragraph.Range.Bold.ToString(),
                Italic = paragraph.Range.Italic.ToString(),
                BoldBi = paragraph.Range.BoldBi.ToString(),
                Bookmarks = paragraph.Range.Bookmarks.ToString(),
                Borders = paragraph.Range.Borders.ToString(),
                Case = paragraph.Range.Case.ToString(),
                Characters = paragraph.Range.Characters.ToString(),
                CharacterWidth = paragraph.Range.CharacterWidth.ToString(),
                CombineCharacters = paragraph.Range.CombineCharacters.ToString(),
                ContentControls = paragraph.Range.ContentControls.ToString(),
                Creator = paragraph.Range.Creator.ToString(),
                DisableCharacterSpaceGrid = paragraph.Range.DisableCharacterSpaceGrid.ToString(),
                Document = paragraph.Range.Document.ToString(),
                Duplicate = paragraph.Range.Duplicate.ToString(),
                Editors = paragraph.Range.Editors.ToString(),
                EmpasisMark = paragraph.Range.EmphasisMark.ToString(),
                End = paragraph.Range.End.ToString(),
                EndnoteOptions = paragraph.Range.EndnoteOptions.ToString(),
                Endnotes = paragraph.Range.Endnotes.ToString(),
                Fields = paragraph.Range.Fields.ToString(),
                Find = paragraph.Range.Find.ToString(),
                FitTextWidth = paragraph.Range.FitTextWidth.ToString(),
                Footnotes = paragraph.Range.Footnotes.ToString(),
                FormattedText = paragraph.Range.FormattedText.ToString(),
                FormFields = paragraph.Range.FormFields.ToString(),
                Frames = paragraph.Range.Frames.ToString(),
                GrammarChecked = paragraph.Range.GrammarChecked.ToString(),
                GrammaticalErrors = paragraph.Range.GrammaticalErrors.ToString(),
                HighlightColorIndex = paragraph.Range.HighlightColorIndex.ToString(),
                HorizontallnVertical = paragraph.Range.HorizontalInVertical.ToString(),
                HTMLDicisions = paragraph.Range.HTMLDivisions.ToString(),
                Hyperlinks = paragraph.Range.Hyperlinks.ToString(),
                InlineShapes = paragraph.Range.InlineShapes.ToString(),
                IsEndOfMark = paragraph.Range.IsEndOfRowMark.ToString(),
                ItalicBi = paragraph.Range.ItalicBi.ToString(),
                Kana = paragraph.Range.Kana.ToString(),
                LanguageDetected = paragraph.Range.LanguageDetected.ToString(),
                LanguageID = paragraph.Range.LanguageID.ToString(),
                LanguageIDFarEst = paragraph.Range.LanguageIDFarEast.ToString(),
                LanguageIDOther = paragraph.Range.LanguageIDOther.ToString(),
                ListFormat = paragraph.Range.ListFormat.ToString(),
                ListParagraphs = paragraph.Range.ListParagraphs.ToString(),
                NoProofing = paragraph.Range.NoProofing.ToString(),
                OMaths = paragraph.Range.OMaths.ToString(),
                Orientation = paragraph.Range.Orientation.ToString(),
                PageSetup = paragraph.Range.PageSetup.ToString(),
                ParagraphFormat = paragraph.Range.ParagraphFormat.ToString(),
                Paragraphs = paragraph.Range.Paragraphs.ToString(),
                PreviousBookmarkID = paragraph.Range.PreviousBookmarkID.ToString(),
                ReadabilityStatistics = paragraph.Range.ReadabilityStatistics.ToString(),
                Revisions = paragraph.Range.Revisions.ToString(),
                Sections = paragraph.Range.Sections.ToString(),
                Sentenses = paragraph.Range.Sentences.ToString(),
                Shading = paragraph.Range.Shading.ToString(),
                ShapeRange = paragraph.Range.ShapeRange.ToString(),
                ShowAll = paragraph.Range.ShowAll.ToString(),
                SmartTags = paragraph.Range.SmartTags.ToString(),
                SpellingChecked = paragraph.Range.SpellingChecked.ToString(),
                SpellingErrors = paragraph.Range.SpellingErrors.ToString(),
                Subdocuments = paragraph.Range.Subdocuments.ToString(),
                SynonymInfo = paragraph.Range.SynonymInfo.ToString(),
                Tables = paragraph.Range.Tables.ToString(),
                TextRetrievalMode = paragraph.Range.TextRetrievalMode.ToString(),
                TextVisibleOnScreen = paragraph.Range.TextVisibleOnScreen.ToString(),
                TopLevelTables = paragraph.Range.TopLevelTables.ToString(),
                TwoLinesInOne = paragraph.Range.TwoLinesInOne.ToString(),
                //WordOpenXML = paragraph.Range.WordOpenXML.ToString(),
                Words = paragraph.Range.Words.ToString(),
                //XML = paragraph.Range.XML.ToString(),
                //XMLNodes = paragraph.Range.XMLNodes.ToString(),
                // Font
                FontName = paragraph.Range.Font.Name.ToString(),
                FontSize = paragraph.Range.Font.Size.ToString(),
                FontTextColorRGB = paragraph.Range.Font.TextColor.RGB.ToString(),
                FontUnderlineColor = paragraph.Range.Font.UnderlineColor.ToString(),
                FontStrikeThrough = paragraph.Range.Font.StrikeThrough.ToString(),
                FontSuperscript = paragraph.Range.Font.Superscript.ToString(),
                FontSubscript = paragraph.Range.Font.Superscript.ToString(),
                FontHidden = paragraph.Range.Font.Hidden.ToString(),
                FontScaling = paragraph.Range.Font.Scaling.ToString(),
                FontPosition = paragraph.Range.Font.Position.ToString(),
                FontKerning = paragraph.Range.Font.Kerning.ToString(),
                FontApplication = paragraph.Range.Font.Application.ToString(),
                FontBoldBi = paragraph.Range.Font.BoldBi.ToString(),
                FontBorders = paragraph.Range.Font.Borders.ToString(),
                FontColor = paragraph.Range.Font.Color.ToString(),
                FontColorIndex = paragraph.Range.Font.ColorIndex.ToString(),
                FontColorIndexBi = paragraph.Range.Font.ColorIndexBi.ToString(),
                FontContextualAlternates = paragraph.Range.Font.ContextualAlternates.ToString(),
                FontCreator = paragraph.Range.Font.Creator.ToString(),
                FontDiacriricColor = paragraph.Range.Font.DiacriticColor.ToString(),
                FontDoubleStrikeThrough = paragraph.Range.Font.DoubleStrikeThrough.ToString(),
                FontDuplicate = paragraph.Range.Font.Duplicate.ToString(),
                FontEmboss = paragraph.Range.Font.Emboss.ToString(),
                FontEmphasisMark = paragraph.Range.Font.EmphasisMark.ToString(),
                FontEngrave = paragraph.Range.Font.Engrave.ToString(),
                // Обратить внимание
                FontFill = paragraph.Range.Font.Fill.ForeColor.RGB.ToString(),
                FontGlow = paragraph.Range.Font.Glow.ToString(),
                FontItalic = paragraph.Range.Font.Italic.ToString(),
                FontItalicBi = paragraph.Range.Font.ItalicBi.ToString(),
                FontLigatures = paragraph.Range.Font.Ligatures.ToString(),
                FontNameAscii = paragraph.Range.Font.NameAscii.ToString(),
                FontNameBi = paragraph.Range.Font.NameBi.ToString(),
                FontNameFarEast = paragraph.Range.Font.NameFarEast.ToString(),
                FontNameOther = paragraph.Range.Font.NameOther.ToString(),
                FontNumberForm = paragraph.Range.Font.NumberForm.ToString(),
                FontNumberSpacing = paragraph.Range.Font.NumberSpacing.ToString(),
                FontOutline = paragraph.Range.Font.Outline.ToString(),
                FontParent = paragraph.Range.Font.Parent.ToString(),
                FontReflection = paragraph.Range.Font.Reflection.ToString(),
                FontShading = paragraph.Range.Font.Shading.ToString(),
                FontShadow = paragraph.Range.Font.Shadow.ToString(),
                FontSizeBi = paragraph.Range.Font.SizeBi.ToString(),
                FontSmallCaps = paragraph.Range.Font.SmallCaps.ToString(),
                FontStylisticSet = paragraph.Range.Font.StylisticSet.ToString(),
                FontTextShadow = paragraph.Range.Font.TextShadow.ToString(),
                FontThreeD = paragraph.Range.Font.ThreeD.ToString(),
                FontUnderline = paragraph.Range.Font.Underline.ToString(),
                // Paragraph 
                OutlineLevel = paragraph.OutlineLevel.ToString(),
                Alignment = paragraph.Alignment.ToString(),
                CharacterUnitLeftIndent = paragraph.CharacterUnitLeftIndent.ToString(),
                LeftIndent = paragraph.LeftIndent.ToString(),
                CharacterUnitRightIndent = paragraph.CharacterUnitLeftIndent.ToString(),
                RightIndent = paragraph.RightIndent.ToString(),
                CharacterUnitFirstLineIndent = paragraph.CharacterUnitFirstLineIndent.ToString(),
                MirrorIndents = paragraph.MirrorIndents.ToString(),
                LineSpacing = paragraph.LineSpacing.ToString(),
                SpaceBefore = paragraph.SpaceBefore.ToString(),
                SpaceAfter = paragraph.SpaceAfter.ToString(),
                PageBreakBefore = paragraph.PageBreakBefore.ToString(),
                AddSpaceBetweenFarEastAndAlpha = paragraph.AddSpaceBetweenFarEastAndAlpha.ToString(),
                AddSpaceBetweenFarEastAndDigit = paragraph.AddSpaceBetweenFarEastAndDigit.ToString(),
                Application = paragraph.Application.ToString(),
                AutoAdjustRightIndent = paragraph.AutoAdjustRightIndent.ToString(),
                BaseLineAlignment = paragraph.BaseLineAlignment.ToString(),
                ParagraphBorders = paragraph.Borders.ToString(),
                CollapsedState = paragraph.CollapsedState.ToString(),
                CollapseHEadingByDefault = paragraph.CollapseHeadingByDefault.ToString(),
                ParagraphCreator = paragraph.Creator.ToString(),
                DisableLineHeightGrid = paragraph.DisableLineHeightGrid.ToString(),
                DropCap = paragraph.DropCap.ToString(),
                FarEastLineBreakControl = paragraph.FarEastLineBreakControl.ToString(),
                FirstLineIndent = paragraph.FirstLineIndent.ToString(),
                HalfWidthPunctuationOnTopOfLine = paragraph.HalfWidthPunctuationOnTopOfLine.ToString(),
                HalfWidthPunctuation = paragraph.HalfWidthPunctuationOnTopOfLine.ToString(),
                Hyphenation = paragraph.Hyphenation.ToString(),
                IsStyleSeparator = paragraph.IsStyleSeparator.ToString(),
                KeepTogether = paragraph.KeepTogether.ToString(),
                KeepWithNext = paragraph.KeepWithNext.ToString(),
                LineSpacingRule = paragraph.LineSpacingRule.ToString(),
                LineUnitAfter = paragraph.LineUnitAfter.ToString(),
                LineUnitBefore = paragraph.LineUnitBefore.ToString(),
                NoLineNumber = paragraph.NoLineNumber.ToString(),
                ParagraphParent = paragraph.Parent.ToString(),
                ReadingOrder = paragraph.ReadingOrder.ToString(),
                ParagraphShading = paragraph.Shading.ToString(),
                SpaceAfterAuto = paragraph.SpaceAfter.ToString(),
                ParagraphStyle = paragraph.get_Style().ToString(),
                TabStops = paragraph.TabStops.ToString(),
                TextboxTightWrap = paragraph.TextboxTightWrap.ToString(),
                TextID = paragraph.TextID.ToString(),
                WindowControl = paragraph.WidowControl.ToString(),
                WordWrap = paragraph.WordWrap.ToString()
            };
            return paragraphProperties;
        }
        
        // Получить свойства страницы
        private PageProperties GetSinglePageProperties(Word.PageSetup pageSetup, int pageNumber)
        {
            PageProperties result = new PageProperties
            {
                PageNumber = pageNumber,
                BottomMargin = pageSetup.BottomMargin,
                DifferentFirstPageHeaderFooter = Convert.ToBoolean(pageSetup.DifferentFirstPageHeaderFooter),
                FooterDistance = pageSetup.FooterDistance,
                Gutter = pageSetup.Gutter,
                HeaderDistance = pageSetup.HeaderDistance,
                LeftMargin = pageSetup.LeftMargin,
                LineNumbering = Convert.ToBoolean(pageSetup.LineNumbering.Active),
                MirrorMargins = Convert.ToBoolean(pageSetup.MirrorMargins),
                OddAndEvenPagesHeaderFooter = Convert.ToBoolean(pageSetup.OddAndEvenPagesHeaderFooter),
                Orientation = Convert.ToString(pageSetup.Orientation),
                PageHeight = pageSetup.PageHeight,
                PageWidth = pageSetup.PageWidth,
                PaperSize = Convert.ToString(pageSetup.PaperSize),
                RightMargin = pageSetup.RightMargin,
                SectionDirection = Convert.ToString(pageSetup.SectionDirection),
                SectionStart = Convert.ToString(pageSetup.SectionStart),
                TextColumns = pageSetup.TextColumns.Count,
                TopMargin = pageSetup.TopMargin,
                TwoPagesOnOne = pageSetup.TwoPagesOnOne,
                VerticalAlignment = Convert.ToString(pageSetup.VerticalAlignment)
            };
            
            return result;
        }

        // Проверить, что первый символ абзаца принадлежит множеству символов
        private int CheckIfFirstSymbolOfParagraphIs(Word.Paragraph paragraph, string[] symbols)
        {
            return Array.IndexOf(symbols, paragraph.Range.Text[0].ToString()) != -1 ? 1 : 0;
        }

        // Проверить, что последний символ абзаца принадлежит можнеству символов
        private int CheckIfLastSymbolOfParagraphIs(Word.Paragraph paragraph, string[] symbols)
        {
            if (paragraph.Range.Text.Length > 1)
            {
                return Array.IndexOf(symbols, paragraph.Range.Text[paragraph.Range.Text.Length - 2].ToString()) != -1 ? 1 : 0;
            }
            else
            {
                return CheckIfFirstSymbolOfParagraphIs(paragraph, symbols);
            }
        }

        // Проверить, что параграф содержит хотя бы один из символов
        private int CheckIfParagraphsContainsOneOf(Word.Paragraph paragraph, string[] symbols)
        {
            foreach (string symbol in symbols)
            {
                if (paragraph.Range.Text.Contains(symbol))
                {
                    return 1;
                }
            }
            return 0;
        }

        // Corrector
        public CorrectorInterop(string filePath = null) : base(filePath) { }

        // Получение JSON-а со списком ошибок
        public override List<ParagraphResult> GetMistakes()
        {
            try
            {
                OpenApp();
                OpenDocument();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                CloseApp();
                return new List<ParagraphResult>();
            }

            List<ParagraphResult> paragraphResults = new List<ParagraphResult>();

            // TODO: - Remove
            ParagraphResult testResult = new ParagraphResult
            {
                ParagraphID = 0,
                Type = ElementType.Paragraph,
                Prefix = "TestParagraph",
                Mistakes = new List<Mistake> { new Mistake { Message = "Русские буквы" } }
            };
            paragraphResults.Add(testResult);

            // TODO: - Implement method

            CloseApp();

            return paragraphResults;
        }

        // Получить свойства всех параграфов
        public override List<ParagraphProperties> GetAllParagraphsProperties()
        {
            try
            {
                OpenApp();
                OpenDocument();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                CloseApp();
                return new List<ParagraphProperties>();
            }

            List<ParagraphProperties> allParagraphsProperties = new List<ParagraphProperties>();

            foreach (Word.Paragraph paragraph in Document.Paragraphs)
            {
                ParagraphProperties paragraphProperties = GetParagraphProperties(paragraph);
                allParagraphsProperties.Add(paragraphProperties);
            }

            CloseApp();

            return allParagraphsProperties;
        }
        
        //Получить свойства всех страниц
        public override List<PageProperties> GetAllPagesProperties()
        {
            List<PageProperties> result = new List<PageProperties>();
            try
            {
                OpenApp();
                OpenDocument();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                CloseApp();
            }
            
            int totalPageNumber = Document.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
            for (int i = 1; i <= totalPageNumber; i++)
            {
                Word.Range pageRange = Document.Range().GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToAbsolute, i);
                PageProperties currentPageProperties = GetSinglePageProperties(pageRange.PageSetup, i);
                result.Add(currentPageProperties);
            }

            CloseDocumentWithoutSavingChanges();
            CloseApp();
            return result;
        }

        // Получить нормализованные свойства параграфов (Для классификатора Ромы)
        public override List<NormalizedProperties> GetNormalizedProperties()
        {
            try
            {
                OpenApp();
                OpenDocument();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                CloseApp();
                return new List<NormalizedProperties>();
            }

            List<NormalizedProperties> allNormalizedProperties = new List<NormalizedProperties>();

            int iteration = 0;
            foreach (Word.Paragraph paragraph in Document.Paragraphs)
            {
                int id = iteration;
                float firstLineIndent = paragraph.FirstLineIndent;
                NormalizedAligment aligment;
                switch (paragraph.Alignment)
                {
                    case Word.WdParagraphAlignment.wdAlignParagraphLeft:
                        aligment = NormalizedAligment.Left;
                        break;
                    case Word.WdParagraphAlignment.wdAlignParagraphCenter:
                        aligment = NormalizedAligment.Center;
                        break;
                    case Word.WdParagraphAlignment.wdAlignParagraphRight:
                        aligment = NormalizedAligment.Right;
                        break;
                    case Word.WdParagraphAlignment.wdAlignParagraphJustify:
                        aligment = NormalizedAligment.Justify;
                        break;
                    default:
                        aligment = NormalizedAligment.Other;
                        break;
                }
                int prefixIsNumber = Char.IsDigit(paragraph.Range.Text[0]) ? 1 : 0;
                int prefixIsLowercase = Char.IsLower(paragraph.Range.Text[0]) ? 1 : 0;
                int prefixIsUppercase = Char.IsUpper(paragraph.Range.Text[0]) ? 1 : 0;
                string[] dashes = new string[] { "-", "־", "᠆", "‐", "‑", "‒", "–", "—", "―", "﹘", "﹣", "－" };
                int prefixIsDash = CheckIfFirstSymbolOfParagraphIs(paragraph, dashes);
                string[] endSigns = new string[] { ".", "!", "?" };
                int suffixIsEndSign = CheckIfLastSymbolOfParagraphIs(paragraph, endSigns);
                string[] colon = new string[] { ":" };
                int suffixIsColon = CheckIfLastSymbolOfParagraphIs(paragraph, colon);
                string[] commaAndSemicolon = new string[] { ",", ";" };
                int suffixIsCommaOrSemicolon = CheckIfLastSymbolOfParagraphIs(paragraph, commaAndSemicolon);
                int containsDash = CheckIfParagraphsContainsOneOf(paragraph, dashes);
                string[] bracket = new string[] { ")" };
                int containsBracket = CheckIfParagraphsContainsOneOf(paragraph, bracket);
                float fontSize = paragraph.Range.Font.Size;
                float lineSpacing = paragraph.LineSpacing;
                LineSpacingRule lineSpacingRule;
                switch (paragraph.LineSpacingRule)
                {
                    case Word.WdLineSpacing.wdLineSpaceSingle:
                        lineSpacingRule = LineSpacingRule.Single;
                        break;
                    case Word.WdLineSpacing.wdLineSpace1pt5:
                        lineSpacingRule = LineSpacingRule.OneAndHalf;
                        break;
                    case Word.WdLineSpacing.wdLineSpaceDouble:
                        lineSpacingRule = LineSpacingRule.Double;
                        break;
                    case Word.WdLineSpacing.wdLineSpaceMultiple:
                        lineSpacingRule = LineSpacingRule.Miltiply;
                        break;
                    default:
                        lineSpacingRule = LineSpacingRule.Other;
                        break;
                }
                ContainsStatus italic;
                switch (paragraph.Range.Italic)
                {
                    case -1:
                        italic = ContainsStatus.Full;
                        break;
                    case 0:
                        italic = ContainsStatus.None;
                        break;
                    default:
                        italic = ContainsStatus.Contains;
                        break;
                }
                ContainsStatus bold;
                switch (paragraph.Range.Bold)
                {
                    case -1:
                        bold = ContainsStatus.Full;
                        break;
                    case 0:
                        bold = ContainsStatus.None;
                        break;
                    default:
                        bold = ContainsStatus.Contains;
                        break;
                }
                int blackColor = (paragraph.Range.Font.Color == Word.WdColor.wdColorBlack) || (paragraph.Range.Font.Color == Word.WdColor.wdColorAutomatic) ? 1 : 0;

                NormalizedProperties normalizedParagraphProperties = new NormalizedProperties
                {
                    Id = id,
                    FirstLineIndent = firstLineIndent,
                    Aligment = (int)aligment,
                    SymbolsCount = paragraph.Range.Text.Length,
                    PrefixIsNumber = prefixIsNumber,
                    PrefixIsLowercase = prefixIsLowercase,
                    PrefixIsUppercase = prefixIsUppercase,
                    PrefixIsDash = prefixIsDash,
                    SuffixIsEndSign = suffixIsEndSign,
                    SuffixIsColon = suffixIsColon,
                    SuffixIsCommaOrSemicolon = suffixIsCommaOrSemicolon,
                    ContainsDash = containsDash,
                    ContainsBracket = containsBracket,
                    FontSize = fontSize,
                    LineSpacing = lineSpacing,
                    LineSpacingRule = (int)lineSpacingRule,
                    Italic = (int)italic,
                    Bold = (int)bold,
                    BlackColor = blackColor

                };
                allNormalizedProperties.Add(normalizedParagraphProperties);
                iteration++;
            }

            CloseApp();

            return allNormalizedProperties;
        }

        // Печать всех абзацев
        public override void PrintAllParagraphs()
        {
            try
            {
                OpenApp();
                OpenDocument();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                CloseApp();
                return;
            }

            foreach (Word.Paragraph paragraph in Document.Paragraphs)
            {
                Console.WriteLine(paragraph.Range.Text);
            }

            CloseApp();
        }
    }
}