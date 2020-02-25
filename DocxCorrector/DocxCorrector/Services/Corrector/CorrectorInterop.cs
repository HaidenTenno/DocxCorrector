using System;
using System.Collections.Generic;
using DocxCorrector.Models;
using DocxCorrector.Services;
using Word = Microsoft.Office.Interop.Word;

namespace DocxCorrector.Services.Corrector
{
    class InteropExeption : Exception
    {
        public InteropExeption(string message) : base(message) { }
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
                Document = App.Documents.Open(FileName: FilePath, Visible: false);
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

        // Получить свойства параграфа
        private ParagraphProperties GetParagraphProperties(Word.Paragraph paragraph)
        {
            var Text = paragraph.Range.Text.ToString();
            var Underline = paragraph.Range.Underline.ToString();
            var Bold = paragraph.Range.Bold.ToString();
            var Italic = paragraph.Range.Italic.ToString();
            var BoldBi = paragraph.Range.BoldBi.ToString();
            var Bookmarks = paragraph.Range.Bookmarks.ToString();
            var Borders = paragraph.Range.Borders.ToString();
            var Case = paragraph.Range.Case.ToString();
            var Characters = paragraph.Range.Characters.ToString();
            var CharacterWidth = paragraph.Range.CharacterWidth.ToString();
            var CombineCharacters = paragraph.Range.CombineCharacters.ToString();
            var ContentControls = paragraph.Range.ContentControls.ToString();
            var Creator = paragraph.Range.Creator.ToString();
            var DisableCharacterSpaceGrid = paragraph.Range.DisableCharacterSpaceGrid.ToString();
            var Document = paragraph.Range.Document.ToString();
            var Duplicate = paragraph.Range.Duplicate.ToString();
            var Editors = paragraph.Range.Editors.ToString();
            var EmpasisMark = paragraph.Range.EmphasisMark.ToString();
            var End = paragraph.Range.End.ToString();
            var EndnoteOptions = paragraph.Range.EndnoteOptions.ToString();
            var Endnotes = paragraph.Range.Endnotes.ToString();
            var Fields = paragraph.Range.Fields.ToString();
            var Find = paragraph.Range.Find.ToString();
            var FitTextWidth = paragraph.Range.FitTextWidth.ToString();
            var Footnotes = paragraph.Range.Footnotes.ToString();
            var FormattedText = paragraph.Range.FormattedText.ToString();
            var FormFields = paragraph.Range.FormFields.ToString();
            var Frames = paragraph.Range.Frames.ToString();
            var GrammarChecked = paragraph.Range.GrammarChecked.ToString();
            var GrammaticalErrors = paragraph.Range.GrammaticalErrors.ToString();
            var HighlightColorIndex = paragraph.Range.HighlightColorIndex.ToString();
            var HorizontallnVertical = paragraph.Range.HorizontalInVertical.ToString();
            var HTMLDicisions = paragraph.Range.HTMLDivisions.ToString();
            var Hyperlinks = paragraph.Range.Hyperlinks.ToString();
            var InlineShapes = paragraph.Range.InlineShapes.ToString();
            var IsEndOfMark = paragraph.Range.IsEndOfRowMark.ToString();
            var ItalicBi = paragraph.Range.ItalicBi.ToString();
            var Kana = paragraph.Range.Kana.ToString();
            var LanguageDetected = paragraph.Range.LanguageDetected.ToString();
            var LanguageID = paragraph.Range.LanguageID.ToString();
            var LanguageIDFarEst = paragraph.Range.LanguageIDFarEast.ToString();
            var LanguageIDOther = paragraph.Range.LanguageIDOther.ToString();
            var ListFormat = paragraph.Range.ListFormat.ToString();
            var ListParagraphs = paragraph.Range.ListParagraphs.ToString();
            var NoProofing = paragraph.Range.NoProofing.ToString();
            var OMaths = paragraph.Range.OMaths.ToString();
            var Orientation = paragraph.Range.Orientation.ToString();
            var PageSetup = paragraph.Range.PageSetup.ToString();
            var ParagraphFormat = paragraph.Range.ParagraphFormat.ToString();
            var Paragraphs = paragraph.Range.Paragraphs.ToString();
            var PreviousBookmarkID = paragraph.Range.PreviousBookmarkID.ToString();
            var ReadabilityStatistics = paragraph.Range.ReadabilityStatistics.ToString();
            var Revisions = paragraph.Range.Revisions.ToString();
            var Sections = paragraph.Range.Sections.ToString();
            var Sentenses = paragraph.Range.Sentences.ToString();
            var Shading = paragraph.Range.Shading.ToString();
            var ShapeRange = paragraph.Range.ShapeRange.ToString();
            var ShowAll = paragraph.Range.ShowAll.ToString();
            var SmartTags = paragraph.Range.SmartTags.ToString();
            var SpellingChecked = paragraph.Range.SpellingChecked.ToString();
            var SpellingErrors = paragraph.Range.SpellingErrors.ToString();
            var Subdocuments = paragraph.Range.Subdocuments.ToString();
            var SynonymInfo = paragraph.Range.SynonymInfo.ToString();
            var Tables = paragraph.Range.Tables.ToString();
            var TextRetrievalMode = paragraph.Range.TextRetrievalMode.ToString();
            var TextVisibleOnScreen = paragraph.Range.TextVisibleOnScreen.ToString();
            var TopLevelTables = paragraph.Range.TopLevelTables.ToString();
            var TwoLinesInOne = paragraph.Range.TwoLinesInOne.ToString();
            var Words = paragraph.Range.Words.ToString();
            var FontName = paragraph.Range.Font.Name.ToString();
            var FontSize = paragraph.Range.Font.Size.ToString();
            //var FontTextColorRGB = paragraph.Range.Font.TextColor.RGB.ToString();
            var FontUnderlineColor = paragraph.Range.Font.UnderlineColor.ToString();
            var FontStrikeThrough = paragraph.Range.Font.StrikeThrough.ToString();
            var FontSuperscript = paragraph.Range.Font.Superscript.ToString();
            var FontSubscript = paragraph.Range.Font.Superscript.ToString();
            var FontHidden = paragraph.Range.Font.Hidden.ToString();
            var FontScaling = paragraph.Range.Font.Scaling.ToString();
            var FontPosition = paragraph.Range.Font.Position.ToString();
            var FontKerning = paragraph.Range.Font.Kerning.ToString();
            var FontApplication = paragraph.Range.Font.Application.ToString();
            var FontBoldBi = paragraph.Range.Font.BoldBi.ToString();
            var FontBorders = paragraph.Range.Font.Borders.ToString();
            var FontColor = paragraph.Range.Font.Color.ToString();
            var FontColorIndex = paragraph.Range.Font.ColorIndex.ToString();
            var FontColorIndexBi = paragraph.Range.Font.ColorIndexBi.ToString();
            var FontContextualAlternates = paragraph.Range.Font.ContextualAlternates.ToString();
            var FontCreator = paragraph.Range.Font.Creator.ToString();
            var FontDiacriricColor = paragraph.Range.Font.DiacriticColor.ToString();
            var FontDoubleStrikeThrough = paragraph.Range.Font.DoubleStrikeThrough.ToString();
            var FontDuplicate = paragraph.Range.Font.Duplicate.ToString();
            var FontEmboss = paragraph.Range.Font.Emboss.ToString();
            var FontEmphasisMark = paragraph.Range.Font.EmphasisMark.ToString();
            var FontEngrave = paragraph.Range.Font.Engrave.ToString();
            // Обратить внимание
            //var FontFill = paragraph.Range.Font.Fill.ForeColor.RGB.ToString();
            //var FontGlow = paragraph.Range.Font.Glow.ToString();
            var FontItalic = paragraph.Range.Font.Italic.ToString();
            var FontItalicBi = paragraph.Range.Font.ItalicBi.ToString();
            var FontLigatures = paragraph.Range.Font.Ligatures.ToString();
            var FontNameAscii = paragraph.Range.Font.NameAscii.ToString();
            var FontNameBi = paragraph.Range.Font.NameBi.ToString();
            var FontNameFarEast = paragraph.Range.Font.NameFarEast.ToString();
            var FontNameOther = paragraph.Range.Font.NameOther.ToString();
            var FontNumberForm = paragraph.Range.Font.NumberForm.ToString();
            var FontNumberSpacing = paragraph.Range.Font.NumberSpacing.ToString();
            var FontOutline = paragraph.Range.Font.Outline.ToString();
            var FontParent = paragraph.Range.Font.Parent.ToString();
            //var FontReflection = paragraph.Range.Font.Reflection.ToString();
            var FontShading = paragraph.Range.Font.Shading.ToString();
            var FontShadow = paragraph.Range.Font.Shadow.ToString();
            var FontSizeBi = paragraph.Range.Font.SizeBi.ToString();
            var FontSmallCaps = paragraph.Range.Font.SmallCaps.ToString();
            var FontStylisticSet = paragraph.Range.Font.StylisticSet.ToString();
            //var FontTextShadow = paragraph.Range.Font.TextShadow.ToString();
            //var FontThreeD = paragraph.Range.Font.ThreeD.ToString();
            var FontUnderline = paragraph.Range.Font.Underline.ToString();
            // Paragraph 
            var OutlineLevel = paragraph.OutlineLevel.ToString();
            var Alignment = paragraph.Alignment.ToString();
            var CharacterUnitLeftIndent = paragraph.CharacterUnitLeftIndent.ToString();
            var LeftIndent = paragraph.LeftIndent.ToString();
            var CharacterUnitRightIndent = paragraph.CharacterUnitLeftIndent.ToString();
            var RightIndent = paragraph.RightIndent.ToString();
            var CharacterUnitFirstLineIndent = paragraph.CharacterUnitFirstLineIndent.ToString();
            var MirrorIndents = paragraph.MirrorIndents.ToString();
            var LineSpacing = paragraph.LineSpacing.ToString();
            var SpaceBefore = paragraph.SpaceBefore.ToString();
            var SpaceAfter = paragraph.SpaceAfter.ToString();
            var PageBreakBefore = paragraph.PageBreakBefore.ToString();
            var AddSpaceBetweenFarEastAndAlpha = paragraph.AddSpaceBetweenFarEastAndAlpha.ToString();
            var AddSpaceBetweenFarEastAndDigit = paragraph.AddSpaceBetweenFarEastAndDigit.ToString();
            var Application = paragraph.Application.ToString();
            var AutoAdjustRightIndent = paragraph.AutoAdjustRightIndent.ToString();
            var BaseLineAlignment = paragraph.BaseLineAlignment.ToString();
            var ParagraphBorders = paragraph.Borders.ToString();
            var CollapsedState = paragraph.CollapsedState.ToString();
            var CollapseHEadingByDefault = paragraph.CollapseHeadingByDefault.ToString();
            var ParagraphCreator = paragraph.Creator.ToString();
            var DisableLineHeightGrid = paragraph.DisableLineHeightGrid.ToString();
            var DropCap = paragraph.DropCap.ToString();
            var FarEastLineBreakControl = paragraph.FarEastLineBreakControl.ToString();
            var FirstLineIndent = paragraph.FirstLineIndent.ToString();
            var HalfWidthPunctuationOnTopOfLine = paragraph.HalfWidthPunctuationOnTopOfLine.ToString();
            var HalfWidthPunctuation = paragraph.HalfWidthPunctuationOnTopOfLine.ToString();
            var Hyphenation = paragraph.Hyphenation.ToString();
            var IsStyleSeparator = paragraph.IsStyleSeparator.ToString();
            var KeepTogether = paragraph.KeepTogether.ToString();
            var KeepWithNext = paragraph.KeepWithNext.ToString();
            var LineSpacingRule = paragraph.LineSpacingRule.ToString();
            var LineUnitAfter = paragraph.LineUnitAfter.ToString();
            var LineUnitBefore = paragraph.LineUnitBefore.ToString();
            var NoLineNumber = paragraph.NoLineNumber.ToString();
            var ParagraphParent = paragraph.Parent.ToString();
            var ReadingOrder = paragraph.ReadingOrder.ToString();
            var ParagraphShading = paragraph.Shading.ToString();
            var SpaceAfterAuto = paragraph.SpaceAfter.ToString();
            var ParagraphStyle = paragraph.get_Style().ToString();
            var TabStops = paragraph.TabStops.ToString();
            var TextboxTightWrap = paragraph.TextboxTightWrap.ToString();
            var TextID = paragraph.TextID.ToString();
            var WindowControl = paragraph.WidowControl.ToString();
            var WordWrap = paragraph.WordWrap.ToString();

            ParagraphProperties paragraphProperties = new ParagraphProperties
            {
                // Range
                Text = Text,
                Underline = Underline,
                Bold = Bold,
                Italic = Italic,
                BoldBi = BoldBi,
                Bookmarks = Bookmarks,
                Borders = Borders,
                Case = Case,
                Characters = Characters,
                CharacterWidth = CharacterWidth,
                CombineCharacters = CombineCharacters,
                ContentControls = ContentControls,
                Creator = Creator,
                DisableCharacterSpaceGrid = DisableCharacterSpaceGrid,
                Document = Document,
                Duplicate = Duplicate,
                Editors = Editors,
                EmpasisMark = EmpasisMark,
                End = End,
                EndnoteOptions = EndnoteOptions,
                Endnotes = Endnotes,
                Fields = Fields,
                Find = Find,
                FitTextWidth = FitTextWidth,
                Footnotes = Footnotes,
                FormattedText = FormattedText,
                FormFields = FormFields,
                Frames = Frames,
                GrammarChecked = GrammarChecked,
                GrammaticalErrors = GrammaticalErrors,
                HighlightColorIndex = HighlightColorIndex,
                HorizontallnVertical = HorizontallnVertical,
                HTMLDicisions = HTMLDicisions,
                Hyperlinks = Hyperlinks,
                InlineShapes = InlineShapes,
                IsEndOfMark = IsEndOfMark,
                ItalicBi = ItalicBi,
                Kana = Kana,
                LanguageDetected = LanguageDetected,
                LanguageID = LanguageID,
                LanguageIDFarEst = LanguageIDFarEst,
                LanguageIDOther = LanguageIDOther,
                ListFormat = ListFormat,
                ListParagraphs = ListParagraphs,
                NoProofing = NoProofing,
                OMaths = OMaths,
                Orientation = Orientation,
                PageSetup = PageSetup,
                ParagraphFormat = ParagraphFormat,
                Paragraphs = Paragraphs,
                PreviousBookmarkID = PreviousBookmarkID,
                ReadabilityStatistics = ReadabilityStatistics,
                Revisions = Revisions,
                Sections = Sections,
                Sentenses = Sentenses,
                Shading = Shading,
                ShapeRange = ShapeRange,
                ShowAll = ShowAll,
                SmartTags = SmartTags,
                SpellingChecked = SpellingChecked,
                SpellingErrors = SpellingErrors,
                Subdocuments = Subdocuments,
                SynonymInfo = SynonymInfo,
                Tables = Tables,
                TextRetrievalMode = TextRetrievalMode,
                TextVisibleOnScreen = TextVisibleOnScreen,
                TopLevelTables = TopLevelTables,
                TwoLinesInOne = TwoLinesInOne,
                //WordOpenXML = WordOpenXML,
                Words = Words,
                //XML = XML,
                //XMLNodes = XMLNodes,
                // Font
                FontName = FontName,
                FontSize = FontSize,
                //FontTextColorRGB = FontTextColorRGB,
                FontUnderlineColor = FontUnderlineColor,
                FontStrikeThrough = FontStrikeThrough,
                FontSuperscript = FontSuperscript,
                FontSubscript = FontSuperscript,
                FontHidden = FontHidden,
                FontScaling = FontScaling,
                FontPosition = FontPosition,
                FontKerning = FontKerning,
                FontApplication = FontApplication,
                FontBoldBi = FontBoldBi,
                FontBorders = FontBorders,
                FontColor = FontColor,
                FontColorIndex = FontColorIndex,
                FontColorIndexBi = FontColorIndexBi,
                FontContextualAlternates = FontContextualAlternates,
                FontCreator = FontCreator,
                FontDiacriricColor = FontDiacriricColor,
                FontDoubleStrikeThrough = FontDoubleStrikeThrough,
                FontDuplicate = FontDuplicate,
                FontEmboss = FontEmboss,
                FontEmphasisMark = FontEmphasisMark,
                FontEngrave = FontEngrave,
                // Обратить внимание
                //FontFill = FontFill,
                //FontGlow = FontGlow,
                FontItalic = FontItalic,
                FontItalicBi = FontItalicBi,
                FontLigatures = FontLigatures,
                FontNameAscii = FontNameAscii,
                FontNameBi = FontNameBi,
                FontNameFarEast = FontNameFarEast,
                FontNameOther = FontNameOther,
                FontNumberForm = FontNumberForm,
                FontNumberSpacing = FontNumberSpacing,
                FontOutline = FontOutline,
                FontParent = FontParent,
                //FontReflection = FontReflection,
                FontShading = FontShading,
                FontShadow = FontShadow,
                FontSizeBi = FontSizeBi,
                FontSmallCaps = FontSmallCaps,
                FontStylisticSet = FontStylisticSet,
                //FontTextShadow = FontTextShadow,
                //FontThreeD = FontThreeD,
                FontUnderline = FontUnderline,
                // Paragraph 
                OutlineLevel = OutlineLevel,
                Alignment = Alignment,
                CharacterUnitLeftIndent = CharacterUnitLeftIndent,
                LeftIndent = LeftIndent,
                CharacterUnitRightIndent = CharacterUnitLeftIndent,
                RightIndent = RightIndent,
                CharacterUnitFirstLineIndent = CharacterUnitFirstLineIndent,
                MirrorIndents = MirrorIndents,
                LineSpacing = LineSpacing,
                SpaceBefore = SpaceBefore,
                SpaceAfter = SpaceAfter,
                PageBreakBefore = PageBreakBefore,
                AddSpaceBetweenFarEastAndAlpha = AddSpaceBetweenFarEastAndAlpha,
                AddSpaceBetweenFarEastAndDigit = AddSpaceBetweenFarEastAndDigit,
                Application = Application,
                AutoAdjustRightIndent = AutoAdjustRightIndent,
                BaseLineAlignment = BaseLineAlignment,
                ParagraphBorders = Borders,
                CollapsedState = CollapsedState,
                CollapseHEadingByDefault = CollapseHEadingByDefault,
                ParagraphCreator = Creator,
                DisableLineHeightGrid = DisableLineHeightGrid,
                DropCap = DropCap,
                FarEastLineBreakControl = FarEastLineBreakControl,
                FirstLineIndent = FirstLineIndent,
                HalfWidthPunctuationOnTopOfLine = HalfWidthPunctuationOnTopOfLine,
                HalfWidthPunctuation = HalfWidthPunctuationOnTopOfLine,
                Hyphenation = Hyphenation,
                IsStyleSeparator = IsStyleSeparator,
                KeepTogether = KeepTogether,
                KeepWithNext = KeepWithNext,
                LineSpacingRule = LineSpacingRule,
                LineUnitAfter = LineUnitAfter,
                LineUnitBefore = LineUnitBefore,
                NoLineNumber = NoLineNumber,
                ParagraphParent = ParagraphParent,
                ReadingOrder = ReadingOrder,
                ParagraphShading = Shading,
                SpaceAfterAuto = SpaceAfter,
                ParagraphStyle = ParagraphStyle,
                TabStops = TabStops,
                TextboxTightWrap = TextboxTightWrap,
                TextID = TextID,
                WindowControl = WindowControl,
                WordWrap = WordWrap
            };
            return paragraphProperties;
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
        public override string GetMistakesJSON()
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
                return "";
            }

            List<ParagraphResult> paragraphResults = new List<ParagraphResult>();

            // TODO: - Remove
            ParagraphResult testResult = new ParagraphResult
            {
                ParagraphID = 0,
                Type = ElementType.Paragraph,
                Prefix = "TestParagraph",
                Mistakes = new List<Mistake> { new Mistake(message: "Not Implemented") }
            };
            paragraphResults.Add(testResult);

            // TODO: - Implement method

            string mistakesJSON = JSONMaker.MakeMistakesJSON(paragraphResults);

            CloseApp();

            return mistakesJSON;
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