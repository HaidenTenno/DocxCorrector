using System;
using System.Collections.Generic;
using DocxCorrector.Models;
using DocxCorrector.Services;
using Word = Microsoft.Office.Interop.Word;

namespace DocxCorrector.Services.Corrector
{
    public sealed class CorrectorInterop : Corrector
    {
        private Word.Application App;
        private Word.Document Document;

        public CorrectorInterop(string filePath) : base(filePath) { }

        // Private
        private void OpenWord()
        {
            App = new Word.Application();
            App.Visible = false;
            Document = App.Documents.Open(FilePath);
        }

        private void QuitWord()
        {
            App.Documents.Close();
            Document = null;
            App.Quit();
            App = null;
        }

        //private void PrintPropertiesOfParagraph(Word.Paragraph paragraph)
        //{
        //    Console.WriteLine($"Уровень заголовка: {paragraph.OutlineLevel}");
        //    Console.WriteLine($"Выравнивание: {paragraph.Alignment}");
        //    Console.WriteLine($"Отступ слева (в знаках): {paragraph.CharacterUnitLeftIndent}");
        //    Console.WriteLine($"Отступ слева (в пунктах): {paragraph.LeftIndent}");
        //    Console.WriteLine($"Отступ справа (в знаках): {paragraph.CharacterUnitRightIndent}");
        //    Console.WriteLine($"Отступ справа (в пунктах): {paragraph.RightIndent}");
        //    Console.WriteLine($"Отступ первой строки: {paragraph.CharacterUnitFirstLineIndent}");
        //    Console.WriteLine($"Зеркальность отступов: {paragraph.MirrorIndents}");
        //    Console.WriteLine($"Междустрочный интервал: {paragraph.LineSpacing}");
        //    Console.WriteLine($"Интервал перед: {paragraph.SpaceBefore}");
        //    Console.WriteLine($"Интервал после: {paragraph.SpaceAfter}");
        //    Console.WriteLine($"Интервал после: {paragraph.PageBreakBefore}");
        //}

        //private void PrintPropertiesOfRange(Word.Range range)
        //{
        //    Console.WriteLine($"Текст: {range.Text}");
        //    Console.WriteLine($"Имя шрифта: {range.Font.Name}");
        //    Console.WriteLine($"Размер шрифта: {range.Font.Size}");
        //    Console.WriteLine($"Жирный: {range.Bold}");
        //    Console.WriteLine($"Курсив: {range.Italic}");
        //    Console.WriteLine($"Цвет текста: {range.Font.TextColor.RGB}");
        //    Console.WriteLine($"Цвет подчеркивания: {range.Font.UnderlineColor}");
        //    Console.WriteLine($"Подчеркнутый: {range.Underline}");
        //    Console.WriteLine($"Зачеркнутый: {range.Font.StrikeThrough}");
        //    Console.WriteLine($"Надстрочность: {range.Font.Superscript}");
        //    Console.WriteLine($"Подстрочность: {range.Font.Subscript}");
        //    Console.WriteLine($"Скрытый: {range.Font.Hidden}");
        //    Console.WriteLine($"Масштаб: {range.Font.Scaling}");
        //    Console.WriteLine($"Смещение: {range.Font.Position}");
        //    Console.WriteLine($"Кернинг: {range.Font.Kerning}");
        //}

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

        // Corrector
        // Получение JSON-а со списком ошибок
        public override string GetMistakesJSON()
        {
            List<ParagraphResult> paragraphResults = new List<ParagraphResult>();
            
            // TODO: - Remove
            ParagraphResult testResult = new ParagraphResult
            {
                ParagraphID = 0,
                Type = ElementType.Paragraph,
                Prefix = "TestParagraph",
                Mistakes = new List<Mistake> { new Mistake { Message = "Not Implemented" } }
            };
            paragraphResults.Add(testResult);

            // TODO: - Implement method

            string mistakesJSON = JSONMaker.MakeMistakesJSON(paragraphResults);
            return mistakesJSON;
        }

        // Получить свойства всех параграфов
        public override List<ParagraphProperties> GetAllParagraphsProperties()
        {
            OpenWord();

            List<ParagraphProperties> allParagraphsProperties = new List<ParagraphProperties>();

            foreach (Word.Paragraph paragraph in Document.Paragraphs)
            {
                ParagraphProperties paragraphProperties = GetParagraphProperties(paragraph);
                allParagraphsProperties.Add(paragraphProperties);
            }

            QuitWord();

            return allParagraphsProperties;
        }

        // Получить нормализованные свойства параграфов (Для классификатора Ромы)
        public override List<NormalizedProperties> GetNormalizedProperties()
        {
            OpenWord();

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

                // TODO: - ВЫНЕСТИ ЭТО
                string[] dashes = new string[] {"-", "־", "᠆", "‐", "‑", "‒", "–", "—", "―", "﹘", "﹣", "－" };
                int prefixIsDash = Array.IndexOf(dashes, paragraph.Range.Text[0].ToString()) != -1 ? 1 : 0;
                string[] endSigns = new string[] { ".", "!", "?" };
                int suffixIsEndSign;
                if (paragraph.Range.Text.Length > 0) 
                {
                    suffixIsEndSign = Array.IndexOf(endSigns, paragraph.Range.Text[paragraph.Range.Text.Length - 2].ToString()) != -1 ? 1 : 0;
                } else
                {
                    suffixIsEndSign = 0;
                }

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
                    SuffixIsSemicolon = 123,
                    SuffixIsCommaOrSemicolon = 123,
                    ContainsDash = 123,
                    ContainsBracket = 123,
                    FontSize = 123,
                    LineSpacing = 123,
                    LineSpacingRule = 123,
                    Italic = 123,
                    Bold = 123,
                    BlackColor = 123

                };
                allNormalizedProperties.Add(normalizedParagraphProperties);
                iteration++;
            }

            QuitWord();

            return allNormalizedProperties;
        }

        // MARK: - Вспомогательные
        // Печать всех абзацев
        public override void PrintAllParagraphs()
        {
            OpenWord();

            foreach (Word.Paragraph paragraph in Document.Paragraphs)
            {
                Console.WriteLine(paragraph.Range.Text);
            }

            QuitWord();
        }
    }
}