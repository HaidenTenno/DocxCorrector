using System;
using System.Collections.Generic;
using DocxCorrector.Models;
using DocxCorrector.Services;
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

        // Проверить параграф под номером paragraphNum в выбранном документе на наличие ошибок АБЗАЦА
        private ParagraphResult GetMistakesSimpleParagraph(int paragraphNum)
        {
            Word.Paragraph paragraph = Document.Paragraphs[paragraphNum];
            
            string prefixOfParagraph;

            if (Document.Paragraphs[paragraphNum].Range.Text.Length > 20)
            {
                prefixOfParagraph = Document.Paragraphs[paragraphNum].Range.Text.ToString().Substring(0, 20);
            }
            else
            {
                prefixOfParagraph = Document.Paragraphs[paragraphNum].Range.Text.ToString();
            }

            ParagraphResult result = new ParagraphResult
            {
                ParagraphID = paragraphNum,
                Type = ElementType.Paragraph,
                Prefix = prefixOfParagraph,
                Mistakes = new List<Mistake>()
            };

            //Условия возникновения ошибок в АБЗАЦЕ

            //Проверка центровки параграфа по ширине страницы
            if (paragraph.Alignment != Word.WdParagraphAlignment.wdAlignParagraphJustify)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Wrong paragraph alignment (must be justify)"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка отступа сверху
            if (paragraph.SpaceBefore != 0 )
            {
                Mistake mistake = new Mistake
                {
                    Message = "Space before the paragraph must be 0"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка отступа снизу
            if (paragraph.SpaceAfter != 0)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Space after the paragraph must be 0"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия заглавной буквы в начале абзаца
            if (prefixOfParagraph[0] != prefixOfParagraph.ToUpper()[0])
            {
                Mistake mistake = new Mistake
                {
                    Message = "Starting capital letter missed"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка 1.5-ого межстрочного интервала
            if (paragraph.LineSpacingRule != Word.WdLineSpacing.wdLineSpace1pt5)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Wrong line spacing (not 1.5 lines)"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка размера шрифта
            if (paragraph.Range.Font.Size != 14.0)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Wrong font size (must be 14 pt)"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка типа шрифта
            if (paragraph.Range.Font.Name.ToString() != "Times New Roman")
            {
                Mistake mistake = new Mistake
                {
                    Message = "Wrong font type (must be Times New Roman)"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия красной строки
            if (paragraph.FirstLineIndent != (float)35.45)
            {
                Mistake mistake = new Mistake
                {
                    Message = "There is no first line indent in this paragraph"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия лишних стилей (лучше FontStyle)
            if (paragraph.Range.Italic != 0 & paragraph.Range.Bold != 0 & paragraph.Range.Underline != 0)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Font style of paragraph must be regular"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка цвета текста
            if (paragraph.Range.Font.TextColor.RGB != -16777216 & paragraph.Range.Font.TextColor.RGB != -587137025 & paragraph.Range.Font.TextColor.RGB != 0)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Font color must be neutral (black)"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка конечных знаков препинания
            if (paragraph.Range.Text.Length > 1)
            {
                if (paragraph.Range.Text[paragraph.Range.Text.Length - 2].ToString() != "." &
                    paragraph.Range.Text[paragraph.Range.Text.Length - 2].ToString() != ":" &
                    paragraph.Range.Text[paragraph.Range.Text.Length - 2].ToString() != "!" &
                    paragraph.Range.Text[paragraph.Range.Text.Length - 2].ToString() != "?")
                {
                    Mistake mistake = new Mistake
                    {
                        Message = "Wrong ending of paragraph (proper punctuation mark expected)"
                    };
                    result.Mistakes.Add(mistake);
                }

            }

            if (result.Mistakes.Count == 0)
            { 
                return null;
            }

            return result;
        }

        // Проверить параграф под номером paragraphNum в выбранном документе на наличие ошибок ЭЛЕМЕНТА СПИСКА
        private ParagraphResult GetMistakesListElement(int paragraphNum)
        {
            Word.Paragraph paragraph = Document.Paragraphs[paragraphNum];

            string prefixOfParagraph;

            if (Document.Paragraphs[paragraphNum].Range.Text.Length > 5)
            {
                prefixOfParagraph = Document.Paragraphs[paragraphNum].Range.Text.ToString().Substring(0, 5);
            }
            else
            {
                prefixOfParagraph = Document.Paragraphs[paragraphNum].Range.Text.ToString();
            }

            ParagraphResult result = new ParagraphResult
            {
                ParagraphID = paragraphNum,
                Type = ElementType.Paragraph,
                Prefix = prefixOfParagraph,
                Mistakes = new List<Mistake>()
            };

            //Условия возникновения ошибок в ЭЛЕМЕНТЕ СПИСКА

            //Проверка центровки элемента списка по ширине страницы
            if (paragraph.Alignment != Word.WdParagraphAlignment.wdAlignParagraphJustify & paragraph.Alignment != Word.WdParagraphAlignment.wdAlignParagraphLeft)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Wrong list element alignment (must be justify or left)"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка отступа сверху
            if (paragraph.SpaceBefore != 0)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Space before the paragraph must be 0"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка отступа снизу
            if (paragraph.SpaceAfter != 0)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Space after the paragraph must be 0"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия заглавной буквы в начале абзаца
            if (prefixOfParagraph[2] != prefixOfParagraph.ToLower()[2])
            {
                Mistake mistake = new Mistake
                {
                    Message = "Wrong starting letter (capital letters are not allowed)"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка 1.5-ого межстрочного интервала
            if (paragraph.LineSpacingRule != Word.WdLineSpacing.wdLineSpace1pt5)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Wrong line spacing (not 1.5 lines)"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка размера шрифта
            if (paragraph.Range.Font.Size != 14.0)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Wrong font size (must be 14 pt)"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка типа шрифта
            if (paragraph.Range.Font.Name.ToString() != "Times New Roman")
            {
                Mistake mistake = new Mistake
                {
                    Message = "Wrong font type (must be Times New Roman)"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия красной строки
            if (paragraph.FirstLineIndent == (float)35.45)
            {
                Mistake mistake = new Mistake
                {
                    Message = "There is no first line indent needed in list element"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия лишних стилей (лучше FontStyle)
            if (paragraph.Range.Italic != 0 & paragraph.Range.Bold != 0 & paragraph.Range.Underline != 0)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Font style of list element must be regular"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка цвета текста
            if (paragraph.Range.Font.TextColor.RGB != -16777216 & paragraph.Range.Font.TextColor.RGB != -587137025 & paragraph.Range.Font.TextColor.RGB != 0)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Font color must be neutral (black)"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка конечных знаков препинания
            if (paragraph.Range.Text.Length > 1)
            {
                if (paragraph.Range.Text[paragraph.Range.Text.Length - 2].ToString() != ";" & 
                    paragraph.Range.Text[paragraph.Range.Text.Length - 2].ToString() != ".")
                {
                    Mistake mistake = new Mistake
                    {
                        Message = "Wrong ending of list element (proper punctuation mark expected)"
                    };
                    result.Mistakes.Add(mistake);
                }

            }

            //Проверка начальных знаков препинания
            if (paragraph.Range.Text.Length > 1)
            {
                if (paragraph.Range.Text[0].ToString() != "-" &
                    paragraph.Range.Text[0].ToString() != "־" &
                    paragraph.Range.Text[0].ToString() != "᠆" &
                    paragraph.Range.Text[0].ToString() != "‐" &
                    paragraph.Range.Text[0].ToString() != "‑" &
                    paragraph.Range.Text[0].ToString() != "‒" &
                    paragraph.Range.Text[0].ToString() != "–" &
                    paragraph.Range.Text[0].ToString() != "—" &
                    paragraph.Range.Text[0].ToString() != "―" &
                    paragraph.Range.Text[0].ToString() != "﹘" &
                    paragraph.Range.Text[0].ToString() != "﹣" &
                    paragraph.Range.Text[0].ToString() != "－" &
                    paragraph.Range.Text[0].ToString() != "1" &
                    paragraph.Range.Text[0].ToString() != "2" &
                    paragraph.Range.Text[0].ToString() != "3" &
                    paragraph.Range.Text[0].ToString() != "4" &
                    paragraph.Range.Text[0].ToString() != "5" &
                    paragraph.Range.Text[0].ToString() != "6" &
                    paragraph.Range.Text[0].ToString() != "7" &
                    paragraph.Range.Text[0].ToString() != "8" &
                    paragraph.Range.Text[0].ToString() != "9")
                    
                {
                    Mistake mistake = new Mistake
                    {
                        Message = "Wrong starting of list element (proper punctuation mark expected)"
                    };
                    result.Mistakes.Add(mistake);
                }

            }

            if (result.Mistakes.Count == 0)
            {
                return null;
            }

            return result;
        }

        // Проверить параграф под номером paragraphNum в выбранном документе на наличие ошибок ПОДПИСИ К РИСУНКУ
        private ParagraphResult GetMistakesImageSign(int paragraphNum)
        {
            Word.Paragraph paragraph = Document.Paragraphs[paragraphNum];

            string prefixOfParagraph;

            if (Document.Paragraphs[paragraphNum].Range.Text.Length > 10)
            {
                prefixOfParagraph = Document.Paragraphs[paragraphNum].Range.Text.ToString().Substring(0, 10);
            }
            else
            {
                prefixOfParagraph = Document.Paragraphs[paragraphNum].Range.Text.ToString();
            }

            ParagraphResult result = new ParagraphResult
            {
                ParagraphID = paragraphNum,
                Type = ElementType.Paragraph,
                Prefix = prefixOfParagraph,
                Mistakes = new List<Mistake>()
            };

            //Условия возникновения ошибок в ПОДПИСИ К КАРТИНКЕ

            //Проверка центровки подписи к картинке по центру
            if (paragraph.Alignment != Word.WdParagraphAlignment.wdAlignParagraphCenter)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Wrong image sign alignment (must be center)"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка отступа сверху
            if (paragraph.SpaceBefore != 0)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Space before image sign must be 0"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка отступа снизу
            if (paragraph.SpaceAfter != 0)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Space after image sign must be 0"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия заглавной буквы в начале абзаца
            if (prefixOfParagraph[0] != prefixOfParagraph.ToUpper()[0])
            {
                Mistake mistake = new Mistake
                {
                    Message = "Wrong starting letter (must be capital)"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка 1.5-ого межстрочного интервала
            if (paragraph.LineSpacingRule != Word.WdLineSpacing.wdLineSpace1pt5)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Wrong line spacing (not 1.5 lines)"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка размера шрифта
            if (paragraph.Range.Font.Size != 14.0)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Wrong font size (must be 14 pt)"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка типа шрифта
            if (paragraph.Range.Font.Name.ToString() != "Times New Roman")
            {
                Mistake mistake = new Mistake
                {
                    Message = "Wrong font type (must be Times New Roman)"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия красной строки
            if (paragraph.FirstLineIndent == (float)35.45)
            {
                Mistake mistake = new Mistake
                {
                    Message = "There is no first line indent needed in image sign"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия лишних стилей (лучше FontStyle)
            if (paragraph.Range.Italic != 0 & paragraph.Range.Bold != 0 & paragraph.Range.Underline != 0)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Font style of image sign must be regular"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка цвета текста
            if (paragraph.Range.Font.TextColor.RGB != -16777216 & paragraph.Range.Font.TextColor.RGB != -587137025 & paragraph.Range.Font.TextColor.RGB != 0)
            {
                Mistake mistake = new Mistake
                {
                    Message = "Font color must be neutral (black)"
                };
                result.Mistakes.Add(mistake);
            }

            //Проверка конечных знаков препинания
            if (paragraph.Range.Text.Length > 1)
            {
                if (Char.IsLetter(paragraph.Range.Text[paragraph.Range.Text.Length - 2]))
                {
                    Mistake mistake = new Mistake
                    {
                        Message = "Wrong ending of image sign (No exclamation mark is required)"
                    };
                    result.Mistakes.Add(mistake);
                }

            }

            if (result.Mistakes.Count == 0)
            {
                return null;
            }

            return result;
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
                Mistakes = new List<Mistake> { new Mistake { Message = "Not Implemented" } }
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

        // Получить JSON со списком ошибок для выбранного документа, с учетом того, что все параграфы в нем типа elementType
        public override string GetMistakesJSONForElementType(ElementType elementType)
        {
            string result = "";

            try
            {
                OpenApp();
                OpenDocument();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                CloseApp();
                return result;
            }

            List<ParagraphResult> paragraphResults = new List<ParagraphResult>();

            int iteration = 1;
            foreach (Word.Paragraph paragraph in Document.Paragraphs)
            {
                ParagraphResult currentParagraphResult = null;
                switch (elementType)
                {
                    case ElementType.Paragraph:
                        currentParagraphResult = GetMistakesSimpleParagraph(paragraphNum: iteration);
                        break;

                    case ElementType.List:
                        currentParagraphResult = GetMistakesListElement(paragraphNum: iteration);
                        break;

                    case ElementType.ImageSign:
                        currentParagraphResult = GetMistakesImageSign(paragraphNum: iteration);
                        break;

                    default:
                        break;
                }

                // Если ошибки параграфа найдены добавить их в общий список
                if (currentParagraphResult != null)
                {
                    paragraphResults.Add(currentParagraphResult);
                }

                iteration++;
            }

            string mistakesJSON = JSONMaker.MakeMistakesJSON(paragraphResults);

            CloseApp();

            return mistakesJSON;
        }
    }
}