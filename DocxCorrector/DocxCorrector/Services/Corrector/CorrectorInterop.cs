#nullable enable
using System;
using System.Collections.Generic;
using DocxCorrector.Models;
using DocxCorrector.Services;
using Word = Microsoft.Office.Interop.Word;

namespace DocxCorrector.Services.Corrector
{
    class CorrectorInteropExeption : Exception
    {
        public CorrectorInteropExeption(string message) : base(message) { }
    }

    public sealed class CorrectorInterop : Corrector
    {
        // Private
        private Word.Application? App;
        private Word.Document? Document;

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
                Document = App!.Documents.Open(FileName: FilePath, Visible: true);
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
#endif
                throw new CorrectorInteropExeption(message: "Can't open document");
            }
        }

        // Закрыть документ
        private void CloseDocument()
        {
            try
            {
                if (Document != null) { App!.Documents.Close(); }
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

        // Проверить параграф под номером paragraphNum в выбранном документе на наличие ошибок АБЗАЦА
        private ParagraphResult? GetMistakesSimpleParagraph(int paragraphNum)
        {
            if (Document == null) { return null; }

            Word.Paragraph paragraph;
            try
            {
                paragraph = Document.Paragraphs[paragraphNum];
            } 
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
#endif
                return null;
            }

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
                Mistake mistake = new Mistake(message: "Wrong paragraph alignment (must be justify)");
                result.Mistakes.Add(mistake);
            }

            //Проверка отступа сверху
            if (paragraph.SpaceBefore != 0 )
            {
                Mistake mistake = new Mistake(message: "Space before the paragraph must be 0");
                result.Mistakes.Add(mistake);
            }

            //Проверка отступа снизу
            if (paragraph.SpaceAfter != 0)
            {
                Mistake mistake = new Mistake(message: "Space after the paragraph must be 0");
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия заглавной буквы в начале абзаца
            if (prefixOfParagraph[0] != prefixOfParagraph.ToUpper()[0])
            {
                Mistake mistake = new Mistake(message: "Starting capital letter missed");
                result.Mistakes.Add(mistake);
            }

            //Проверка 1.5-ого межстрочного интервала
            if (paragraph.LineSpacingRule != Word.WdLineSpacing.wdLineSpace1pt5)
            {
                Mistake mistake = new Mistake(message: "Wrong line spacing (not 1.5 lines)");
                result.Mistakes.Add(mistake);
            }

            //Проверка размера шрифта
            if (paragraph.Range.Font.Size != 14.0)
            {
                Mistake mistake = new Mistake(message: "Wrong font size (must be 14 pt)");
                result.Mistakes.Add(mistake);
            }

            //Проверка типа шрифта
            if (paragraph.Range.Font.Name.ToString() != "Times New Roman")
            {
                Mistake mistake = new Mistake(message: "Wrong font type (must be Times New Roman)");
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия красной строки
            if (paragraph.FirstLineIndent != (float)35.45)
            {
                Mistake mistake = new Mistake(message: "There is no first line indent in this paragraph");
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия лишних стилей (лучше FontStyle)
            if (paragraph.Range.Italic != 0 & paragraph.Range.Bold != 0 & paragraph.Range.Underline != 0)
            {
                Mistake mistake = new Mistake(message: "Font style of paragraph must be regular");
                result.Mistakes.Add(mistake);
            }

            //Проверка цвета текста
            if (paragraph.Range.Font.TextColor.RGB != -16777216 & paragraph.Range.Font.TextColor.RGB != -587137025 & paragraph.Range.Font.TextColor.RGB != 0)
            {
                Mistake mistake = new Mistake(message: "Font color must be neutral (black)");
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
                    Mistake mistake = new Mistake(message: "Wrong ending of paragraph (proper punctuation mark expected)");
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
        private ParagraphResult? GetMistakesListElement(int paragraphNum)
        {
            if (Document == null) { return null; }

            Word.Paragraph paragraph;
            try
            {
                paragraph = Document.Paragraphs[paragraphNum];
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
#endif
                return null;
            }

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
                Mistake mistake = new Mistake(message: "Wrong list element alignment (must be justify or left)");
                result.Mistakes.Add(mistake);
            }

            //Проверка отступа сверху
            if (paragraph.SpaceBefore != 0)
            {
                Mistake mistake = new Mistake(message: "Space before the paragraph must be 0");
                result.Mistakes.Add(mistake);
            }

            //Проверка отступа снизу
            if (paragraph.SpaceAfter != 0)
            {
                Mistake mistake = new Mistake(message: "Space after the paragraph must be 0");
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия заглавной буквы в начале абзаца
            if (prefixOfParagraph[2] != prefixOfParagraph.ToLower()[2])
            {
                Mistake mistake = new Mistake(message: "Wrong starting letter (capital letters are not allowed)");
                result.Mistakes.Add(mistake);
            }

            //Проверка 1.5-ого межстрочного интервала
            if (paragraph.LineSpacingRule != Word.WdLineSpacing.wdLineSpace1pt5)
            {
                Mistake mistake = new Mistake(message: "Wrong line spacing (not 1.5 lines)");
                result.Mistakes.Add(mistake);
            }

            //Проверка размера шрифта
            if (paragraph.Range.Font.Size != 14.0)
            {
                Mistake mistake = new Mistake(message: "Wrong font size (must be 14 pt)");
                result.Mistakes.Add(mistake);
            }

            //Проверка типа шрифта
            if (paragraph.Range.Font.Name.ToString() != "Times New Roman")
            {
                Mistake mistake = new Mistake(message: "Wrong font type (must be Times New Roman)");
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия красной строки
            if (paragraph.FirstLineIndent == (float)35.45)
            {
                Mistake mistake = new Mistake(message: "There is no first line indent needed in list element");
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия лишних стилей (лучше FontStyle)
            if (paragraph.Range.Italic != 0 & paragraph.Range.Bold != 0 & paragraph.Range.Underline != 0)
            {
                Mistake mistake = new Mistake(message: "Font style of list element must be regular");
                result.Mistakes.Add(mistake);
            }

            //Проверка цвета текста
            if (paragraph.Range.Font.TextColor.RGB != -16777216 & paragraph.Range.Font.TextColor.RGB != -587137025 & paragraph.Range.Font.TextColor.RGB != 0)
            {
                Mistake mistake = new Mistake(message: "Font color must be neutral (black)");
                result.Mistakes.Add(mistake);
            }

            //Проверка конечных знаков препинания
            if (paragraph.Range.Text.Length > 1)
            {
                if (paragraph.Range.Text[paragraph.Range.Text.Length - 2].ToString() != ";" & 
                    paragraph.Range.Text[paragraph.Range.Text.Length - 2].ToString() != ".")
                {
                    Mistake mistake = new Mistake(message: "Wrong ending of list element (proper punctuation mark expected)");
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
                    Mistake mistake = new Mistake(message: "Wrong starting of list element (proper punctuation mark expected)");
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
        private ParagraphResult? GetMistakesImageSign(int paragraphNum)
        {
            if (Document == null) { return null; }

            Word.Paragraph paragraph;
            try
            {
                paragraph = Document.Paragraphs[paragraphNum];
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
#endif
                return null;
            }

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

            //Условия возникновения ошибок в ПОДПИСИ К РИСУНКУ

            //Проверка центровки подписи к картинке по центру
            if (paragraph.Alignment != Word.WdParagraphAlignment.wdAlignParagraphCenter)
            {
                Mistake mistake = new Mistake(message: "Wrong image sign alignment (must be center)");
                result.Mistakes.Add(mistake);
            }

            //Проверка отступа сверху
            if (paragraph.SpaceBefore != 0)
            {
                Mistake mistake = new Mistake(message: "Space before image sign must be 0");
                result.Mistakes.Add(mistake);
            }

            //Проверка отступа снизу
            if (paragraph.SpaceAfter != 0)
            {
                Mistake mistake = new Mistake(message: "Space after image sign must be 0");
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия заглавной буквы в начале абзаца
            if (prefixOfParagraph[0] != prefixOfParagraph.ToUpper()[0])
            {
                Mistake mistake = new Mistake(message: "Wrong starting letter (must be capital)");
                result.Mistakes.Add(mistake);
            }

            //Проверка 1.5-ого межстрочного интервала
            if (paragraph.LineSpacingRule != Word.WdLineSpacing.wdLineSpace1pt5)
            {
                Mistake mistake = new Mistake(message: "Wrong line spacing (not 1.5 lines)");
                result.Mistakes.Add(mistake);
            }

            //Проверка размера шрифта
            if (paragraph.Range.Font.Size != 14.0)
            {
                Mistake mistake = new Mistake(message: "Wrong font size (must be 14 pt)");
                result.Mistakes.Add(mistake);
            }

            //Проверка типа шрифта
            if (paragraph.Range.Font.Name.ToString() != "Times New Roman")
            {
                Mistake mistake = new Mistake(message: "Wrong font type (must be Times New Roman)");
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия красной строки
            if (paragraph.FirstLineIndent == (float)35.45)
            {
                Mistake mistake = new Mistake(message: "There is no first line indent needed in image sign");
                result.Mistakes.Add(mistake);
            }

            //Проверка наличия лишних стилей (лучше FontStyle)
            if (paragraph.Range.Italic != 0 & paragraph.Range.Bold != 0 & paragraph.Range.Underline != 0)
            {
                Mistake mistake = new Mistake(message: "Font style of image sign must be regular");
                result.Mistakes.Add(mistake);
            }

            //Проверка цвета текста
            if (paragraph.Range.Font.TextColor.RGB != -16777216 & paragraph.Range.Font.TextColor.RGB != -587137025 & paragraph.Range.Font.TextColor.RGB != 0)
            {
                Mistake mistake = new Mistake(message: "Font color must be neutral (black)");
                result.Mistakes.Add(mistake);
            }

            //Проверка конечных знаков препинания
            if (paragraph.Range.Text.Length > 1)
            {
                if (Char.IsLetter(paragraph.Range.Text[paragraph.Range.Text.Length - 2]))
                {
                    Mistake mistake = new Mistake(message: "Wrong ending of image sign (No exclamation mark is required)");
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
        public CorrectorInterop(string? filePath = null) : base(filePath) { }

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
                Mistakes = new List<Mistake> { new Mistake(message: "Русские буквы") }
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

            foreach (Word.Paragraph paragraph in Document!.Paragraphs)
            {
                ParagraphProperties paragraphProperties = new ParagraphProperties(paragraph);
                allParagraphsProperties.Add(paragraphProperties);
            }

            CloseApp();

            return allParagraphsProperties;
        }
        
        //Получить свойства всех страниц
        public override List<PageProperties> GetAllPagesProperties()
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
                return new List<PageProperties>();
            }

            List<PageProperties> result = new List<PageProperties>();

            int totalPageNumber = Document!.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
            for (int pageNumber = 1; pageNumber <= totalPageNumber; pageNumber++)
            {
                Word.Range pageRange = Document.Range().GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToAbsolute, pageNumber);
                PageProperties currentPageProperties = new PageProperties(pageSetup: pageRange.PageSetup, pageNumber: pageNumber);
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
            foreach (Word.Paragraph paragraph in Document!.Paragraphs)
            {
                NormalizedProperties normalizedParagraphProperties = new NormalizedProperties(paragraph: paragraph, paragraphId: iteration);
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

            foreach (Word.Paragraph paragraph in Document!.Paragraphs)
            {
                Console.WriteLine(paragraph.Range.Text);
            }

            CloseApp();
        }

        // Получить списк ошибок для выбранного документа, с учетом того, что все параграфы в нем типа elementType
        public override List<ParagraphResult> GetMistakesForElementType(ElementType elementType)
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

            int iteration = 1;
            foreach (Word.Paragraph paragraph in Document!.Paragraphs)
            {
                ParagraphResult? currentParagraphResult;

                currentParagraphResult = elementType switch
                {
                    ElementType.Paragraph => GetMistakesSimpleParagraph(paragraphNum: iteration),
                    ElementType.List => GetMistakesListElement(paragraphNum: iteration),
                    ElementType.ImageSign => currentParagraphResult = GetMistakesImageSign(paragraphNum: iteration),
                    _ => null
                };


                // Если ошибки параграфа найдены добавить их в общий список
                if (currentParagraphResult != null)
                {
                    paragraphResults.Add(currentParagraphResult);
                }

                iteration++;
            }

            CloseApp();

            return paragraphResults;
        }
    }
}