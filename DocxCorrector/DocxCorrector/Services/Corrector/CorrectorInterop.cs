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

        // Получить результат для параграфа открытого документа под номером paragraphNum, рассматривая его как тип type
        private ParagraphResult? GetResultForParagraph(ElementType type, int paragraphNum)
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

            // Первые 20 символов параграфа
            string prefixOfParagraph = InteropHelper.GetParagraphPrefix(paragraph: paragraph, prefixLength: 20);

            ParagraphResult result = new ParagraphResult
            {
                ParagraphID = paragraphNum,
                Type = type,
                Prefix = prefixOfParagraph,
                Mistakes = new List<Mistake>()
            };

            result.Mistakes.AddRange(GetGeneralMistakesForParagraph(paragraph: paragraph, type: type));
            // Тут можно вызвать отдельную функцию для особенных элементов (например, проверяющих следующий и предыдущий абзац для списка)

            if (result.Mistakes.Count == 0)
            {
                return null;
            }

            return result;
        }

        // Получить список основных ошибок для параграфа paragraph, рассматривая его как тип type
        private List<Mistake> GetGeneralMistakesForParagraph(Word.Paragraph paragraph, ElementType type)
        {
            List<Mistake> result = new List<Mistake>();
            // Проверка общих свойств для всех типов
            // Отступ сверху
            if (paragraph.SpaceBefore != 0)
            {
                Mistake mistake = new Mistake(message: "Неверный отступ сверху (должен быть 0)");
                result.Add(mistake);
            }

            // Отступ снизу
            if (paragraph.SpaceAfter != 0)
            {
                Mistake mistake = new Mistake(message: "Неверный отступ снизу (должен быть 0)");
                result.Add(mistake);
            }

            // Междустрочный интервал
            if (paragraph.LineSpacingRule != Word.WdLineSpacing.wdLineSpace1pt5)
            {
                Mistake mistake = new Mistake(message: "Неверный междустрочный интервал (должен быть 1.5)");
                result.Add(mistake);
            }

            // Название шрифта
            if (paragraph.Range.Font.Name.ToString() != "Times New Roman")
            {
                Mistake mistake = new Mistake(message: "Неверный шрифт (должен быть Times New Roman)");
                result.Add(mistake);
            }

            // Размер шрифта
            if (paragraph.Range.Font.Size != 14.0)
            {
                Mistake mistake = new Mistake(message: "Неверный размер шрифта (должен быть 14)");
                result.Add(mistake);
            }

            // Отступ первой строки
            if (paragraph.FirstLineIndent != (float)35.45)
            {
                Mistake mistake = new Mistake(message: "Неверный отступ первой строки (должен быть 1.25 см)");
                result.Add(mistake);
            }

            // Курсив
            if (paragraph.Range.Italic != 0)
            {
                Mistake mistake = new Mistake(message: "Параграф не может быть оформлен курсивом");
                result.Add(mistake);
            }

            // Жирный
            if (paragraph.Range.Bold != 0)
            {
                Mistake mistake = new Mistake(message: "Параграф не может быть оформлен жирным");
                result.Add(mistake);
            }

            // Подчеркнутость
            if (paragraph.Range.Underline != 0)
            {
                Mistake mistake = new Mistake(message: "Параграф не может быть подчернутым");
                result.Add(mistake);
            }

            // Зачеркнутость
            if (paragraph.Range.Font.StrikeThrough != 0 | paragraph.Range.Font.DoubleStrikeThrough != 0)
            {
                Mistake mistake = new Mistake(message: "Параграф не может быть зачернутым");
                result.Add(mistake);
            }

            // Цвет текста
            if (paragraph.Range.Font.TextColor.RGB != -16777216 & paragraph.Range.Font.TextColor.RGB != -587137025 & paragraph.Range.Font.TextColor.RGB != 0)
            {
                Mistake mistake = new Mistake(message: "Неверный цвет шрифта (должен быть черный)");
                result.Add(mistake);
            }

            // Проверка свойств спецефичных для каждого из типов
            switch (type)
            {
                // АБЗАЦ
                case ElementType.Paragraph:
                    // Положение на странице
                    if (paragraph.Alignment != Word.WdParagraphAlignment.wdAlignParagraphJustify)
                    {
                        Mistake mistake = new Mistake(message: "Неверное положение на странице (должно быть по ширине)");
                        result.Add(mistake);
                    }

                    // Заглавная буква в начале
                    if (!Char.IsUpper(paragraph.Range.Text[0]))
                    {
                        Mistake mistake = new Mistake(message: "Элемент должен начинаться с большой буквы");
                        result.Add(mistake);
                    }

                    // Символ окончания
                    if (!Convert.ToBoolean(InteropHelper.CheckIfLastSymbolOfParagraphIs(paragraph: paragraph, new string[] { ".", ":", "!", "?" })))
                    {
                        Mistake mistake = new Mistake(message: "Неверный последний символ");
                        result.Add(mistake);
                    }

                    break;

                // ЭЛЕМЕНТ СПИСКА
                case ElementType.List:
                    // Положение на странице
                    if (paragraph.Alignment != Word.WdParagraphAlignment.wdAlignParagraphJustify & paragraph.Alignment != Word.WdParagraphAlignment.wdAlignParagraphLeft)
                    {
                        Mistake mistake = new Mistake(message: "Неверное положение на странице (должно быть по ширине или слева)");
                        result.Add(mistake);
                    }

                    // Первый символ - черта или число
                    if (Convert.ToBoolean(InteropHelper.CheckIfFirstSymbolOfParagraphIs(paragraph: paragraph, new string[] { "-", "־", "᠆", "‐", "‑", "‒", "–", "—", "―", "﹘", "﹣", "－" })) &
                        !Char.IsNumber(paragraph.Range.Text[0]))
                    {
                        Mistake mistake = new Mistake(message: "Неверный первый символ");
                        result.Add(mistake);
                    }

                    // Начало со строчной буквы
                    if (paragraph.Range.Words.Count > 2)
                    {
                        if (Char.IsLower(paragraph.Range.Words[1].Text[0]))
                        {
                            Mistake mistake = new Mistake(message: "Пункт должен начинаться со строчной буквы");
                            result.Add(mistake);
                        }
                    }
                   
                    // Символ окончания
                    if (!Convert.ToBoolean(InteropHelper.CheckIfLastSymbolOfParagraphIs(paragraph: paragraph, new string[] { ".", ",", ";" })))
                    {
                        Mistake mistake = new Mistake(message: "Неверный последний символ");
                        result.Add(mistake);
                    }
                    
                    break;

                // ПОДПИСЬ К РИСУНКУ
                case ElementType.ImageSign:
                    // Положение на странице
                    if (paragraph.Alignment != Word.WdParagraphAlignment.wdAlignParagraphCenter)
                    {
                        Mistake mistake = new Mistake(message: "Неверное положение на странице (должно быть по центру)");
                        result.Add(mistake);
                    }

                    // Первое слово
                    // TODO: - Проверить регуляркой
                    if (paragraph.Range.Words.Count > 2)
                    {
                        if (paragraph.Range.Words[1].Text != "Рисунок")
                        {
                            Mistake mistake = new Mistake(message: "Подпись к рисунку должна начинаться со слова Рисунок");
                            result.Add(mistake);
                        }
                    }

                    // Символ окончания
                    if (paragraph.Range.Text.Length > 1)
                    {
                        if (!Char.IsLetter(paragraph.Range.Text[paragraph.Range.Text.Length - 2]))
                        {
                            Mistake mistake = new Mistake(message: "Подпись к рисунку не должна заканичиваться знаком препинания");
                            result.Add(mistake);
                        }

                    }

                    break;
                default:
                    throw new NotImplementedException();
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
                ParagraphProperties paragraphProperties = new ParagraphPropertiesInterop(paragraph);
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
                PageProperties currentPageProperties = new PagePropertiesInterop(pageSetup: pageRange.PageSetup, pageNumber: pageNumber);
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
                NormalizedProperties normalizedParagraphProperties = new NormalizedPropertiesInterop(paragraph: paragraph, paragraphId: iteration);
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

                currentParagraphResult = GetResultForParagraph(type: elementType, paragraphNum: iteration);

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