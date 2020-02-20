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

        private void PrintPropertiesOfParagraph(Word.Paragraph paragraph)
        {
            Console.WriteLine($"Уровень заголовка: {paragraph.OutlineLevel}");
            Console.WriteLine($"Выравнивание: {paragraph.Alignment}");
            Console.WriteLine($"Отступ слева (в знаках): {paragraph.CharacterUnitLeftIndent}");
            Console.WriteLine($"Отступ слева (в пунктах): {paragraph.LeftIndent}");
            Console.WriteLine($"Отступ справа (в знаках): {paragraph.CharacterUnitRightIndent}");
            Console.WriteLine($"Отступ справа (в пунктах): {paragraph.RightIndent}");
            Console.WriteLine($"Отступ первой строки: {paragraph.CharacterUnitFirstLineIndent}");
            Console.WriteLine($"Зеркальность отступов: {paragraph.MirrorIndents}");
            Console.WriteLine($"Междустрочный интервал: {paragraph.LineSpacing}");
            Console.WriteLine($"Интервал перед: {paragraph.SpaceBefore}");
            Console.WriteLine($"Интервал после: {paragraph.SpaceAfter}");
            Console.WriteLine($"Интервал после: {paragraph.PageBreakBefore}");
        }

        private void PrintPropertiesOfRange(Word.Range range)
        {
            Console.WriteLine($"Текст: {range.Text}");
            Console.WriteLine($"Имя шрифта: {range.Font.Name}");
            Console.WriteLine($"Размер шрифта: {range.Font.Size}");
            Console.WriteLine($"Жирный: {range.Bold}");
            Console.WriteLine($"Курсив: {range.Italic}");
            Console.WriteLine($"Цвет текста: {range.Font.TextColor.RGB}");
            Console.WriteLine($"Цвет выделения: {range.Font.UnderlineColor}");
            Console.WriteLine($"Подчеркнутый: {range.Underline}");
            Console.WriteLine($"Зачеркнутый: {range.Font.StrikeThrough}");
            Console.WriteLine($"Надстрочность: {range.Font.Superscript}");
            Console.WriteLine($"Подстрочность: {range.Font.Subscript}");
            Console.WriteLine($"Скрытый: {range.Font.Hidden}");
            Console.WriteLine($"Масштаб: {range.Font.Scaling}");
            Console.WriteLine($"Смещение: {range.Font.Position}");
            Console.WriteLine($"Кернинг: {range.Font.Kerning}");
        }

        // Corrector
        public override string GetMistakesJSON()
        {
            List<ParagraphResult> paragraphResults = new List<ParagraphResult>();
            
            // TODO: - Remove
            ParagraphResult testResult = new ParagraphResult
            {
                ParagraphID = 0,
                Type = ElementType.Paragraph,
                Suffix = "TestParagraph",
                Mistakes = new List<Mistake> { new Mistake { Message = "Not Implemented" } }
            };
            paragraphResults.Add(testResult);

            // TODO: - Implement method

            string mistakesJSON = JSONMaker.MakeMistakesJSON(paragraphResults);
            return mistakesJSON;
        }

        public override void PrintAllParagraphs()
        {
            OpenWord();

            foreach (Word.Paragraph paragraph in Document.Paragraphs)
            {
                Console.WriteLine(paragraph.Range.Text);
            }

            QuitWord();
        }

        public override void PrintFirstParagraphProperties()
        {
            OpenWord();

            Word.Paragraph paragraph = Document.Paragraphs.First;

            PrintPropertiesOfParagraph(paragraph);
            PrintPropertiesOfRange(paragraph.Range);

            QuitWord();
        }

        public override void PrintFirstTwoWordsProperties()
        {
            throw new NotImplementedException();
        }
    }
}