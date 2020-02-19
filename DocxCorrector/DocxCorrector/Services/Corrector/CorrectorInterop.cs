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

            Console.WriteLine($"Текст: {paragraph.Range.Text}");
            Console.WriteLine($"Имя шрифта: {paragraph.Range.Font.Name}");
            Console.WriteLine($"Размер шрифта: {paragraph.Range.Font.Size}");
            Console.WriteLine($"Уровень заголовка: {paragraph.OutlineLevel}");
            Console.WriteLine($"Жирный: {paragraph.Range.Bold}");
            Console.WriteLine($"Курсив: {paragraph.Range.Italic}");
            Console.WriteLine($"Цвет текста: {paragraph.Range.Font.TextColor.RGB}");
            Console.WriteLine($"Цвет выделения: {paragraph.Range.Font.UnderlineColor}");
            Console.WriteLine($"Подчеркнутый: {paragraph.Range.Underline}");
            Console.WriteLine($"Зачеркнутый: {paragraph.Range.Font.StrikeThrough}");
            Console.WriteLine($"Надстрочность: {paragraph.Range.Font.Superscript}");
            Console.WriteLine($"Подстрочность: {paragraph.Range.Font.Subscript}");
            Console.WriteLine($"Скрытый: {paragraph.Range.Font.Hidden}");
            Console.WriteLine($"Масштаб: {paragraph.Range.Font.Scaling}");
            Console.WriteLine($"Смещение: {paragraph.Range.Font.Position}");
            Console.WriteLine($"Кернинг: {paragraph.Range.Font.Kerning}");
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

            QuitWord();
        }
    }
}