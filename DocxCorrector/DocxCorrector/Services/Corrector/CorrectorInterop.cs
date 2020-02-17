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
        public override void PrintAllParagraphs()
        {
            OpenWord();

            foreach (Word.Paragraph paragraph in Document.Paragraphs)
            {
                Console.WriteLine(paragraph.Range.Text);
            }

            QuitWord();
        }

        public override string GetMistakesJSON()
        {
            List<Mistake> mistakes = new List<Mistake> { new Mistake { ParagraphID = 0, Message = "NOT IMPLEMENTED" } };

            // TODO: - Implement

            string mistakesJSON = JSONMaker.MakeMistakesJSON(mistakes);
            return mistakesJSON;
        }
    }
}